# services/payroll_calculation_service.py
# ================================================================
# 劳务发薪核算核心服务：工时降序漏斗筛选、额度反推与储备池管理
# ================================================================

import os
import re
import json
from datetime import datetime
import pandas as pd
import numpy as np
from typing import Dict, Tuple, List, Any, Optional

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')
BACKLOG_FILE = os.path.join(DATA_DIR, 'payroll_backlog.json')


def ensure_data_dir():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR, exist_ok=True)


def number_to_rmb_capital(n: float) -> str:
    """将数字金额转换为人民币大写汉字（如：12345.67 -> 壹万贰仟叁佰肆拾伍元陆角柒分）。"""
    if n is None or pd.isna(n):
        return "零元整"
    try:
        n = round(float(n), 2)
    except (ValueError, TypeError):
        return "零元整"

    if n == 0:
        return "零元整"
    if n < 0:
        return "负" + number_to_rmb_capital(-n)

    digits = "零壹贰叁肆伍陆柒捌玖"
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "万", "亿"]

    integer_part = int(n)
    fractional_part = int(round((n - integer_part) * 100))

    int_str = str(integer_part)
    res = ""
    zero_flag = False

    # 按 4 位分段处理整数部分
    groups = []
    while int_str:
        groups.append(int_str[-4:])
        int_str = int_str[:-4]

    for gi, grp in enumerate(groups):
        grp_str = ""
        g_zero = False
        for i, ch in enumerate(reversed(grp)):
            d = int(ch)
            if d != 0:
                if g_zero:
                    grp_str = digits[0] + grp_str
                    g_zero = False
                grp_str = digits[d] + units[i] + grp_str
            else:
                if grp_str:
                    g_zero = True
        if grp_str:
            res = grp_str + big_units[gi] + res
        elif gi == 0 and not res:
            res = digits[0]

    res = res.rstrip("零") + "元" if res else "零元"

    # 处理小数部分
    jiao = fractional_part // 10
    fen = fractional_part % 10

    if jiao == 0 and fen == 0:
        res += "整"
    else:
        if jiao > 0:
            res += digits[jiao] + "角"
        elif integer_part > 0:
            res += "零"
        if fen > 0:
            res += digits[fen] + "分"

    return res


def filter_top_working_days(daily_records: Dict[Any, Any], target_days: float) -> Tuple[Dict[int, str], Dict[int, Any]]:
    """
    按每日工时降序排序，提取前 target_days 个最长工时日期。
    
    参数:
        daily_records: {日期: 工时/打卡记录}，日期可以是 1~31 或带月份日期。
        target_days: 班组长或额度核定的目标出勤天数 (支持整天或 0.5 天)。
        
    返回:
        selected: {日期: '✓' 或 '0.5'}
        rejected: {日期: 原始工时} (放入差额池，保留具体日期)
    """
    if target_days <= 0:
        return {}, {int(k): v for k, v in daily_records.items() if str(k).isdigit()}

    # 解析工时为可排序浮点数
    def _parse_hours(val):
        if val is None or val == '':
            return 0.0
        try:
            return float(val)
        except (ValueError, TypeError):
            # 非数字标记，如 '√', '✓' 等，默认为 8 小时
            s = str(val).strip()
            if '✓' in s or '√' in s or '打卡' in s or '出勤' in s:
                return 8.0
            return 0.0

    # 规范化日期键为整数
    norm_records = {}
    for k, v in daily_records.items():
        try:
            day_int = int(k)
            norm_records[day_int] = v
        except (ValueError, TypeError):
            continue

    sorted_items = sorted(
        norm_records.items(),
        key=lambda x: (_parse_hours(x[1]), -x[0]),  # 工时降序；同工时按日期升序
        reverse=True
    )

    selected = {}
    rejected = {}

    current_days = 0.0
    for day, val in sorted_items:
        hours = _parse_hours(val)
        if current_days < target_days and hours > 0:
            rem = target_days - current_days
            if rem >= 1.0:
                selected[day] = '✓'
                current_days += 1.0
            elif rem >= 0.5:
                selected[day] = '0.5'
                current_days += 0.5
            else:
                rejected[day] = val
        else:
            rejected[day] = val

    # 如果目标天数大于有效打卡天数（现场特殊特批情况），在当月其余空缺日期按顺序补齐打钩
    if current_days < target_days:
        for day in range(1, 32):
            if current_days >= target_days:
                break
            if day not in selected:
                rem = target_days - current_days
                if rem >= 1.0:
                    selected[day] = '✓'
                    current_days += 1.0
                elif rem >= 0.5:
                    selected[day] = '0.5'
                    current_days += 0.5

    return selected, rejected


def calculate_budget_allocation(
    workers_df: pd.DataFrame,
    total_budget: float,
    strategy: str = 'proportional',
    strict_half_days: bool = False
) -> pd.DataFrame:
    """
    额度比例反推与尾差平齐算法。
    
    参数:
        workers_df: 包含 ['姓名', '提取日薪', '最终核定天数']，可选 ['锁定'] (bool)。
        total_budget: 本次资金总额度 (元)。
        strategy: 策略，目前支持 'proportional' (等比缩减)。
        strict_half_days: 
            True  -> 天数严格规整为 0.5 的整数倍（通过贪心增补 0.5 天调平，总额不超过目标额度）。
            False -> 允许微小尾差直接调整至目标额度（分毫不差匹配测试与高精度场景）。
            
    返回:
        包含 ['实发天数', '实发工资', '应发工资'] 的 DataFrame。
    """
    df = workers_df.copy()

    if df.empty or total_budget <= 0:
        df['实发天数'] = 0.0
        df['实发工资'] = 0.0
        df['应发工资'] = 0.0
        return df

    if '锁定' not in df.columns:
        df['锁定'] = False

    df['提取日薪'] = pd.to_numeric(df['提取日薪'], errors='coerce').fillna(0.0) if '提取日薪' in df.columns else 0.0
    df['最终核定天数'] = pd.to_numeric(df['最终核定天数'], errors='coerce').fillna(0.0) if '最终核定天数' in df.columns else 0.0
    df['津贴'] = pd.to_numeric(df['津贴'], errors='coerce').fillna(0.0) if '津贴' in df.columns else 0.0
    df['应发工资'] = (df['最终核定天数'] * df['提取日薪'] + df['津贴']).round(2)

    locked_mask = df['锁定'] == True
    unlocked_mask = ~locked_mask

    locked_total = df.loc[locked_mask, '应发工资'].sum()
    available_budget = total_budget - locked_total

    if available_budget <= 0:
        df.loc[unlocked_mask, '实发天数'] = 0.0
        df.loc[unlocked_mask, '实发工资'] = 0.0
        df.loc[locked_mask, '实发天数'] = df.loc[locked_mask, '最终核定天数']
        df.loc[locked_mask, '实发工资'] = df.loc[locked_mask, '应发工资']
        return df

    unlocked_original_total = df.loc[unlocked_mask, '应发工资'].sum()

    if unlocked_original_total <= available_budget:
        # 额度充裕，无需缩减
        df['实发天数'] = df['最终核定天数']
        df['实发工资'] = df['应发工资']
        return df

    # 按比例缩减
    ratio = available_budget / unlocked_original_total
    df.loc[locked_mask, '实发天数'] = df.loc[locked_mask, '最终核定天数']

    # 初步向下取整到 0.5 的倍数
    df.loc[unlocked_mask, '实发天数'] = np.floor(
        df.loc[unlocked_mask, '最终核定天数'] * ratio * 2.0
    ) / 2.0
    df['实发工资'] = (df['实发天数'] * df['提取日薪'] + df['津贴']).round(2)

    current_actual_total = df['实发工资'].sum()
    diff = round(total_budget - current_actual_total, 2)

    if diff > 0 and unlocked_mask.any():
        if strict_half_days:
            # 严格以 0.5 天为步长贪心回填
            candidates = df[unlocked_mask & (df['提取日薪'] > 0)].copy()
            candidates['gap'] = candidates['最终核定天数'] - candidates['实发天数']
            candidates = candidates.sort_values(by=['gap', '提取日薪'], ascending=[False, True])

            rem_budget = diff
            for idx in candidates.index:
                rate = df.loc[idx, '提取日薪']
                half_cost = round(rate * 0.5, 2)
                if rem_budget >= half_cost and (df.loc[idx, '实发天数'] + 0.5 <= df.loc[idx, '最终核定天数']):
                    df.loc[idx, '实发天数'] += 0.5
                    rem_budget = round(rem_budget - half_cost, 2)
            df['实发工资'] = (df['实发天数'] * df['提取日薪'] + df['津贴']).round(2)
            
            # 强制吸收剩余尾差，确保把额度发够，但绝对不破坏天数 0.5 的规整性
            if rem_budget > 0 and not candidates.empty:
                target_idx = candidates.index[0]
                # 天数保持原有的 0.5 倍数不变，直接将剩余的尾差金额（元）加到该工人的实发工资上
                df.loc[target_idx, '实发工资'] = round(
                    df.loc[target_idx, '实发工资'] + rem_budget, 2
                )
        else:
            # 允许微调吸收尾差，实现总金额 100% 精确等于目标额度
            valid_indices = df[unlocked_mask & (df['提取日薪'] > 0)].index
            if not valid_indices.empty:
                target_idx = valid_indices[0]
                extra_days = diff / df.loc[target_idx, '提取日薪']
                df.loc[target_idx, '实发天数'] = round(df.loc[target_idx, '实发天数'] + extra_days, 4)
                df['实发工资'] = (df['实发天数'] * df['提取日薪'] + df['津贴']).round(2)
                df.loc[target_idx, '实发工资'] = round(
                    df.loc[target_idx, '实发工资'] + (total_budget - df['实发工资'].sum()), 2
                )

    return df


# ================================================================
# 差额蓄水池管理（第二批工资单数据底账）
# ================================================================

def load_backlog() -> Dict[str, Any]:
    """读取所有留存的差额备忘池。"""
    ensure_data_dir()
    if os.path.exists(BACKLOG_FILE):
        try:
            with open(BACKLOG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_backlog(data: Dict[str, Any]):
    """持久化差额池数据。"""
    ensure_data_dir()
    with open(BACKLOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_to_backlog(
    period: str,
    company: str,
    worker_id: str,
    worker_name: str,
    days_deducted: float,
    amount_deducted: float,
    rejected_days: dict,
    daily_rate: float = 0.0
):
    """记录某工人因额度不足扣减的工时与欠款明细。"""
    backlog = load_backlog()
    key = f"{company}_{period}"
    if key not in backlog:
        backlog[key] = {}

    worker_key = worker_id if worker_id else f"name:{worker_name}"

    if worker_key not in backlog[key]:
        backlog[key][worker_key] = {
            '姓名': worker_name,
            '身份证号': worker_id,
            '日薪': daily_rate,
            '扣减天数': 0.0,
            '欠额': 0.0,
            '扣减具体日期': {}
        }

    backlog[key][worker_key]['扣减天数'] = round(backlog[key][worker_key]['扣减天数'] + days_deducted, 2)
    backlog[key][worker_key]['欠额'] = round(backlog[key][worker_key]['欠额'] + amount_deducted, 2)
    if daily_rate > 0:
        backlog[key][worker_key]['日薪'] = daily_rate
    backlog[key][worker_key]['扣减具体日期'].update(rejected_days)

    save_backlog(backlog)


def get_backlog_for_company(company: str) -> Dict[str, Any]:
    """查询指定分包公司所有的待发差额批次。"""
    backlog = load_backlog()
    res = {}
    for k, v in backlog.items():
        if k.startswith(company + "_"):
            period = k.split("_", 1)[1]
            res[period] = v
    return res


def extract_from_backlog(company: str, period: str, worker_keys: List[str]) -> List[Dict[str, Any]]:
    """从差额池提取并移除指定人员的待发记录，用于生成第二批工资单。"""
    backlog = load_backlog()
    key = f"{company}_{period}"
    if key not in backlog:
        return []

    extracted = []
    for wk in worker_keys:
        if wk in backlog[key]:
            record = backlog[key].pop(wk)
            record['_worker_key'] = wk
            record['period'] = period
            record['company'] = company
            extracted.append(record)

    if not backlog[key]:
        del backlog[key]

    save_backlog(backlog)
    return extracted


# ================================================================
# 总包系统打卡流水解析与三方智能裁决引擎
# ================================================================

def parse_sanju_punch_excel(file_input) -> Tuple[pd.DataFrame, str]:
    """
    解析中建三局劳务工人考勤明细表（带每日打卡时间及工时）。
    返回:
        (workers_df, period_str)
    """
    if hasattr(file_input, 'seek'):
        file_input.seek(0)

    head_raw = pd.read_excel(file_input, header=None, nrows=6)
    period_str = '8月'
    for _, row in head_raw.iterrows():
        row_text = ' '.join(str(v) for v in row.values)
        m = re.search(r'(\d{4}年\d{1,2}月|\d{1,2}月)', row_text)
        if m:
            period_str = m.group(1)
            break

    header_idx = 2
    for i, row in head_raw.iterrows():
        cells = [str(v) for v in row.values]
        if any('工人姓名' in c or '姓名' in c for c in cells) and any('考勤日期' in c or '工作时长' in c for c in cells):
            header_idx = i
            break

    if hasattr(file_input, 'seek'):
        file_input.seek(0)
    df = pd.read_excel(file_input, header=header_idx)

    col_map = {
        '工人姓名': '姓名',
        '身份证号': '身份证号',
        '分包企业': '分包/所属企业',
        '班组': '班组',
        '工种': '工种',
        '考勤日期': '考勤日期',
        '工作时长(小时)': '工作时长',
        '项目名称': '项目名称',
    }
    for orig, target in col_map.items():
        if orig in df.columns:
            df[target] = df[orig]

    df['工作时长'] = pd.to_numeric(df.get('工作时长', 0), errors='coerce').fillna(0.0)

    def _extract_day(d):
        if isinstance(d, (pd.Timestamp, datetime)):
            return d.day
        s = str(d).strip()
        try:
            return pd.to_datetime(s).day
        except Exception:
            return None

    df['day'] = df['考勤日期'].apply(_extract_day)
    df = df.dropna(subset=['day'])
    df['day'] = df['day'].astype(int)

    group_cols = ['姓名', '身份证号']
    for c in ['分包/所属企业', '工种', '班组']:
        if c in df.columns:
            group_cols.append(c)

    workers = []
    for keys, grp in df.groupby(group_cols):
        row_meta = grp.iloc[0]
        daily_records = {}
        for _, row in grp.iterrows():
            d = row['day']
            h = float(row['工作时长'])
            daily_records[d] = round(h, 2)

        # 统计有效打卡天数（大于0工时）
        punch_days = sum(1 for h in daily_records.values() if h > 0.0)
        total_hours = sum(daily_records.values())
        avg_hours = total_hours / punch_days if punch_days > 0 else 0.0

        workers.append({
            '姓名': str(row_meta.get('姓名', '')).strip(),
            '身份证号': str(row_meta.get('身份证号', '')).strip(),
            '分包/所属企业': str(row_meta.get('分包/所属企业', '')).strip(),
            '工种': str(row_meta.get('工种', '普工')).strip(),
            '班组': str(row_meta.get('班组', '')).strip(),
            '总包系统天数': float(punch_days),
            '总工时': round(total_hours, 2),
            '平均工时': round(avg_hours, 2),
            '_raw_punch_hours': dict(daily_records),
            '__考勤明细__': daily_records,
            '月份': period_str
        })

    res_df = pd.DataFrame(workers)
    return res_df, period_str


def reconcile_attendance(
    system_df: pd.DataFrame,
    my_df: Optional[pd.DataFrame] = None,
    foreman_days_map: Optional[Dict[str, float]] = None,
    final_days_map: Optional[Dict[str, float]] = None,
    allowance_map: Optional[Dict[str, float]] = None
) -> pd.DataFrame:
    """
    三方考勤智能对账与裁决引擎：
    基准铁律：系统考勤为核查基线，默认最终有效工日 = min(班组长核定工日, 总包系统打卡天数)。
    实战特批：支持劳资员人工调整最终裁定工日（不锁死，应对断电故障与特批补卡）。
    当核减时，依据每日工时降序筛选打 ✓；当特批超额时，自动在空缺日补齐打钩。
    """
    df = system_df.copy()
    if '_raw_punch_hours' not in df.columns:
        df['_raw_punch_hours'] = [
            dict(r.get('__考勤明细__', {})) if isinstance(r.get('__考勤明细__'), dict) else {}
            for _, r in df.iterrows()
        ]

    # 1. 匹配【我的考勤】
    my_map = {}
    if my_df is not None and not my_df.empty:
        id_col = next((c for c in my_df.columns if any(k in c for k in ['身份证', '证件号'])), None)
        name_col = next((c for c in my_df.columns if any(k in c for k in ['姓名', '人员'])), None)

        # 优先精确查找代表天数的列，严格避开 '__考勤明细__' 等字典列
        day_cands = [
            '最终核定天数', '解析出勤天数', '出勤天数', '实际出勤天数',
            '考勤天数', '实际天数', '核定天数', '天数', '出勤工日', '工日'
        ]
        days_col = next((c for c in day_cands if c in my_df.columns), None)
        if not days_col:
            days_col = next(
                (c for c in my_df.columns if any(k in c for k in ['天数', '工日', '出勤']) and not str(c).startswith('__') and '明细' not in str(c)),
                None
            )

        if days_col and (id_col or name_col):
            for _, r in my_df.iterrows():
                k_id = str(r.get(id_col, '')).strip() if id_col else ''
                k_name = str(r.get(name_col, '')).strip() if name_col else ''
                raw_val = r.get(days_col, 0)

                # 防御性解析：如果单元格是字典（如考勤明细），统计有效出勤天数；否则转为浮点数
                if isinstance(raw_val, dict):
                    val = float(sum(1 for v in raw_val.values() if v and str(v).strip() not in ['', '0', '0.0', '休', '-']))
                else:
                    try:
                        val = float(raw_val)
                    except (ValueError, TypeError):
                        val = np.nan

                if pd.notna(val):
                    if k_id and len(k_id) >= 15:
                        my_map[k_id] = float(val)
                    if k_name:
                        my_map[f"name:{k_name}"] = float(val)

    def _get_my_days(row):
        id_card = str(row.get('身份证号', '')).strip()
        name = str(row.get('姓名', '')).strip()
        if id_card in my_map:
            return my_map[id_card]
        if f"name:{name}" in my_map:
            return my_map[f"name:{name}"]
        return float(row.get('总包系统天数', 0.0))

    df['我的考勤天数'] = df.apply(_get_my_days, axis=1)

    # 2. 匹配【班组长核定工日】
    f_map = foreman_days_map or {}
    def _get_foreman_days(row):
        id_card = str(row.get('身份证号', '')).strip()
        name = str(row.get('姓名', '')).strip()
        if id_card in f_map:
            return float(f_map[id_card])
        if f"name:{name}" in f_map:
            return float(f_map[f"name:{name}"])
        if name in f_map:
            return float(f_map[name])
        return float(row['我的考勤天数'])

    df['班组长核定工日'] = df.apply(_get_foreman_days, axis=1)

    # 3. 匹配【最终裁定工日】：支持现场特殊特批修改，不写死
    fin_map = final_days_map or {}
    def _get_final_days(row):
        id_card = str(row.get('身份证号', '')).strip()
        name = str(row.get('姓名', '')).strip()
        # 优先使用人工调整的裁定工日
        if id_card in fin_map:
            return float(fin_map[id_card])
        if f"name:{name}" in fin_map:
            return float(fin_map[f"name:{name}"])
        if name in fin_map:
            return float(fin_map[name])
        # 默认基准裁决：min(班组长核定工日, 总包系统天数)
        return float(min(row['班组长核定工日'], row['总包系统天数']))

    df['最终裁定工日'] = df.apply(_get_final_days, axis=1)

    # 4. 匹配【津贴】
    alw_map = allowance_map or {}
    def _get_allowance(row):
        id_card = str(row.get('身份证号', '')).strip()
        name = str(row.get('姓名', '')).strip()
        if id_card in alw_map:
            return float(alw_map[id_card])
        if f"name:{name}" in alw_map:
            return float(alw_map[f"name:{name}"])
        if name in alw_map:
            return float(alw_map[name])
        return float(row.get('津贴', 0.0) or 0.0)

    df['津贴'] = df.apply(_get_allowance, axis=1).round(2)

    # 5. 状态标记与工时降序筛选打钩
    status_list = []
    selected_details = []
    rejected_details = []

    for _, row in df.iterrows():
        sys_d = float(row['总包系统天数'])
        f_d = float(row['班组长核定工日'])
        fin_d = float(row['最终裁定工日'])
        raw_hours = row.get('_raw_punch_hours')
        if not isinstance(raw_hours, dict) or not raw_hours:
            raw_hours = row.get('__考勤明细__', {})
            if not isinstance(raw_hours, dict):
                raw_hours = {}

        if fin_d <= 0:
            status = "⚪ 0考勤 (不纳入考勤表)"
        elif fin_d > sys_d:
            diff = fin_d - sys_d
            status = f"⚡ 特殊特批工日 (超出系统{diff:.1f}天)"
        elif fin_d < sys_d:
            diff = sys_d - fin_d
            status = f"✂️ 班组长核减{diff:.1f}天 (剔除短工时)"
        elif f_d > sys_d:
            status = f"⚠️ 受限系统打卡 (系统仅{sys_d:.1f}天)"
        else:
            status = "✔️ 完全一致"
        status_list.append(status)

        # 依据真实工时由长到短排布打 ✓，若特批超额则补齐打钩
        sel, rej = filter_top_working_days(raw_hours, fin_d)
        selected_details.append(sel)
        rejected_details.append(rej)

    df['裁定状态'] = status_list
    df['__考勤明细__'] = selected_details
    df['最终核定天数'] = df['最终裁定工日']
    df['解析出勤天数'] = df['总包系统天数']
    df['_rejected_foreman'] = rejected_details
    if '提取日薪' in df.columns:
        df['提取日薪'] = pd.to_numeric(df['提取日薪'], errors='coerce').fillna(300.0)
        df['应发工资'] = (df['最终核定天数'] * df['提取日薪'] + df['津贴']).round(2)

    return df

