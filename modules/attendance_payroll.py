# modules/attendance_payroll.py
# ============================================================
# 考勤对账、定稿与工资联动结算模块  v2
# 架构: 三层解耦 — 数据层 / 导出层 / 展示层
# CSS: 严格复用 assets/style.css，禁止内联硬编码样式
#
# Tab 1 流程门控设计:
#   Step A -> 三方差异对比报告（只读，供领导确认）
#   Step B -> 定稿录入（领导确认后，录入最终核定天数）
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
from copy import copy
from pathlib import Path
from datetime import datetime
from modules.master_data import load_master_df

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# ================================================================
# 数据层
# ================================================================

# ── 列名候选表（精确匹配单元格值）─────────────────────────────
_NAME_CANDS = ['姓名', '名', '名字', '人员姓名', 'Name']
_ID_CANDS = ['身份证号', '身份证号码', '身份证', '证件号码', '证件号']
_COMPANY_CANDS = ['分包/所属企业', '所属企业', '分包单位', '单位', '劳务单位']
_DAYS_CANDS = [
    '考勤天数', '出勤天数', '实际天数', '天数', '实出勤', '实际出勤天数',
    '当月出勤天数', '小计', '合计', '合计天数', '总计', '天', '出勤合计', '出勤', 'Days',
]
_SERIAL_CANDS = ['序号', '编号', '序', 'No.', 'NO.']
_NAME_EXCLUDED_HINTS = (
    '签字', '签名', '申明', '声明', '项目', '工程', '班组', '负责人',
    '名称', '编号', '序号', '日期', '备注', '单位', '工资', '考勤', '人员',
)


def _normalize_header(value):
    """统一处理 Excel 表头中的换行、普通空格和全角空格。"""
    return re.sub(r'[\s\u3000]+', '', str(value).strip())


def _clean_identity(value):
    """清理身份证号中的空格和 Excel 文本化产生的尾部 .0。"""
    text = _raw_attendance_value(value).replace(' ', '').upper()
    return text[:-2] if text.endswith('.0') else text


def _is_valid_identity(value):
    return len(_clean_identity(value)) >= 15


def _looks_like_person_name(value):
    """过滤模板里的声明、签字、项目名称等非人员文本。"""
    name = _raw_attendance_value(value)
    compact = _normalize_header(name)
    if not compact or any(hint in compact for hint in _NAME_EXCLUDED_HINTS):
        return False
    return bool(re.fullmatch(r'[\u4e00-\u9fff·]{2,12}', compact))


def _detect_header_row(file_obj) -> int:
    """
    扫描前 12 行，找到包含姓名字段的真实表头行。
    兼容“日期\n姓名”“姓名（签字）”等组合表头。
    """
    try:
        raw = pd.read_excel(file_obj, dtype=str, header=None, nrows=12)
        for i in range(min(12, len(raw))):
            row_cells = [_normalize_header(v) for v in raw.iloc[i].tolist()]
            if any(
                ('姓名' in cell or '人员姓名' in cell or '名字' in cell or cell == '名' or cell.lower() == 'name')
                for cell in row_cells
            ):
                return i
    except Exception:
        pass
    return 0


def _safe_read(file_obj, label: str) -> pd.DataFrame:
    """
    智能读取考勤 Excel：
    1. 自动检测真实表头行（精确匹配姓名候选词）
    2. 统一字符串类型、清洗列名
    3. 过滤空行和纯数字序号/分隔行
    """
    if file_obj is None:
        return pd.DataFrame()
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        header_row = _detect_header_row(file_obj)

        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        df = pd.read_excel(file_obj, dtype=str, header=header_row)
        df.columns = [re.sub(r'[\s\u3000]+', '', str(c).strip()) for c in df.columns]

        name_col = _resolve_col(df, _NAME_CANDS)
        if name_col:
            id_col = _resolve_col(df, _ID_CANDS)
            serial_col = _resolve_col(df, _SERIAL_CANDS)

            def _is_person_row(row):
                name = row.get(name_col, '')
                if _is_valid_identity(row.get(id_col, '')) if id_col else False:
                    return True
                serial = str(row.get(serial_col, '')).strip() if serial_col else ''
                return _looks_like_person_name(name) and (bool(re.fullmatch(r'\d+(?:\.0)?', serial)) or not serial_col)

            mask = (
                df[name_col].notna() &
                (df[name_col].astype(str).str.strip() != '') &
                df.apply(_is_person_row, axis=1)
            )
            df = df[mask].reset_index(drop=True)

        df['__来源__'] = label
        id_col = _resolve_col(df, _ID_CANDS)
        if id_col and id_col != '身份证号':
            df['身份证号'] = df[id_col].apply(_clean_identity)
        elif id_col:
            df['身份证号'] = df[id_col].apply(_clean_identity)
        company_col = _resolve_col(df, _COMPANY_CANDS)
        if company_col and company_col != '分包/所属企业':
            df['分包/所属企业'] = df[company_col].astype(str).str.strip()
        return df
    except Exception as exc:
        st.error(f'读取【{label}】文件失败：{exc}')
        return pd.DataFrame()


def _resolve_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    normalized_columns = {
        column: _normalize_header(column)
        for column in df.columns
    }
    # 兼容组合表头，例如“日期姓名”“姓名（签字）”；单字“名”只允许精确匹配，
    # 防止把“项目名称”“班组名称”误认为姓名列。
    for candidate in candidates:
        normalized_candidate = _normalize_header(candidate)
        for column, normalized_column in normalized_columns.items():
            if normalized_candidate == '名':
                if normalized_column == '名':
                    return column
            elif normalized_candidate and normalized_candidate in normalized_column:
                return column
    return None


def _day_number(column):
    """识别考勤表中的日期列：1、01、1日、01日等均视为同一天。"""
    text = str(column).strip().replace("号", "").replace("日", "")
    if text.isdigit():
        day = int(text)
        if 1 <= day <= 31:
            return day
    return None


def _raw_attendance_value(value):
    """保留官网/纸质表原始内容，不预设最终输出为√或时间段。"""
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null", "<na>"}:
        return ""
    return text


def _attendance_state(value):
    """仅做比对状态识别，不决定最终写回工资表的格式。"""
    text = _raw_attendance_value(value)
    if not text or text in {"--", "—", "-", "/"}:
        return "缺勤/无记录"
    if any(mark in text for mark in ["√", "✓", "✔", "出勤", "正常"]):
        return "有考勤"
    if re.search(r"\d{1,2}:\d{2}", text):
        return "有考勤"
    return "待确认"


def _extract_daily_source(df, source_name):
    if df is None or df.empty:
        return {}
    name_col = _resolve_col(df, _NAME_CANDS)
    id_col = _resolve_col(df, _ID_CANDS)
    if name_col is None:
        return {}

    day_cols = {day: col for col in df.columns if (day := _day_number(col)) is not None}
    records = {}
    for _, row in df.iterrows():
        name = _raw_attendance_value(row.get(name_col, ""))
        identity = _raw_attendance_value(row.get(id_col, "")) if id_col else ""
        key = identity if len(identity) >= 15 else f"姓名:{name}"
        if not name or not key:
            continue
        item = records.setdefault(
            key,
            {"姓名": name, "身份证号": identity, "来源": source_name, "days": {}},
        )
        if not item["身份证号"] and identity:
            item["身份证号"] = identity
        for day, col in day_cols.items():
            value = _raw_attendance_value(row.get(col, ""))
            if value:
                item["days"][day] = value
    return records


def build_daily_diff_table(df_a, df_b, df_paper, company=""):
    """按人员+日期保留三份考勤原始单元格，供人工确认差异。

    这里故意不生成“最终√”或“最终时间段”，只保存原始值和识别状态，
    等领导确认后再决定最终输出格式。
    """
    sources = [
        ("三局", _extract_daily_source(df_a, "三局")),
        ("护薪", _extract_daily_source(df_b, "护薪")),
        ("纸质", _extract_daily_source(df_paper, "纸质")),
    ]
    all_keys = set()
    for _, records in sources:
        all_keys.update(records.keys())

    rows = []
    for key in sorted(all_keys):
        records = {name: data.get(key, {}) for name, data in sources}
        base = next((data for data in records.values() if data), {})
        for day in range(1, 32):
            values = {name: data.get("days", {}).get(day, "") for name, data in records.items()}
            states = [_attendance_state(values[name]) for name in ["三局", "护薪", "纸质"]]
            if all(state == "缺勤/无记录" for state in states):
                status = "— 数据缺失"
            elif len(set(states)) == 1 and "待确认" not in states:
                status = "✔ 一致"
            else:
                status = "⚠ 有差异"
            rows.append({
                "分包/所属企业": company,
                "姓名": base.get("姓名", ""),
                "身份证号": base.get("身份证号", ""),
                "日期": day,
                "三局原始内容": values["三局"],
                "护薪原始内容": values["护薪"],
                "纸质原始内容": values["纸质"],
                "三局识别": states[0],
                "护薪识别": states[1],
                "纸质识别": states[2],
                "日差异状态": status,
            })
    return pd.DataFrame(rows)


def _to_float(v):
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return np.nan


def build_diff_table(df_a, df_b, df_paper, company=''):

    def _extract(df, alias):
        if df.empty:
            return pd.DataFrame(columns=['姓名', '身份证号', alias, '__key__'])
        name_col = _resolve_col(df, _NAME_CANDS)
        id_col = _resolve_col(df, _ID_CANDS)
        days_col = _resolve_col(df, _DAYS_CANDS)
        if name_col is None:
            return pd.DataFrame(columns=['姓名', '身份证号', alias, '__key__'])
        out = pd.DataFrame()
        out['姓名'] = df[name_col].astype(str).str.strip()
        out['身份证号'] = df[id_col].apply(_clean_identity) if id_col else ''
        out[alias] = df[days_col].apply(_to_float) if days_col else np.nan
        out = out[
            out.apply(
                lambda row: _is_valid_identity(row['身份证号']) or _looks_like_person_name(row['姓名']),
                axis=1,
            )
        ]
        return out.reset_index(drop=True)

    ea = _extract(df_a, '三局天数')
    eb = _extract(df_b, '护薪天数')
    ep = _extract(df_paper, '纸质天数')

    # 先收集三份考勤中“姓名 -> 身份证号”的唯一映射；再补充主数据中的同公司唯一姓名。
    # 这使没有身份证号的 6 月纸质模板仍能合并到三局/护薪/主数据里的同一人员。
    name_to_ids = {}

    def _add_name_id(name, identity):
        clean_name = _raw_attendance_value(name)
        clean_id = _clean_identity(identity)
        if clean_name and _is_valid_identity(clean_id):
            name_to_ids.setdefault(clean_name, set()).add(clean_id)

    for part in [ea, eb, ep]:
        for _, row in part.iterrows():
            _add_name_id(row.get('姓名', ''), row.get('身份证号', ''))

    master = _load_master_data()
    if master is not None and not master.empty and {'姓名', '身份证号'}.issubset(master.columns):
        master_scope = master.copy()
        if company and '分包/所属企业' in master_scope.columns:
            scoped = master_scope[master_scope['分包/所属企业'].astype(str).str.strip() == str(company).strip()]
            if not scoped.empty:
                master_scope = scoped
        for _, row in master_scope.iterrows():
            _add_name_id(row.get('姓名', ''), row.get('身份证号', ''))

    for part in [ea, eb, ep]:
        if part.empty:
            continue
        for idx, row in part.iterrows():
            identity = _clean_identity(row.get('身份证号', ''))
            if not _is_valid_identity(identity):
                candidates = name_to_ids.get(_raw_attendance_value(row.get('姓名', '')), set())
                if len(candidates) == 1:
                    identity = next(iter(candidates))
                    part.at[idx, '身份证号'] = identity
            part.at[idx, '__key__'] = identity if _is_valid_identity(identity) else f"姓名:{row.get('姓名', '')}"
        part.drop_duplicates(subset='__key__', keep='first', inplace=True)
        part.reset_index(drop=True, inplace=True)

    merged = ea.copy()
    for part in [eb, ep]:
        if part.empty or '__key__' not in part.columns:
            continue
        merged = merged.merge(part, on='__key__', how='outer', suffixes=('', '_new'))
        for col in ['姓名', '身份证号']:
            new_col = f'{col}_new'
            if new_col in merged.columns:
                merged[col] = merged[col].replace('', pd.NA).fillna(merged[new_col])
                merged.drop(columns=[new_col], inplace=True)

    for col in ['三局天数', '护薪天数', '纸质天数']:
        if col not in merged.columns:
            merged[col] = np.nan

    def _status(row):
        vals = [row['三局天数'], row['护薪天数'], row['纸质天数']]
        valid = [v for v in vals if not (isinstance(v, float) and np.isnan(v))]
        if len(valid) == 0:
            return '— 数据缺失'
        if len(set(valid)) == 1:
            return '✔ 一致'
        return '⚠ 有差异'

    merged['差异状态'] = merged.apply(_status, axis=1)
    merged = merged.reset_index(drop=True)
    result = merged[['姓名', '身份证号', '三局天数', '护薪天数', '纸质天数', '差异状态']]
    if company:
        result.insert(0, '分包/所属企业', company)
    return result


# ================================================================
# 导出层
# ================================================================

def _thin_border():
    s = Side(style='thin', color='D1D5DB')
    return Border(top=s, left=s, right=s, bottom=s)


def _cell_style(cell, font=None, fill=None, align=None, border=None):
    if font:   cell.font      = font
    if fill:   cell.fill      = fill
    if align:  cell.alignment = align
    if border: cell.border    = border


def _add_title_row(ws, title, n_cols, fill):
    ws.insert_rows(1)
    ws.row_dimensions[1].height = 36
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(bold=True, size=14, color='FFFFFF', name='微软雅黑')
    c.fill      = fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border    = _thin_border()
    if n_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)


def _set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _ensure_data_rows(ws, start_row, count, footer_row):
    """在模板签字/合计区域前扩展数据行，并复制上一行的样式。"""
    available = max(0, footer_row - start_row)
    extra = count - available
    if extra <= 0:
        return

    # openpyxl 的 insert_rows 不会自动移动 merged_cells；先拆分底部签字区，
    # 插入数据行后再按新行号恢复合并，否则新增数据会写入 MergedCell。
    footer_merges = []
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= footer_row:
            footer_merges.append((merged.min_row, merged.min_col, merged.max_row, merged.max_col))
            ws.unmerge_cells(str(merged))

    ws.insert_rows(footer_row, extra)
    source_row = footer_row - 1
    for row_idx in range(footer_row, footer_row + extra):
        ws.row_dimensions[row_idx].height = ws.row_dimensions[source_row].height
        for col_idx in range(1, ws.max_column + 1):
            source = ws.cell(source_row, col_idx)
            target = ws.cell(row_idx, col_idx)
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.alignment:
                target.alignment = copy(source.alignment)

    for min_row, min_col, max_row, max_col in footer_merges:
        ws.merge_cells(
            start_row=min_row + extra,
            start_column=min_col,
            end_row=max_row + extra,
            end_column=max_col,
        )


def _find_footer_row(ws, start_row):
    """识别模板底部声明/签字区域的起始行，兼容旧版和新版工资模板。"""
    keywords = ('申明', '声明', '班组长签字', '第  1  页', '第1页')
    for row_idx in range(start_row, ws.max_row + 1):
        values = [ws.cell(row_idx, col_idx).value for col_idx in range(1, min(ws.max_column, 5) + 1)]
        text = ' '.join(str(value) for value in values if value is not None)
        if any(keyword in text for keyword in keywords):
            return row_idx
    return ws.max_row + 1


def _select_company_sheet(wb, company):
    """按公司简称选择模板 Sheet；没有匹配时才回退到 active。"""
    aliases = []
    if '江苏旭之升' in str(company):
        aliases = ['旭之升', '江苏']
    elif '青海久昌' in str(company):
        aliases = ['青海久昌', '久昌']
    for ws in wb.worksheets:
        if any(alias in ws.title for alias in aliases):
            return ws
    return wb.active


def _keep_only_sheet(wb, ws):
    """公司级导出只保留当前公司的 Sheet，避免打开文件时落到另一家公司空白页。"""
    for other in list(wb.worksheets):
        if other is not ws:
            wb.remove(other)
    wb.active = 0
    return ws


def _set_wage_sheet_header(ws, comp_df, company):
    """更新新版公司工资模板的合并表头，保留模板原有月份文本。"""
    project = ''
    for col in ['项目全称', '项目简称']:
        if col in comp_df.columns:
            values = comp_df[col].dropna().astype(str).str.strip()
            values = values[values.ne('')]
            if not values.empty:
                project = values.iloc[0]
                break
    if not project:
        project = '科技文化中心—国际体育中心（足球场项目）'

    team = ''
    if '班组' in comp_df.columns:
        teams = comp_df['班组'].dropna().astype(str).str.strip()
        teams = teams[teams.ne('')].drop_duplicates().tolist()
        team = '、'.join(teams[:3]) if teams else '各班组'

    original = str(ws['A2'].value or '')
    month_match = re.search(r'\d{4}\s*年[^\n]*?月', original)
    month_text = month_match.group(0) if month_match else '年  月'
    ws['A2'] = f'项目名称（全称）：{project}             班组名称：{team}                    {month_text}'


def build_diff_report_xlsx(diff_df):
    if not OPENPYXL_OK:
        return b''
    wb = Workbook()
    ws = wb.active
    ws.title = '三方考勤差异报告'
    ws.freeze_panes = 'A3'

    fill_title  = PatternFill('solid', fgColor='7C3AED')
    fill_header = PatternFill('solid', fgColor='5B21B6')
    fill_ok     = PatternFill('solid', fgColor='D1FAE5')
    fill_warn   = PatternFill('solid', fgColor='FEF3C7')
    fill_miss   = PatternFill('solid', fgColor='F1F5F9')
    fh  = Font(bold=True, color='FFFFFF', size=11, name='微软雅黑')
    fb  = Font(size=10, color='334155', name='微软雅黑')
    fw  = Font(size=10, color='92400E', name='微软雅黑', bold=True)
    fm  = Font(size=10, color='94A3B8', name='微软雅黑', italic=True)
    ac  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bd  = _thin_border()
    today = datetime.now().strftime('%Y年%m月')

    headers = ['序号', '公司', '姓名', '身份证号', '三局系统天数', '护薪系统天数', '纸质天数', '差异状态', '领导批示（手填）']
    ws.row_dimensions[1].height = 30
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        _cell_style(c, font=fh, fill=fill_header, align=ac, border=bd)

    for i, row in diff_df.iterrows():
        r = i + 2
        ws.row_dimensions[r].height = 24
        status = str(row.get('差异状态', ''))
        if status == '✔ 一致':
            row_fill, row_font = fill_ok, fb
        elif status.startswith('⚠'):
            row_fill, row_font = fill_warn, fw
        else:
            row_fill, row_font = fill_miss, fm

        def _clean(k):
            v = row.get(k)
            return '' if (v is None or (isinstance(v, float) and np.isnan(v))) else v

        vals = [
            i + 1,
            row.get('分包/所属企业', ''),
            row.get('姓名', ''),
            _clean('身份证号'),
            _clean('三局天数'),
            _clean('护薪天数'),
            _clean('纸质天数'),
            status,
            '',
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            _cell_style(c, font=row_font, fill=row_fill, align=ac, border=bd)

    _set_col_widths(ws, [6, 22, 12, 20, 11, 11, 11, 14, 24])
    _add_title_row(ws, f'三方考勤差异核查报告 — {today}（仅供领导确认，请勿直接发放）', len(headers), fill_title)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _load_master_data():
    try:
        df = load_master_df()
        if df.empty:
            return None
        for col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()
        
        def extract_wage(val):
            if pd.isna(val): return None
            match = re.search(r'(\d+)', str(val))
            return float(match.group(1)) if match else None

        if '结算单价/标准' in df.columns:
            df['提取日薪'] = df['结算单价/标准'].apply(extract_wage)
        else:
            df['提取日薪'] = None
            
        if '身份证号' in df.columns:
            df['身份证号'] = df['身份证号'].astype(str).str.strip()
            df = df.sort_values('身份证号').drop_duplicates(subset=['身份证号'], keep='last')
        else:
            df = df.drop_duplicates(subset=['姓名'])
        return df
    except Exception:
        return None


def enrich_with_master(final_df):
    """按身份证号补齐主数据；身份证号缺失时仅对唯一姓名做安全兜底。"""
    master = _load_master_data()
    if final_df is None or final_df.empty or master is None or master.empty:
        return final_df.copy() if isinstance(final_df, pd.DataFrame) else pd.DataFrame()

    out = final_df.copy()
    for col in ['身份证号', '姓名']:
        if col in out.columns:
            out[col] = out[col].fillna('').astype(str).str.strip()
    master = master.copy()
    master['身份证号'] = master.get('身份证号', '').fillna('').astype(str).str.strip()
    master['姓名'] = master.get('姓名', '').fillna('').astype(str).str.strip()

    master_cols = [c for c in [
        '姓名', '性别', '身份证号', '手机号', '班组', '工种', '人员类型',
        '分包/所属企业', '工资卡号', '开户银行', '提取日薪', '项目全称'
    ] if c in master.columns]
    lookup = master[master_cols].copy()
    lookup = lookup.rename(columns={c: f'主表_{c}' for c in master_cols if c not in ['身份证号']})
    out = out.merge(lookup, on='身份证号', how='left')

    # 主表是最终权威来源；考勤文件只用于三方天数和异常核对。
    for col in ['姓名', '性别', '手机号', '班组', '工种', '分包/所属企业', '工资卡号', '开户银行', '提取日薪', '项目全称']:
        master_col = f'主表_{col}'
        if master_col in out.columns:
            if col not in out.columns:
                out[col] = ''
            out[col] = out[col].replace('', pd.NA).fillna(out[master_col])
            out.drop(columns=[master_col], inplace=True)

    # 只对身份证号确实无法匹配、且主表姓名唯一的人员进行姓名兜底。
    if '姓名' in out.columns and '姓名' in master.columns:
        unique_name = master[master['姓名'].ne('')].groupby('姓名').filter(lambda x: len(x) == 1)
        name_map = unique_name.set_index('姓名').to_dict('index')
        for idx, row in out.iterrows():
            if row.get('身份证号', '') in master['身份证号'].values:
                continue
            candidate = name_map.get(str(row.get('姓名', '')).strip())
            if candidate:
                for col in ['性别', '手机号', '班组', '工种', '分包/所属企业', '工资卡号', '开户银行', '提取日薪', '项目全称']:
                    if (not str(row.get(col, '')).strip()) and candidate.get(col):
                        out.at[idx, col] = candidate[col]

    # 统一导出字段别名，避免模板导出阶段因界面字段名不同而漏填电话/卡号。
    if '联系电话' not in out.columns:
        out['联系电话'] = ''
    if '手机号' in out.columns:
        out['联系电话'] = out['联系电话'].replace('', pd.NA).fillna(out['手机号'])
    if '银行卡号' not in out.columns:
        out['银行卡号'] = ''
    if '工资卡号' in out.columns:
        out['银行卡号'] = out['银行卡号'].replace('', pd.NA).fillna(out['工资卡号'])

    return out


import zipfile
from openpyxl import load_workbook

def build_all_exports_zip(salary_df):
    if not OPENPYXL_OK:
        return b''
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    zip_buf = io.BytesIO()
    
    def _value(row, *names):
        for name in names:
            value = row.get(name, '')
            if value is None:
                continue
            try:
                if pd.isna(value):
                    continue
            except (TypeError, ValueError):
                pass
            return value
        return ''

    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        companies = salary_df.get('分包/所属企业', pd.Series(dtype=str)).fillna('').astype(str).str.strip().unique()
        companies = [c for c in companies if c] or ['未知公司']

        for comp in companies:
            comp_df = salary_df[salary_df.get('分包/所属企业', '') == comp].copy() if '分包/所属企业' in salary_df.columns else salary_df.copy()
            comp_df = comp_df.reset_index(drop=True)
            
            # --- 1. 公司标准_考勤表 ---
            try:
                wb = load_workbook('load-data/muban/gongsi/附件7：务工人员（含队长、班组长、弄民工）考勤表.xlsx')
                ws = _select_company_sheet(wb, comp)
                ws = _keep_only_sheet(wb, ws)
                start_r = 7
                for i, (_, row) in enumerate(comp_df.iterrows(), start=start_r):
                    ws.cell(row=i, column=2, value=_value(row, '姓名'))
                    ws.cell(row=i, column=3, value=_value(row, '工种'))
                    ws.cell(row=i, column=35, value=_value(row, '最终核定天数'))
                buf = io.BytesIO()
                wb.save(buf)
                zf.writestr(f'{comp}/公司标准_考勤表_{timestamp}.xlsx', buf.getvalue())
            except Exception as e:
                st.error(f'公司标准考勤表导出失败（{comp}）：{e}')
                
            # --- 2. 公司标准_工资表 ---
            try:
                wb = load_workbook('load-data/muban/gongsi/副本工资发放表.xlsx')
                ws = _select_company_sheet(wb, comp)
                ws = _keep_only_sheet(wb, ws)
                start_r = 5
                footer_row = _find_footer_row(ws, start_r)
                _ensure_data_rows(ws, start_r, len(comp_df), footer_row=footer_row)
                _set_wage_sheet_header(ws, comp_df, comp)
                for i, (_, row) in enumerate(comp_df.iterrows(), start=start_r):
                    ws.cell(row=i, column=1, value=i - start_r + 1)
                    ws.cell(row=i, column=2, value=_value(row, '姓名'))
                    ws.cell(row=i, column=3, value=_value(row, '工种'))
                    ws.cell(row=i, column=4, value=_value(row, '最终核定天数'))
                    ws.cell(row=i, column=5, value=_value(row, '日薪', '提取日薪'))
                    ws.cell(row=i, column=6, value=_value(row, '应发工资'))
                    ws.cell(row=i, column=11, value=_value(row, '应发工资'))
                    # L/M 为支付确认签字和领款签字，保持空白；N 为新版模板的银行卡号列。
                    ws.cell(row=i, column=14, value=_value(row, '银行卡号', '工资卡号', '卡号'))
                buf = io.BytesIO()
                wb.save(buf)
                zf.writestr(f'{comp}/公司标准_工资确认表_{timestamp}.xlsx', buf.getvalue())
            except Exception as e:
                st.error(f'公司标准工资表导出失败（{comp}）：{e}')
                
            # --- 3. 总包标准_考勤表 ---
            try:
                wb = load_workbook('load-data/muban/zongbao/考勤表（一式两份本人签字摁手印）.xlsx')
                ws = wb.active
                start_r = 5
                for i, (_, row) in enumerate(comp_df.iterrows(), start=start_r):
                    ws.cell(row=i, column=1, value=i-start_r+1)
                    ws.cell(row=i, column=2, value=_value(row, '姓名'))
                    ws.cell(row=i, column=3, value=_value(row, '性别'))
                    ws.cell(row=i, column=35, value=_value(row, '最终核定天数'))
                buf = io.BytesIO()
                wb.save(buf)
                zf.writestr(f'{comp}/总包标准_考勤表_{timestamp}.xlsx', buf.getvalue())
            except Exception as e:
                st.error(f'总包标准考勤表导出失败（{comp}）：{e}')

            # --- 4. 总包标准_工资表 ---
            try:
                wb = load_workbook('load-data/muban/zongbao/附件3：农民工资发放确认表(一式两份本人签字摁手印).xlsx')
                ws = wb.active
                start_r = 4
                for i, (_, row) in enumerate(comp_df.iterrows(), start=start_r):
                    ws.cell(row=i, column=1, value=i-start_r+1)
                    ws.cell(row=i, column=2, value=_value(row, '姓名'))
                    ws.cell(row=i, column=3, value=_value(row, '身份证号'))
                    ws.cell(row=i, column=4, value=_value(row, '联系电话', '手机号', '电话', '手机号码'))
                    ws.cell(row=i, column=5, value=_value(row, '开户银行'))
                    ws.cell(row=i, column=6, value=_value(row, '银行卡号', '工资卡号', '卡号'))
                    ws.cell(row=i, column=7, value=_value(row, '最终核定天数'))
                    ws.cell(row=i, column=8, value=_value(row, '应发工资'))
                buf = io.BytesIO()
                wb.save(buf)
                zf.writestr(f'{comp}/总包标准_工资确认表_{timestamp}.xlsx', buf.getvalue())
            except Exception as e:
                st.error(f'总包标准工资表导出失败（{comp}）：{e}')

    return zip_buf.getvalue()


# ================================================================
# 展示层
# ================================================================

def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="header-emoji">🗂️</span>
        <div class="header-text">
            <h2>考勤对账与工资结算</h2>
            <p>三方差异可视化 → 领导确认 → 在线定稿 → 一键多模板导出</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)

    tab_check, tab_export = st.tabs([
        ':material/fact_check: 考勤对账与在线定稿',
        ':material/payments: 考勤表与工资表导出',
    ])

    with tab_check:
        diff_df   = st.session_state.get('_att_diff_df')
        daily_df  = st.session_state.get('_att_daily_df')
        confirmed = st.session_state.get('_att_leader_confirmed', False)
        final_df  = st.session_state.get('final_attendance')

        # ── Step A：生成差异对比报告 ────────────────────────────
        with st.container(border=True):
            st.markdown("""
            <div class="step-indicator">
                <span class="step-num">A</span>
                <span>生成三方差异对比报告 → 发给领导确认</span>
            </div>
            """, unsafe_allow_html=True)

            st.caption('每家公司分别上传三份考勤，系统不会跨公司合并同名人员。')
            company_inputs = {}
            company_names = ['江苏旭之升建筑工程有限公司', '青海久昌建筑装饰工程有限公司']
            for company in company_names:
                with st.container(border=True):
                    st.markdown(f'**{company}**')
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        file_a = st.file_uploader('三局系统考勤', type=['xlsx', 'xls'], key=f'att_{company}_a')
                    with c2:
                        file_b = st.file_uploader('智慧护薪考勤', type=['xlsx', 'xls'], key=f'att_{company}_b')
                    with c3:
                        file_paper = st.file_uploader('纸质/水印考勤', type=['xlsx', 'xls'], key=f'att_{company}_paper')
                    company_inputs[company] = (file_a, file_b, file_paper)

            if st.button(':material/compare_arrows: 生成三方差异对比报告', type='primary', key='btn_gen_diff'):
                results = []
                daily_results = []
                missing_companies = []
                for company, (file_a, file_b, file_paper) in company_inputs.items():
                    if not all([file_a, file_b, file_paper]):
                        missing_companies.append(company)
                        continue
                    df_a = _safe_read(file_a, f'{company}-三局系统')
                    df_b = _safe_read(file_b, f'{company}-护薪系统')
                    df_p = _safe_read(file_paper, f'{company}-纸质核对')
                    results.append(build_diff_table(df_a, df_b, df_p, company=company))
                    daily_results.append(build_daily_diff_table(df_a, df_b, df_p, company=company))
                if missing_companies or len(results) != len(company_inputs):
                    st.warning(f':material/info: 请为以下公司各上传三份考勤后再执行：{"、".join(missing_companies)}')
                else:
                    with st.spinner('正在生成差异对比报告……'):
                        result = pd.concat(results, ignore_index=True)
                        daily_result = pd.concat(daily_results, ignore_index=True) if daily_results else pd.DataFrame()
                    st.session_state['_att_diff_df']            = result
                    st.session_state['_att_daily_df']           = daily_result
                    st.session_state['_att_missing_companies']   = missing_companies
                    st.session_state['_att_leader_confirmed']   = False
                    st.session_state.pop('final_attendance', None)
                    st.session_state.pop('_att_quickfill', None)
                    st.session_state.pop('att_data_editor', None)
                    st.rerun()

        # ── 差异报告展示（只读） ─────────────────────────────────
        if diff_df is not None and not diff_df.empty:
            n_ok   = (diff_df['差异状态'] == '✔ 一致').sum()
            n_warn = diff_df['差异状态'].str.startswith('⚠').sum()
            n_miss = (diff_df['差异状态'] == '— 数据缺失').sum()

            st.markdown(
                f'''
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-title">三方一致</div>
                        <div class="metric-value" style="color:#059669">{n_ok}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">存在差异</div>
                        <div class="metric-value" style="color:#D97706">{n_warn}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">数据缺失</div>
                        <div class="metric-value" style="color:#94A3B8">{n_miss}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">合计人数</div>
                        <div class="metric-value">{len(diff_df)}</div>
                    </div>
                </div>
                ''',
                unsafe_allow_html=True,
            )

            st.markdown('##### :material/table_view: 三方考勤对比一览（只读）')

            def _highlight_row(row):
                status = row['差异状态']
                if status == '✔ 一致':
                    bg = 'background-color:#D1FAE5;color:#065F46'
                elif status.startswith('⚠'):
                    bg = 'background-color:#FEF3C7;color:#92400E;font-weight:600'
                else:
                    bg = 'background-color:#F1F5F9;color:#94A3B8;font-style:italic'
                return [bg] * len(row)

            display_df = diff_df.copy()
            # 保持考勤天数列为纯数值类型；不要用 '' 混入 float 列，避免 Arrow 序列化失败。
            for col in ['三局天数', '护薪天数', '纸质天数']:
                if col in display_df.columns:
                    display_df[col] = pd.to_numeric(display_df[col], errors='coerce')
            styled = display_df.style.apply(_highlight_row, axis=1)
            st.dataframe(styled, use_container_width=True, hide_index=True)

            if daily_df is not None and not daily_df.empty:
                st.markdown('##### :material/calendar_month: 每日考勤原始明细（暂不决定最终输出格式）')
                st.caption('系统保留三份表格中每天的原始内容，并只做“有考勤/缺勤/待确认”识别。领导确认最终使用√还是时间段后，再接入最终输出。')
                daily_view = st.radio(
                    '每日明细查看范围',
                    ['仅看差异', '查看全部'],
                    horizontal=True,
                    key='att_daily_view_mode',
                )
                daily_view_df = daily_df
                if daily_view == '仅看差异':
                    daily_view_df = daily_df[daily_df['日差异状态'].str.startswith('⚠')]
                st.dataframe(daily_view_df, use_container_width=True, hide_index=True, height=420)

            st.markdown(
                '<div class="hint-box">'
                ':material/info: 下载彩色差异报告发给领导审阅。领导确认后点击下方按钮进入定稿环节。'
                '</div>',
                unsafe_allow_html=True,
            )

            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            col_dl, col_confirm, _ = st.columns([2, 2, 3])

            with col_dl:
                if OPENPYXL_OK:
                    st.download_button(
                        label=':material/download: 下载差异对比报告 (Excel)',
                        data=build_diff_report_xlsx(diff_df),
                        file_name=f'三方考勤差异报告_{timestamp}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='dl_diff_report', use_container_width=True,
                    )

            with col_confirm:
                if not confirmed:
                    if st.button(':material/how_to_reg: 领导已确认，进入定稿', key='btn_leader_confirm', type='primary', use_container_width=True):
                        st.session_state['_att_leader_confirmed'] = True
                        st.rerun()
                else:
                    st.markdown(
                        '<div class="celebrate-banner" style="margin:0;">'
                        '<span class="celebrate-icon">✅</span>'
                        '领导已确认，可在下方录入定稿天数。'
                        '</div>',
                        unsafe_allow_html=True,
                    )

            # ── Step B：定稿录入（门控） ─────────────────────────
            if confirmed:
                st.markdown('---')
                with st.container(border=True):
                    st.markdown("""
                    <div class="step-indicator">
                        <span class="step-num">B</span>
                        <span>定稿录入 — 根据领导批示填写最终核定天数</span>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(
                        '<div class="hint-box">'
                        '可先用下方<b>快速填充</b>按钮批量填入，再手动微调个别差异行；'
                        '或直接双击 <b>✏ 最终核定天数</b> 逐行编辑。'
                        '</div>',
                        unsafe_allow_html=True,
                    )

                    # ── 构建基础 DataFrame ─────────────────────────
                    edit_base = diff_df.copy()
                    for extra in ['工种', '班组', '身份证号', '联系电话', '开户银行', '银行卡号', '性别', '手机号', '分包/所属企业']:
                        if extra not in edit_base.columns:
                            edit_base[extra] = ''

                    def _default_final(row):
                        if row['差异状态'] == '✔ 一致':
                            for col in ['三局天数', '护薪天数', '纸质天数']:
                                v = row[col]
                                if not (isinstance(v, float) and np.isnan(v)):
                                    return v
                        return None

                    if '最终核定天数' not in edit_base.columns:
                        edit_base['最终核定天数'] = edit_base.apply(_default_final, axis=1)

                    # 优先加载已保存的定稿数据
                    def _row_key(row):
                        return f"{row.get('分包/所属企业', '')}|{row.get('身份证号', '')}|{row.get('姓名', '')}"

                    if final_df is not None and '最终核定天数' in final_df.columns:
                        saved_map = {_row_key(r): r.get('最终核定天数') for _, r in final_df.iterrows()}
                        edit_base['最终核定天数'] = edit_base.apply(
                            lambda r: saved_map.get(_row_key(r), r.get('最终核定天数')), axis=1
                        )

                    # 若有快速填充覆盖，使用 session 中存储的结果
                    if '_att_quickfill' in st.session_state:
                        qf = st.session_state['_att_quickfill']
                        edit_base['最终核定天数'] = edit_base.apply(
                            lambda r: qf.get(_row_key(r), r.get('最终核定天数')), axis=1
                        )

                    # ── 批量快速填充按钮区 ─────────────────────────
                    st.markdown('**:material/bolt: 批量快速填充核定天数：**')

                    def _do_fill(rule: str):
                        """根据规则批量填充，存入 session 并刷新。"""
                        result = {}
                        for _, r in edit_base.iterrows():
                            name = _row_key(r)
                            vals = {
                                '三局': r.get('三局天数'),
                                '护薪': r.get('护薪天数'),
                                '纸质': r.get('纸质天数'),
                            }
                            valid_vals = {
                                k: v for k, v in vals.items()
                                if v is not None and not (isinstance(v, float) and np.isnan(v))
                            }
                            if rule == '最大值':
                                result[name] = max(valid_vals.values()) if valid_vals else np.nan
                            elif rule in valid_vals:
                                result[name] = valid_vals[rule]
                            else:
                                # 该来源无数据，保留原值
                                orig = r.get('最终核定天数')
                                result[name] = orig if orig is not None else np.nan
                        st.session_state['_att_quickfill'] = result
                        # 清除编辑器缓存，确保 Streamlit 用新数据重绘
                        st.session_state.pop('att_data_editor', None)
                        st.rerun()

                    qcols = st.columns(4)
                    with qcols[0]:
                        if st.button(
                            ':material/trending_up: 按最大值填入',
                            key='qf_max', use_container_width=True,
                            help='哪个系统天数最多就填哪个（取三方最大值）',
                        ):
                            _do_fill('最大值')
                    with qcols[1]:
                        if st.button(
                            ':material/domain: 按三局系统填入',
                            key='qf_a', use_container_width=True,
                            help='全部按三局智慧工地平台的天数填入',
                        ):
                            _do_fill('三局')
                    with qcols[2]:
                        if st.button(
                            ':material/shield_person: 按护薪系统填入',
                            key='qf_b', use_container_width=True,
                            help='全部按智慧护薪平台的天数填入',
                        ):
                            _do_fill('护薪')
                    with qcols[3]:
                        if st.button(
                            ':material/description: 按纸质核对填入',
                            key='qf_paper', use_container_width=True,
                            help='全部按纸质核对表的天数填入',
                        ):
                            _do_fill('纸质')

                    st.markdown('<div class="soft-divider"></div>', unsafe_allow_html=True)

                    # ── 在线编辑器 ─────────────────────────────────
                    display_edit_cols = [
                        '分包/所属企业', '姓名', '身份证号', '三局天数', '护薪天数', '纸质天数',
                        '差异状态', '最终核定天数',
                        '工种', '班组', '联系电话', '开户银行', '银行卡号',
                    ]
                    display_edit_cols = [c for c in display_edit_cols if c in edit_base.columns]

                    col_cfg = {
                        '姓名':         st.column_config.TextColumn('姓名', disabled=True),
                        '分包/所属企业': st.column_config.TextColumn('公司', disabled=True),
                        '三局天数':     st.column_config.NumberColumn('三局系统天数', disabled=True, format='%.1f'),
                        '护薪天数':     st.column_config.NumberColumn('护薪系统天数', disabled=True, format='%.1f'),
                        '纸质天数':     st.column_config.NumberColumn('纸质天数',     disabled=True, format='%.1f'),
                        '差异状态':     st.column_config.TextColumn('差异状态',       disabled=True),
                        '最终核定天数': st.column_config.NumberColumn(
                            '✏ 最终核定天数',
                            min_value=0, max_value=31, step=0.5, format='%.1f',
                            help='批量填充后可双击单行微调',
                        ),
                        '工种':     st.column_config.TextColumn('工种'),
                        '班组':     st.column_config.TextColumn('班组'),
                        '身份证号': st.column_config.TextColumn('身份证号'),
                        '联系电话': st.column_config.TextColumn('联系电话'),
                        '开户银行': st.column_config.TextColumn('开户银行'),
                        '银行卡号': st.column_config.TextColumn('银行卡号'),
                    }

                    edited = st.data_editor(
                        edit_base[display_edit_cols],
                        column_config=col_cfg,
                        use_container_width=True,
                        num_rows='fixed',
                        hide_index=True,
                        key='att_data_editor',
                    )

                    c_save, c_info = st.columns([2, 5])
                    with c_save:
                        if st.button(':material/save: 保存定稿数据', type='primary', key='btn_save_final', use_container_width=True):
                            missing = edited[
                                edited['最终核定天数'].isna() &
                                (edited['差异状态'].str.startswith('⚠') | (edited['差异状态'] == '— 数据缺失'))
                            ]
                            if not missing.empty:
                                names = '、'.join(missing['姓名'].tolist())
                                st.warning(f':material/warning: 以下差异人员尚未填写核定天数：**{names}**')
                            else:
                                st.session_state['final_attendance'] = enrich_with_master(edited)
                                st.success(':material/check_circle: 定稿已保存！请切换至【考勤表与工资表导出】标签页。')
                    with c_info:
                        if final_df is not None:
                            st.markdown('<div class="hint-box" style="margin:0;">:material/info: 已有上次保存的定稿数据，重新保存将覆盖。</div>', unsafe_allow_html=True)

        elif diff_df is None:
            st.markdown('<div class="hint-box" style="margin-top:16px;">:material/arrow_upward: 请上传文件后点击【生成三方差异对比报告】。</div>', unsafe_allow_html=True)

        if diff_df is not None or final_df is not None:
            with st.expander(':material/refresh: 重置本期对账（开始新一期）'):
                st.warning('将清除当前所有对账及定稿数据，无法撤销。')
                if st.button(':material/delete_forever: 确认重置', key='btn_reset_tab1', type='secondary'):
                    for k in ['_att_diff_df', '_att_daily_df', '_att_leader_confirmed', 'final_attendance']:
                        st.session_state.pop(k, None)
                    st.rerun()

    # ============================================================
    # Tab 2 — 导出
    # ============================================================
    with tab_export:
        final_df = st.session_state.get('final_attendance')

        if final_df is None or final_df.empty:
            st.markdown(
                '<div class="alert-box alert-danger" style="margin-top:20px;">'
                ':material/warning: 尚未完成考勤定稿。请先在【考勤对账与在线定稿】标签页完成 Step B 并保存。'
                '</div>',
                unsafe_allow_html=True,
            )
            return

        with st.container(border=True):
            st.markdown('#### :material/calculate: 工资联动计算')
            wage_mode = st.radio('日薪配置模式', ['使用主表日薪', '全局统一日薪', '按工种分别设置'], horizontal=True, key='wage_mode')
            salary_df = enrich_with_master(final_df)

            if wage_mode == '使用主表日薪':
                salary_df['日薪'] = pd.to_numeric(salary_df.get('提取日薪', 0), errors='coerce')
                missing_wage = salary_df['日薪'].isna() | (salary_df['日薪'] <= 0)
                if missing_wage.any():
                    st.warning(f'有 {int(missing_wage.sum())} 人未能从 data/master 取得有效日薪，请核对主表后再导出。')
            elif wage_mode == '全局统一日薪':
                global_wage = st.number_input('统一日薪（元/天）', min_value=0.0, value=300.0, step=10.0, key='global_wage')
                salary_df['日薪'] = global_wage
            else:
                unique_types = [t for t in final_df.get('工种', pd.Series(dtype=str)).dropna().unique() if str(t).strip()] if '工种' in final_df.columns else []
                if not unique_types:
                    st.info('未检测到工种信息，将使用默认 300 元/天。')
                    salary_df['日薪'] = 300.0
                else:
                    cols = st.columns(min(len(unique_types), 4))
                    wage_map = {}
                    for idx, wt in enumerate(unique_types):
                        with cols[idx % 4]:
                            w = st.number_input(f'「{wt}」日薪', min_value=0.0, value=300.0, step=10.0, key=f'wage_{wt}')
                            wage_map[str(wt).strip()] = w
                    salary_df['日薪'] = salary_df['工种'].apply(lambda x: wage_map.get(str(x).strip(), 300.0))

            def _calc(row):
                try:
                    return round(float(row['最终核定天数']) * float(row['日薪']), 2)
                except (ValueError, TypeError):
                    return ''

            salary_df['应发工资'] = salary_df.apply(_calc, axis=1)

        st.markdown('---')
        st.markdown('#### :material/table_view: 工资预览')
        preview_cols = [c for c in ['姓名','工种','班组','最终核定天数','日薪','应发工资'] if c in salary_df.columns]
        st.dataframe(salary_df[preview_cols], use_container_width=True, hide_index=True)

        valid_salaries = pd.to_numeric(salary_df['应发工资'], errors='coerce').dropna()
        total_pay = valid_salaries.sum()
        avg_days  = pd.to_numeric(salary_df['最终核定天数'], errors='coerce').mean()

        st.markdown(
            f'''
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">参与结算人数</div><div class="metric-value">{len(salary_df)}</div></div>
                <div class="metric-card"><div class="metric-title">平均核定天数</div><div class="metric-value">{avg_days:.1f} 天</div></div>
                <div class="metric-card"><div class="metric-title">工资总额</div><div class="metric-value">¥{total_pay:,.0f}</div></div>
            </div>
            ''',
            unsafe_allow_html=True,
        )

        st.markdown('---')
        st.markdown('#### :material/download: 按标准模板一键导出')
        st.markdown('<div class="hint-box">系统已读取 `load-data/muban` 目录下的 4 套官方模板，按公司自动分拣，保留所有红头、签字栏及样式进行套打。</div>', unsafe_allow_html=True)

        if not OPENPYXL_OK:
            st.error('未安装 openpyxl，请运行 pip install openpyxl 后重启应用。')
        else:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            
            with st.container(border=True):
                st.markdown('<div style="text-align:center;padding:12px 0"><span style="font-size:40px">🗂️</span><br><b>全量台账压缩包 (ZIP)</b><br><small style="color:#64748B">内含：公司标准考勤、工资表 + 总包标准考勤、工资表<br>已按所属分公司分类整理</small></div>', unsafe_allow_html=True)
                
                zip_data = build_all_exports_zip(salary_df)
                if zip_data:
                    st.download_button(
                        label=':material/archive: 一键打包下载全部 4 套台账',
                        data=zip_data,
                        file_name=f'劳务考勤发薪全套台账_{timestamp}.zip',
                        mime='application/zip',
                        use_container_width=True,
                        type='primary'
                    )
                else:
                    st.error("生成打包文件失败，请检查模板文件是否存在。")
