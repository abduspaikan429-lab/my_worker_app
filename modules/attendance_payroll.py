# -*- coding: utf-8 -*-
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
    if not text or text in {"--", "—", "-", "/", "0"}:
        return "缺勤/无记录"
    if any(mark in text for mark in ["√", "✓", "✔", "出勤", "正常"]):
        return "有考勤"
    if re.search(r"\d{1,2}:\d{2}", text):
        return "有考勤"
    try:
        if float(text) > 0:
            return "有考勤"
    except (ValueError, TypeError):
        pass
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

def parse_watermark_attendance(df, company='', period=''):
    """
    单源水印签到表解析。
    基于 _safe_read 的结果，优先通过识别每日考勤计算出勤天数，
    如果未找到每日明细，则尝试读取总天数列。
    """
    if df.empty:
        return __import__('pandas').DataFrame()
    
    days_col = _resolve_col(df, _DAYS_CANDS)
    records = _extract_daily_source(df, '水印签到表')
    
    rows = []
    if records:
        for key, item in records.items():
            name = item.get('姓名', '')
            identity = item.get('身份证号', '')
            days_dict = item.get('days', {})
            
            actual_days = 0
            has_daily_records = len(days_dict) > 0
            
            if has_daily_records:
                for day, val in days_dict.items():
                    state = _attendance_state(val)
                    if state == '有考勤':
                        actual_days += 1
            
            if not has_daily_records and days_col:
                name_col = _resolve_col(df, _NAME_CANDS)
                mask = (df[name_col] == name)
                if identity:
                    id_c = _resolve_col(df, _ID_CANDS)
                    if id_c:
                        mask = mask & (df[id_c].apply(_clean_identity) == identity)
                match = df[mask]
                if not match.empty:
                    val = _to_float(match.iloc[0][days_col])
                    actual_days = val if not __import__('pandas').isna(val) else 0
                    
            rows.append({
                '分包/所属企业': company,
                '姓名': name,
                '身份证号': identity,
                '月份': period,
                '__考勤明细__': dict(days_dict),
                '解析出勤天数': float(actual_days),
                '最终核定天数': float(actual_days)
            })
    else:
        name_col = _resolve_col(df, _NAME_CANDS)
        id_col = _resolve_col(df, _ID_CANDS)
        for _, row in df.iterrows():
            name = _raw_attendance_value(row.get(name_col, ''))
            identity = _clean_identity(row.get(id_col, '')) if id_col else ''
            if not _looks_like_person_name(name) and not _is_valid_identity(identity):
                continue
            actual_days = _to_float(row.get(days_col)) if days_col else 0
            if __import__('pandas').isna(actual_days):
                actual_days = 0
            rows.append({
                '分包/所属企业': company,
                '姓名': name,
                '身份证号': identity,
                '月份': period,
                '__考勤明细__': {},
                '解析出勤天数': float(actual_days),
                '最终核定天数': float(actual_days)
            })

    res_df = __import__('pandas').DataFrame(rows)
    if not res_df.empty:
        res_df['__key__'] = res_df.apply(lambda r: r['身份证号'] if _is_valid_identity(r['身份证号']) else f"姓名:{r['姓名']}", axis=1)
        res_df.drop_duplicates(subset='__key__', keep='first', inplace=True)
        res_df.drop(columns=['__key__'], inplace=True)
    return res_df


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
            if pd.isna(val) or not str(val).strip():
                return None
            match = re.search(r'(\d+(?:\.\d+)?)', str(val))
            return float(match.group(1)) if match else None

        wage_col = next((c for c in ['结算单价/标准', '日薪', '日工资', '工资标准', '单价', '基本工资', '提取日薪'] if c in df.columns), None)
        if wage_col:
            df['提取日薪'] = df[wage_col].apply(extract_wage)
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


def _is_empty_val(v):
    if v is None or pd.isna(v):
        return True
    s = str(v).strip()
    return s == '' or s.lower() in ['nan', 'none', 'null', '<na>', '0', '0.0']


def enrich_with_master(final_df):
    """
    按主表权威数据补齐人员档案与日薪：
    1. 身份证号精确匹配（优先，支持大小写与去空格）
    2. 姓名 + 班组 匹配
    3. 姓名 匹配（唯一姓名优先，非唯一取首条保底）
    主表作为最终权威来源，提取日薪、身份证号、银行卡等自动以主表最新数据为准。
    """
    master = _load_master_data()
    if final_df is None or final_df.empty or master is None or master.empty:
        return final_df.copy() if isinstance(final_df, pd.DataFrame) else pd.DataFrame()

    out = final_df.copy()

    # 建立主表索引
    id_map = {}
    for _, r in master.iterrows():
        sid = str(r.get('身份证号', '')).strip().upper()
        if sid and len(sid) >= 15:
            id_map[sid] = r.to_dict()

    name_team_map = {}
    for _, r in master.iterrows():
        nm = str(r.get('姓名', '')).strip()
        tm = str(r.get('班组', '')).strip()
        if nm and tm:
            name_team_map[f"{nm}_{tm}"] = r.to_dict()

    unique_names = master[master['姓名'].ne('')].groupby('姓名').filter(lambda x: len(x) == 1)
    name_map = {str(r.get('姓名', '')).strip(): r.to_dict() for _, r in unique_names.iterrows()}

    all_name_map = {}
    for _, r in master.iterrows():
        nm = str(r.get('姓名', '')).strip()
        if nm and nm not in all_name_map:
            all_name_map[nm] = r.to_dict()

    fields_to_enrich = [
        '身份证号', '姓名', '性别', '手机号', '班组', '工种', '人员类型',
        '分包/所属企业', '工资卡号', '开户银行', '提取日薪', '项目全称'
    ]

    for col in fields_to_enrich:
        if col not in out.columns:
            out[col] = ''

    for idx in range(len(out)):
        sid = str(out.at[idx, '身份证号'] if '身份证号' in out.columns else '').strip().upper()
        nm = str(out.at[idx, '姓名'] if '姓名' in out.columns else '').strip()
        tm = str(out.at[idx, '班组'] if '班组' in out.columns else '').strip()

        matched_record = None
        if sid and len(sid) >= 15 and sid in id_map:
            matched_record = id_map[sid]
        elif f"{nm}_{tm}" in name_team_map:
            matched_record = name_team_map[f"{nm}_{tm}"]
        elif nm in name_map:
            matched_record = name_map[nm]
        elif nm in all_name_map:
            matched_record = all_name_map[nm]

        if matched_record:
            # 提取日薪：主表有有效日薪时直接采用
            master_wage = matched_record.get('提取日薪')
            if master_wage is not None and not pd.isna(master_wage) and float(master_wage or 0) > 0:
                out.at[idx, '提取日薪'] = float(master_wage)
            elif _is_empty_val(out.at[idx, '提取日薪']):
                if master_wage is not None and not pd.isna(master_wage):
                    out.at[idx, '提取日薪'] = master_wage

            # 其它关键字段覆盖或补齐
            for f in fields_to_enrich:
                if f == '提取日薪':
                    continue
                mv = matched_record.get(f, '')
                if mv is not None and not pd.isna(mv) and str(mv).strip() != '':
                    cur_val = out.at[idx, f]
                    if _is_empty_val(cur_val) or f in ['身份证号', '工资卡号', '开户银行', '分包/所属企业', '班组', '工种']:
                        out.at[idx, f] = mv

    # 统一导出字段别名，避免模板导出阶段因界面字段名不同而漏填电话/卡号。
    if '联系电话' not in out.columns:
        out['联系电话'] = ''
    if '手机号' in out.columns:
        out['联系电话'] = out['联系电话'].replace('', pd.NA).fillna(out['手机号'])
    if '银行卡号' not in out.columns:
        out['银行卡号'] = ''
    if '工资卡号' in out.columns:
        out['银行卡号'] = out['银行卡号'].replace('', pd.NA).fillna(out['工资卡号'])
    if '工资卡号' not in out.columns:
        out['工资卡号'] = out['银行卡号']
    else:
        out['工资卡号'] = out['工资卡号'].replace('', pd.NA).fillna(out['银行卡号'])

    return out


import zipfile
from openpyxl import load_workbook

def _excel_value(row, *names):
    """按候选字段读取值；跳过 None、NaN 与空文本，回退到后续候选字段。"""
    for name in names:
        if hasattr(row, 'get'):
            value = row.get(name)
        elif hasattr(row, '__getitem__') and name in row:
            value = row[name]
        else:
            continue
        if value is None:
            continue
        try:
            if pd.isna(value):
                continue
        except (TypeError, ValueError):
            pass
        if str(value).strip() == '':
            continue
        return value
    return ''


def _excel_text(value):
    if value is None:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _period_label(value):
    text = _excel_text(value)
    if not text:
        return ''
    match = re.search(r'(\d{1,2})\s*月', text)
    if match:
        return f'{int(match.group(1))}月'
    match = re.search(r'(?<!\d)(\d{1,2})(?!\d)', text)
    return f'{int(match.group(1))}月' if match else text


def _period_number(value):
    match = re.search(r'(\d{1,2})\s*月', _excel_text(value))
    if match:
        return int(match.group(1))
    return None


def _infer_period(value):
    """从上传文件名或标题提取月份，缺失时保留空值而不擅自改变业务月份。"""
    match = re.search(r'(\d{1,2})\s*月', _excel_text(value))
    return f'{int(match.group(1))}月' if match else ''


def _short_company(company, attendance=False):
    text = _excel_text(company)
    if '江苏旭之升' in text:
        return '旭之升' if attendance else '江苏旭之升'
    if '青海久昌' in text:
        return '青海久昌'
    return text or '未知公司'


def _first_nonempty(df, columns, fallback=''):
    for column in columns:
        if column not in df.columns:
            continue
        values = df[column].dropna().map(_excel_text)
        values = values[values.ne('')]
        if not values.empty:
            return values.iloc[0]
    return fallback


def _export_batches(salary_df):
    """按公司+月份组织批次；一个批次最终对应一个明细 sheet。"""
    if salary_df is None or salary_df.empty:
        return []
    data = salary_df.copy()
    if '分包/所属企业' not in data.columns:
        data['分包/所属企业'] = '未知公司'
    if '月份' not in data.columns:
        data['月份'] = ''
    data['分包/所属企业'] = data['分包/所属企业'].map(_excel_text).replace('', '未知公司')
    data['月份'] = data['月份'].map(_period_label)
    batches = []
    for (company, period), group in data.groupby(['分包/所属企业', '月份'], sort=False, dropna=False):
        batches.append({
            'company': _excel_text(company) or '未知公司',
            'period': _period_label(period),
            'df': group.reset_index(drop=True),
        })
    return batches


def _safe_sheet_title(title, used_titles):
    title = re.sub(r'[\\/*?:\[\]]', '', _excel_text(title)) or '明细'
    title = title[:31]
    base = title
    suffix = 2
    while title in used_titles:
        tail = f'_{suffix}'
        title = f'{base[:31-len(tail)]}{tail}'
        suffix += 1
    used_titles.add(title)
    return title


def _copy_template_sheets(template_path, batches, name_builder, company_specific=True):
    wb = load_workbook(template_path)
    source_pairs = []
    for batch in batches:
        source = _select_company_sheet(wb, batch['company']) if company_specific else wb.active
        source_pairs.append((batch, source))

    outputs = []
    used_titles = set()
    used_sources = set()

    for batch, source in source_pairs:
        target_title = _safe_sheet_title(name_builder(batch), used_titles)
        if source not in used_sources:
            # 优先直接复用模板原 sheet，保留页眉图片/Logo、图表与打印排版（openpyxl copy_worksheet 会丢失图片）
            sheet = source
            used_sources.add(source)
        else:
            sheet = wb.copy_worksheet(source)

        existing = wb[target_title] if target_title in wb.sheetnames else None
        if existing and existing != sheet:
            existing.title = f"__temp_{id(existing)}"

        sheet.title = target_title
        outputs.append((batch, sheet))

    output_sheets = {sheet for _, sheet in outputs}
    for sheet in list(wb.worksheets):
        if sheet not in output_sheets:
            wb.remove(sheet)
    if wb.worksheets:
        wb.active = 0
    return wb, outputs


def _copy_worksheet_content(source_ws, target_ws):
    for row in source_ws.iter_rows():
        for cell in row:
            target_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                target_cell.font = copy(cell.font)
                target_cell.border = copy(cell.border)
                target_cell.fill = copy(cell.fill)
                target_cell.number_format = copy(cell.number_format)
                target_cell.protection = copy(cell.protection)
                target_cell.alignment = copy(cell.alignment)
    for col_letter, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_idx].height = dim.height
    for merged in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged))


def _clear_rows(ws, start_row, end_row, start_col=1, end_col=None):
    if end_row < start_row:
        return
    end_col = end_col or ws.max_column
    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(start_col, end_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            try:
                cell.value = None
            except AttributeError:
                pass


def _find_attendance_footer_row(ws, start_row):
    keywords = ('分包项目负责人', '分包班组长', '制表人', '总包负责人', '申明', '声明', '第  1  页', '第1页')
    for row_idx in range(start_row, ws.max_row + 1):
        values = [ws.cell(row_idx, col_idx).value for col_idx in range(1, ws.max_column + 1)]
        text = ' '.join(_excel_text(value) for value in values if value is not None)
        if any(keyword in text for keyword in keywords):
            return row_idx
    return ws.max_row + 1


def _prepare_data_area(ws, start_row, count, footer_finder):
    footer_row = footer_finder(ws, start_row)
    available = max(0, footer_row - start_row)
    if count > available:
        _ensure_data_rows(ws, start_row, count, footer_row=footer_row)
        footer_row = footer_finder(ws, start_row)
    elif count < available:
        # 删除多余空行，彻底避免留白行
        to_delete = available - count
        delete_start = start_row + count
        footer_merges = []
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= footer_row:
                footer_merges.append((merged.min_row, merged.min_col, merged.max_row, merged.max_col))
                ws.unmerge_cells(str(merged))

        ws.delete_rows(delete_start, to_delete)

        for min_row, min_col, max_row, max_col in footer_merges:
            ws.merge_cells(
                start_row=min_row - to_delete,
                start_column=min_col,
                end_row=max_row - to_delete,
                end_column=max_col,
            )
        footer_row = start_row + count

    _clear_rows(ws, start_row, footer_row - 1)
    return footer_row


def _header_row(ws, required='姓名'):
    for row_idx in range(1, min(ws.max_row, 12) + 1):
        labels = [_normalize_header(ws.cell(row_idx, col_idx).value) for col_idx in range(1, ws.max_column + 1)]
        if any(required in label for label in labels):
            return row_idx
    return 1


def _header_col(ws, row_idx, labels):
    normalized_labels = [_normalize_header(label) for label in labels]
    for col_idx in range(1, ws.max_column + 1):
        current = _normalize_header(ws.cell(row_idx, col_idx).value)
        if any(label == current or label in current for label in normalized_labels):
            return col_idx
    return None


def _write_text_cell(ws, row_idx, col_idx, value):
    if not col_idx:
        return
    cell = ws.cell(row=row_idx, column=col_idx)
    cell.value = _excel_text(value)
    cell.number_format = '@'


def _write_number_cell(ws, row_idx, col_idx, value):
    if not col_idx:
        return
    text = _excel_text(value)
    if not text:
        ws.cell(row=row_idx, column=col_idx).value = None
        return
    number = _to_float(text)
    if pd.isna(number):
        ws.cell(row=row_idx, column=col_idx).value = value
    else:
        ws.cell(row=row_idx, column=col_idx).value = int(number) if float(number).is_integer() else number


def _daily_marks(row):
    marks = _excel_value(row, '__考勤明细__')
    return marks if isinstance(marks, dict) else {}


def _write_daily_cells(ws, row_idx, row, marker=8, start_col=4, day_count=31):
    marks = _daily_marks(row)
    font_day = Font(name='仿宋_GB2312', size=9)
    align_center = Alignment(horizontal='center', vertical='center')
    bd = _thin_border()
    for day in range(1, day_count + 1):
        value = marks.get(day, marks.get(str(day), ''))
        cell = ws.cell(row=row_idx, column=start_col + day - 1)
        cell.font = font_day
        cell.alignment = align_center
        cell.number_format = 'General'
        cell.border = bd
        if _attendance_state(value) == '有考勤':
            if str(value).strip() in ('0.5', '半天', '4', '4.0'):
                cell.value = 4 if marker == 8 else value
            else:
                cell.value = marker
        else:
            cell.value = None


def _batch_project(group):
    return _first_nonempty(group, ['项目全称', '项目简称'], '科技文化中心—国际体育中心（足球场项目）')


def _batch_team(group):
    return _first_nonempty(group, ['班组'], '各班组')


def _set_attendance_sheet_header(ws, batch, standard):
    group = batch['df']
    project = _batch_project(group)
    team = _batch_team(group)
    period = batch['period']
    if not period:
        return
    year = datetime.now().year
    month_number = _period_number(period)
    month_text = f'{year}年{month_number}月' if month_number else period
    if standard == 'zongbao':
        ws['A1'] = f'{project}工程{period}考勤表'
        ws['A2'] = batch['company']
        ws['O2'] = f'班组：{team}'
        ws['AE2'] = f'日期：{month_text}'
    else:
        ws['A4'] = f'项目名称（全称）：{project}                        班组名称：{team}                        {month_text}'


def _normalize_company_attendance_grid(ws):
    """将公司模板的“上月26-31+本月1-25”表头统一为参考表的1-31日列。"""
    header_styles = {}
    for col_idx in range(1, 36):
        source = ws.cell(row=6, column=col_idx)
        header_styles[col_idx] = copy(source._style)
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row <= 6 and merged.max_row >= 5 and merged.min_col <= 35 and merged.max_col >= 1:
            ws.unmerge_cells(str(merged))
    ws['A5'] = '编号'
    ws['B5'] = '姓名'
    ws['C5'] = '工种'
    for day in range(1, 32):
        cell = ws.cell(row=5, column=3 + day)
        cell.value = day
        if header_styles[3 + day]:
            cell._style = copy(header_styles[3 + day])
    ws['AI5'] = '合计'
    if header_styles[35]:
        ws['AI5']._style = copy(header_styles[35])
    return 6


def _populate_company_attendance_sheet(ws, batch):
    group = batch['df']
    # 考勤为0的直接不要做到考勤表里面
    group = group[pd.to_numeric(group.get('最终核定天数', 0), errors='coerce').fillna(0) > 0].copy()
    start_row = _normalize_company_attendance_grid(ws)
    footer_row = _prepare_data_area(ws, start_row, len(group), _find_attendance_footer_row)
    _set_attendance_sheet_header(ws, batch, 'company')

    font_data = Font(name='仿宋_GB2312', size=10)
    align_center = Alignment(horizontal='center', vertical='center')
    bd = _thin_border()

    for row_idx, (_, row) in enumerate(group.iterrows(), start=start_row):
        ws.row_dimensions[row_idx].height = 19.5
        # 编号
        c_no = ws.cell(row=row_idx, column=1, value=row_idx - start_row + 1)
        c_no.font = font_data
        c_no.alignment = align_center
        c_no.number_format = 'General'
        c_no.border = bd

        # 姓名
        _write_text_cell(ws, row_idx, 2, _excel_value(row, '姓名'))
        c_name = ws.cell(row=row_idx, column=2)
        c_name.font = font_data
        c_name.alignment = align_center
        c_name.border = bd

        # 工种
        _write_text_cell(ws, row_idx, 3, _excel_value(row, '工种'))
        c_job = ws.cell(row=row_idx, column=3)
        c_job.font = font_data
        c_job.alignment = align_center
        c_job.border = bd

        # 每日考勤打卡工时 (统一数字8与半天4)
        _write_daily_cells(ws, row_idx, row, 8)

        # 合计天数（3、考勤表在合计处填上公式以便合计=COUNTIF(D6:AH6,8)+COUNTIF(D6:AH6,4)*0.5）
        c_total = ws.cell(row=row_idx, column=35)
        c_total.value = f"=COUNTIF(D{row_idx}:AH{row_idx},8)+COUNTIF(D{row_idx}:AH{row_idx},4)*0.5"
        c_total.font = font_data
        c_total.alignment = align_center
        c_total.number_format = 'General'
        c_total.border = bd

    # 考勤总计（2、考勤的总计没有填上）
    c_sum = ws.cell(row=footer_row, column=35)
    if footer_row > start_row:
        c_sum.value = f"=SUM(AI{start_row}:AI{footer_row - 1})"
    else:
        c_sum.value = 0
    c_sum.font = font_data
    c_sum.alignment = align_center
    c_sum.number_format = 'General'
    c_sum.border = bd

    return {'start_row': start_row, 'footer_row': footer_row}


def _populate_zongbao_attendance_sheet(ws, batch):
    group = batch['df']
    # 考勤为0的直接不要做到考勤表里面
    group = group[pd.to_numeric(group.get('最终核定天数', 0), errors='coerce').fillna(0) > 0].copy()
    start_row = 5
    footer_row = _prepare_data_area(ws, start_row, len(group), _find_attendance_footer_row)
    _set_attendance_sheet_header(ws, batch, 'zongbao')
    for row_idx, (_, row) in enumerate(group.iterrows(), start=start_row):
        ws.cell(row=row_idx, column=1, value=row_idx - start_row + 1)
        _write_text_cell(ws, row_idx, 2, _excel_value(row, '姓名'))
        _write_text_cell(ws, row_idx, 3, _excel_value(row, '性别'))
        _write_daily_cells(ws, row_idx, row, 8)
        ws.cell(row=row_idx, column=35).value = f"=COUNTIF(D{row_idx}:AH{row_idx},8)+COUNTIF(D{row_idx}:AH{row_idx},4)*0.5"
    if footer_row > start_row:
        ws.cell(row=footer_row, column=35, value=f"=SUM(AI{start_row}:AI{footer_row - 1})")
    else:
        ws.cell(row=footer_row, column=35, value=0)
    return {'start_row': start_row, 'footer_row': footer_row}


def _set_total_wage_header(ws, batch):
    period = batch['period']
    if not period:
        return
    ws['A2'] = f'项目名称：{_batch_project(batch["df"])}'
    ws['D2'] = f'劳务单位：{batch["company"]}'
    ws['H2'] = f'月度：{period}'


def _populate_zongbao_wage_sheet(ws, batch):
    group = batch['df']
    group = group[pd.to_numeric(group.get('最终核定天数', 0), errors='coerce').fillna(0) > 0].copy()
    start_row = 4
    footer_row = _prepare_data_area(ws, start_row, len(group), lambda sheet, start: _find_footer_row(sheet, start))
    _set_total_wage_header(ws, batch)
    for row_idx, (_, row) in enumerate(group.iterrows(), start=start_row):
        ws.cell(row=row_idx, column=1, value=row_idx - start_row + 1)
        _write_text_cell(ws, row_idx, 2, _excel_value(row, '姓名'))
        _write_text_cell(ws, row_idx, 3, _excel_value(row, '身份证号'))
        _write_text_cell(ws, row_idx, 4, _excel_value(row, '联系电话', '手机号', '电话', '手机号码'))
        _write_text_cell(ws, row_idx, 5, _excel_value(row, '开户银行'))
        _write_text_cell(ws, row_idx, 6, _excel_value(row, '银行卡号', '工资卡号', '卡号'))
        _write_number_cell(ws, row_idx, 7, _excel_value(row, '最终核定天数'))
        _write_number_cell(ws, row_idx, 8, _excel_value(row, '应发工资'))
    ws.column_dimensions['C'].width = max(ws.column_dimensions['C'].width or 0, 24)
    ws.column_dimensions['D'].width = max(ws.column_dimensions['D'].width or 0, 16)
    ws.column_dimensions['F'].width = max(ws.column_dimensions['F'].width or 0, 24)
    return {'start_row': start_row, 'footer_row': footer_row, 'pay_col': 8}


def _set_wage_sheet_header(ws, comp_df, company, period=''):
    """更新新版公司工资模板的合并表头，按导出批次写入月份。"""
    project = _first_nonempty(comp_df, ['项目全称', '项目简称'], '科技文化中心—国际体育中心（足球场项目）')
    team = _first_nonempty(comp_df, ['班组'], '各班组')
    original = str(ws['A2'].value or '')
    month_match = re.search(r'\d{4}\s*年[^\n]*?月', original)
    month_text = month_match.group(0) if month_match else '年  月'
    if period:
        month_number = _period_number(period)
        month_text = f'{datetime.now().year}年{month_number}月' if month_number else period
    ws['A2'] = f'项目名称（全称）：{project}             班组名称：{team}                    {month_text}'


def _populate_company_wage_sheet(ws, batch, att_footer_row=None, att_sheet_name='考勤表'):
    group = batch['df']
    # 与考勤表严格对齐：按最终核定天数大于0过滤，确保两表人员名单与顺序1-to-1完全一致
    group = group[pd.to_numeric(group.get('最终核定天数', 0), errors='coerce').fillna(0) > 0].copy()

    start_row = 5
    count = len(group)
    total_count = count + 1  # 工人行 + 底部合计行
    footer_row = _prepare_data_area(ws, start_row, total_count, lambda sheet, start: _find_footer_row(sheet, start))
    _set_wage_sheet_header(ws, group, batch['company'], batch['period'])
    header_row = _header_row(ws, '姓名')
    name_col = _header_col(ws, header_row, ['姓名']) or 2
    job_col = _header_col(ws, header_row, ['工种']) or 3
    days_col = _header_col(ws, header_row, ['出勤工日', '考勤天数']) or 4
    rate_col = _header_col(ws, header_row, ['工资标准', '日薪']) or 5
    gross_col = _header_col(ws, header_row, ['工资总额', '应发工资']) or 6
    allowance_col = _header_col(ws, header_row, ['津贴']) or 7
    pay_col = _header_col(ws, header_row, ['应支付', '实发工资']) or 12
    card_col = _header_col(ws, header_row, ['银行卡号', '工资卡号', '卡号', '银行卡']) or 15

    days_let = get_column_letter(days_col)
    rate_let = get_column_letter(rate_col)
    gross_let = get_column_letter(gross_col)
    alw_let = get_column_letter(allowance_col)
    pay_let = get_column_letter(pay_col)

    font_data = Font(name='宋体', size=9)
    font_bold = Font(name='宋体', size=9, bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    bd = _thin_border()

    att_end = max(6, (att_footer_row - 1) if att_footer_row else (6 + count - 1))

    for idx, (_, row) in enumerate(group.iterrows()):
        row_idx = start_row + idx
        ws.cell(row=row_idx, column=1, value=idx + 1)
        _write_text_cell(ws, row_idx, name_col, _excel_value(row, '姓名'))
        _write_text_cell(ws, row_idx, job_col, _excel_value(row, '工种'))

        # 出勤日期用VLOOKUP公式从考勤表引用
        name_ref = f"{get_column_letter(name_col)}{row_idx}"
        c_days = ws.cell(row=row_idx, column=days_col)
        c_days.value = f"=VLOOKUP({name_ref},{att_sheet_name}!$B$6:$AI${att_end},34,FALSE)"
        c_days.font = font_data
        c_days.alignment = align_center
        c_days.number_format = 'General'
        c_days.border = bd

        _write_number_cell(ws, row_idx, rate_col, _excel_value(row, '日薪', '提取日薪'))

        # 工资总额公式：出勤工日*工资标准
        c_gross = ws.cell(row=row_idx, column=gross_col)
        c_gross.value = f"={days_let}{row_idx}*{rate_let}{row_idx}"
        c_gross.font = font_data
        c_gross.alignment = align_center
        c_gross.border = bd

        # 津贴列
        alw_val = row.get('津贴', 0)
        _write_number_cell(ws, row_idx, allowance_col, alw_val if not pd.isna(alw_val) else 0)
        c_alw = ws.cell(row=row_idx, column=allowance_col)
        c_alw.font = font_data
        c_alw.alignment = align_center
        c_alw.border = bd

        # 应支付公式：工资总额+津贴
        c_pay = ws.cell(row=row_idx, column=pay_col)
        c_pay.value = f"={gross_let}{row_idx}+{alw_let}{row_idx}"
        c_pay.font = font_data
        c_pay.alignment = align_center
        c_pay.border = bd

        _write_text_cell(ws, row_idx, card_col, _excel_value(row, '银行卡号', '工资卡号', '卡号', '银行卡'))

    # 在序号最后一行新增一行，合并单元格作为合计，合计工资上面的总额
    total_row_idx = start_row + count
    ws.row_dimensions[total_row_idx].height = 20
    ws.merge_cells(start_row=total_row_idx, start_column=1, end_row=total_row_idx, end_column=5)
    c_tot_label = ws.cell(row=total_row_idx, column=1, value='合计')
    c_tot_label.font = font_bold
    c_tot_label.alignment = align_center
    for c in range(1, 6):
        ws.cell(row=total_row_idx, column=c).border = bd

    last_worker_row = total_row_idx - 1
    # 工资总额合计
    c_tot_gross = ws.cell(row=total_row_idx, column=gross_col)
    if count > 0:
        c_tot_gross.value = f"=SUM({gross_let}{start_row}:{gross_let}{last_worker_row})"
    else:
        c_tot_gross.value = 0
    c_tot_gross.font = font_bold
    c_tot_gross.alignment = align_center
    c_tot_gross.border = bd

    # 津贴合计
    c_tot_alw = ws.cell(row=total_row_idx, column=allowance_col)
    if count > 0:
        c_tot_alw.value = f"=SUM({alw_let}{start_row}:{alw_let}{last_worker_row})"
    else:
        c_tot_alw.value = 0
    c_tot_alw.font = font_bold
    c_tot_alw.alignment = align_center
    c_tot_alw.border = bd

    # 扣款部分单元格补齐边框
    for c in range(allowance_col + 1, pay_col):
        ws.cell(row=total_row_idx, column=c).border = bd

    # 应支付合计
    c_tot_pay = ws.cell(row=total_row_idx, column=pay_col)
    if count > 0:
        c_tot_pay.value = f"=SUM({pay_let}{start_row}:{pay_let}{last_worker_row})"
    else:
        c_tot_pay.value = 0
    c_tot_pay.font = font_bold
    c_tot_pay.alignment = align_center
    c_tot_pay.border = bd

    # 其余列补齐边框
    for c in range(pay_col + 1, ws.max_column + 1):
        ws.cell(row=total_row_idx, column=c).border = bd

    if card_col:
        col_letter = get_column_letter(card_col)
        ws.column_dimensions[col_letter].width = max(ws.column_dimensions[col_letter].width or 0, 26)

    return {'start_row': start_row, 'footer_row': footer_row, 'pay_col': pay_col, 'total_row': total_row_idx}


def _add_company_wage_summary(wb, details):
    summary = wb.create_sheet('总计', 0)
    summary.merge_cells('A1:E1')
    summary['A1'] = '中建二局安装工程有限公司-工资发放表'
    summary['A1'].font = Font(bold=True, size=16, name='宋体')
    summary['A1'].alignment = Alignment(horizontal='center', vertical='center')
    summary.merge_cells('A2:E2')
    summary['A2'] = '项目名称（全称）：科技文化中心—国际体育中心（足球场项目）'
    summary['A2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    headers = ['序号', '月份', '劳务分包公司名称', '总额（元）', '合计']
    for col_idx, header in enumerate(headers, start=1):
        summary.cell(row=3, column=col_idx, value=header)
    month_rows = {}
    detail_rows = []
    for idx, detail in enumerate(details, start=1):
        batch = detail['batch']
        row_idx = 3 + idx
        month = _period_number(batch['period'])
        summary.cell(row=row_idx, column=1, value=idx)
        summary.cell(row=row_idx, column=2, value=month)
        summary.cell(row=row_idx, column=3, value=_short_company(batch['company']))
        sheet_name = detail['sheet'].title.replace("'", "''")
        pay_letter = get_column_letter(detail['pay_col'])
        start_row = detail['start_row']
        end_row = start_row + len(batch['df']) - 1
        summary.cell(row=row_idx, column=4, value=f"=SUM('{sheet_name}'!{pay_letter}{start_row}:{pay_letter}{end_row})")
        detail_rows.append(row_idx)
        month_rows.setdefault(month, []).append(row_idx)
    for rows in month_rows.values():
        first = rows[0]
        summary.cell(row=first, column=5, value=f'=SUM(D{rows[0]}:D{rows[-1]})')
        for row_idx in rows[1:]:
            summary.cell(row=row_idx, column=2, value=None)
            summary.cell(row=row_idx, column=5, value=None)
    total_row = 4 + len(detail_rows)
    summary.cell(row=total_row, column=1, value='合计')
    if detail_rows:
        summary.cell(row=total_row, column=4, value=f'=SUM(D{detail_rows[0]}:D{detail_rows[-1]})')
        summary.cell(row=total_row, column=5, value=f'=SUM(E{detail_rows[0]}:E{detail_rows[-1]})')
    summary.column_dimensions['A'].width = 10
    summary.column_dimensions['B'].width = 10
    summary.column_dimensions['C'].width = 24
    summary.column_dimensions['D'].width = 16
    summary.column_dimensions['E'].width = 16
    for row_idx in range(1, total_row + 1):
        for col_idx in range(1, 6):
            summary.cell(row=row_idx, column=col_idx).border = _thin_border()
            summary.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = 'auto'
    except AttributeError:
        pass


def _workbook_bytes(wb):
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_salary_card_ledger_xlsx(salary_df):
    if not OPENPYXL_OK or salary_df is None or salary_df.empty:
        return b''
    # 确保主表档案全量贯通补齐
    salary_df = enrich_with_master(salary_df)
    valid_mask = (pd.to_numeric(salary_df.get('最终核定天数', 0), errors='coerce').fillna(0) > 0) | \
                 (pd.to_numeric(salary_df.get('应发工资', 0), errors='coerce').fillna(0) > 0)
    if valid_mask.any():
        salary_df = salary_df[valid_mask].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = '工资卡台账'
    
    ft_title = Font(name='宋体', size=16, bold=True)
    ft_header = Font(name='宋体', size=11, bold=True)
    ft_body = Font(name='宋体', size=11)
    ac = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bd = _thin_border()
    
    ws.row_dimensions[1].height = 30
    title_cell = ws.cell(row=1, column=1, value='农民工工资卡登记台账')
    title_cell.font = ft_title
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    for col in range(1, 10):
        ws.cell(row=1, column=col).border = bd
        
    headers = ['序号', '所属项目', '劳务单位', '姓名', '身份证号', '电话号码', '开户行', '卡号', '备注']
    ws.row_dimensions[2].height = 20
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = ft_header
        c.alignment = ac
        c.border = bd
        
    start_row = 3
    for idx, (_, row) in enumerate(salary_df.iterrows()):
        r = start_row + idx
        project = _excel_value(row, '项目全称', '项目简称') or '科技文化中心—国际体育中心（足球场项目）'
        company = _excel_value(row, '分包/所属企业', '劳务公司', '所属企业')
        name = _excel_value(row, '姓名')
        id_num = _excel_value(row, '身份证号')
        phone = _excel_value(row, '联系电话', '手机号', '电话', '手机号码')
        bank = _excel_value(row, '开户银行', '开户行')
        card = _excel_value(row, '银行卡号', '工资卡号', '卡号', '银行卡')
        
        vals = [idx + 1, project, company, name, id_num, phone, bank, card, '']
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=_excel_text(v))
            c.font = ft_body
            c.alignment = ac
            c.border = bd
            if ci in (5, 6, 8):
                c.number_format = '@'

    _set_col_widths(ws, [6, 25, 25, 12, 22, 15, 20, 22, 15])
    return _workbook_bytes(wb)



def build_all_exports_zip(salary_df, include_ledger=False):
    """将每个公司/月份批次生成包含公司标准考勤表与工资确认表的单一 Excel 文件，并打包为 ZIP。"""
    if not OPENPYXL_OK:
        return b''
    # 确保主表档案全量贯通补齐
    salary_df = enrich_with_master(salary_df)
    batches = _export_batches(salary_df)
    if not batches:
        return b''
    
    zip_buf = io.BytesIO()
    errors = []
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        if include_ledger:
            try:
                ledger_bytes = build_salary_card_ledger_xlsx(salary_df)
                if ledger_bytes:
                    zf.writestr('农民工工资卡登记台账.xlsx', ledger_bytes)
            except Exception as exc:
                errors.append(f'工资卡台账生成失败：{exc}')

        for batch in batches:
            c_name = _short_company(batch['company'], attendance=False)
            p_name = batch['period'] or "未知月份"
            
            company_str_gongsi = c_name
            period_str_gongsi = p_name

            # 统一过滤并对齐人员名单与顺序：只保留最终核定天数大于0的人员，确保两表完全一致
            batch_df = batch['df'].copy()
            valid_mask = pd.to_numeric(batch_df.get('最终核定天数', 0), errors='coerce').fillna(0) > 0
            batch_df = batch_df[valid_mask].copy().reset_index(drop=True)
            batch['df'] = batch_df

            try:
                # 1. 考勤表为主体加载，完整保留表头Logo图片与排版
                wb = load_workbook('load-data/muban/gongsi/附件7：务工人员（含队长、班组长、弄民工）考勤表.xlsx')
                att_source = _select_company_sheet(wb, batch['company'])
                for s in list(wb.worksheets):
                    if s != att_source:
                        wb.remove(s)
                att_ws = att_source
                att_ws.title = '考勤表'
                att_result = _populate_company_attendance_sheet(att_ws, batch)

                # 2. 工资确认表复制进同一 Workbook（1、工资表和考勤表放在一起，不要总表）
                wb_wage_temp = load_workbook('load-data/muban/gongsi/副本工资发放表.xlsx')
                wage_source = _select_company_sheet(wb_wage_temp, batch['company'])
                wage_ws = wb.create_sheet('工资确认表')
                _copy_worksheet_content(wage_source, wage_ws)
                _populate_company_wage_sheet(wage_ws, batch, att_footer_row=att_result['footer_row'], att_sheet_name='考勤表')

                # 设置自动重算
                try:
                    wb.calculation.fullCalcOnLoad = True
                    wb.calculation.forceFullCalc = True
                    wb.calculation.calcMode = 'auto'
                except AttributeError:
                    pass

                # 工资表和考勤表放在一起，命名为：公司-月份-考勤与工资确认表.xlsx
                filename = f'{company_str_gongsi}-{period_str_gongsi}-考勤与工资确认表.xlsx'
                zf.writestr(filename, _workbook_bytes(wb))
            except Exception as exc:
                errors.append(f'{c_name} 考勤与工资确认表生成失败：{exc}')

    for error in errors:
        st.error(f'导出失败：{error}')
    return zip_buf.getvalue()


# ================================================================
# 展示层
# ================================================================

def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="header-emoji">🗂️</span>
        <div class="header-text">
            <h2>劳务发薪全流程流水线</h2>
            <p>工时核算与额度反推 → 第二批工资待发储备池管理 → 多批次工资发放总台账</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)

    tab_calc, tab_backlog, tab_ledger = st.tabs([
        ':material/calculate: 考勤核定与额度反推',
        ':material/inventory_2: 第二批工资待发储备池',
        ':material/receipt_long: 工资发放总台账 (多批次动态累计)'
    ])

    from services.payroll_calculation_service import (
        filter_top_working_days,
        calculate_budget_allocation,
        add_to_backlog,
        get_backlog_for_company,
        extract_from_backlog,
        number_to_rmb_capital,
        parse_sanju_punch_excel,
        reconcile_attendance,
    )

    # ── Tab 1: 💰 考勤核定与额度反推 ──────────────────────────────────
    with tab_calc:
        # Step 1: 双表上传与解析
        with st.container(border=True):
            st.markdown("""
            <div class="step-indicator">
                <span class="step-num">1</span>
                <span>上传总包系统打卡流水表 & 我的考勤表</span>
            </div>
            """, unsafe_allow_html=True)
            
            c_p1, c_p2 = st.columns([1, 1])
            with c_p1:
                period_input = st.text_input("发薪月份/周期 (如: 8月 或 6、7月)", value=st.session_state.get('_payroll_period', '8月'), key='payroll_period_input')
            with c_p2:
                company_input = st.selectbox("默认分包劳务单位", ["江苏旭之升建筑工程有限公司", "青海久昌建筑装饰工程有限公司"], key='payroll_comp_select')

            c_u1, c_u2 = st.columns(2)
            with c_u1:
                file_system = st.file_uploader(
                    "1. 【必传】总包系统考勤打卡流水 / 月度出勤统计表",
                    type=['xlsx', 'xls', 'csv'],
                    key='payroll_file_system',
                    help="支持导出自总包系统/三局/护薪系统的出勤统计表（横向月度宽表或流水明细），系统考勤严格按【有效工天】自动核算"
                )
            with c_u2:
                file_my = st.file_uploader(
                    "2. 【推荐】我的考勤表 / 水印照片核对表",
                    type=['xlsx', 'xls', 'csv'],
                    key='payroll_file_my',
                    help="劳资员水印照片核对表（以水印照片为主，防止代打卡；若系统考勤小于水印照片，多出天数自动转为阶段性结算）"
                )

            if st.button(":material/compare_arrows: 开始三方智能对账与解析", type='primary', key='btn_parse_att'):
                if not file_system and not file_my:
                    st.warning(":material/info: 请至少上传一份考勤文件（推荐上传总包系统考勤打卡流水）")
                else:
                    with st.spinner("正在解析总包每日打卡工时（按有效工天）并匹配人员档案……"):
                        if file_system:
                            sys_df, detected_period = parse_sanju_punch_excel(file_system)
                            if detected_period and detected_period != '8月':
                                period_input = detected_period

                            my_parsed = None
                            if file_my:
                                df_my_raw = _safe_read(file_my, '我的考勤表')
                                my_parsed = parse_watermark_attendance(df_my_raw, company=company_input, period=period_input)

                            enriched = enrich_with_master(sys_df)
                            if '提取日薪' not in enriched.columns:
                                enriched['提取日薪'] = 300.0
                            enriched['提取日薪'] = pd.to_numeric(enriched['提取日薪'], errors='coerce').fillna(300.0)
                            if '津贴' not in enriched.columns:
                                enriched['津贴'] = 0.0
                            enriched['津贴'] = pd.to_numeric(enriched['津贴'], errors='coerce').fillna(0.0)

                            reconciled = reconcile_attendance(enriched, my_df=my_parsed)
                            reconciled['锁定'] = False

                            st.session_state['_payroll_working_df'] = reconciled
                            st.session_state['_payroll_period'] = period_input
                            st.session_state.pop('_payroll_allocated_df', None)
                            st.success(f"✅ 考勤解析成功！共识别到 {len(reconciled)} 名务工人员（已严格按有效工天核算）。")
                            st.rerun()
                        else:
                            # 仅上传了水印签到表
                            df_raw = _safe_read(file_my, '考勤打卡表')
                            parsed = parse_watermark_attendance(df_raw, company=company_input, period=period_input)
                            enriched = enrich_with_master(parsed)
                            if '提取日薪' not in enriched.columns:
                                enriched['提取日薪'] = 300.0
                            enriched['提取日薪'] = pd.to_numeric(enriched['提取日薪'], errors='coerce').fillna(300.0)
                            if '津贴' not in enriched.columns:
                                enriched['津贴'] = 0.0
                            enriched['津贴'] = pd.to_numeric(enriched['津贴'], errors='coerce').fillna(0.0)
                            enriched['总包系统天数'] = enriched['解析出勤天数']
                            enriched['有效工天'] = enriched['解析出勤天数']
                            enriched['出勤天数'] = enriched['解析出勤天数']
                            enriched['我的考勤天数'] = enriched['解析出勤天数']
                            enriched['水印照片天数'] = enriched['解析出勤天数']
                            enriched['班组长核定工日'] = enriched['解析出勤天数']
                            enriched['最终裁定工日'] = enriched['解析出勤天数']
                            enriched['最终核定天数'] = enriched['解析出勤天数']
                            enriched['正常打卡天数'] = enriched['解析出勤天数']
                            enriched['阶段性结算天数'] = 0.0
                            enriched['正常打卡工资'] = (enriched['正常打卡天数'] * enriched['提取日薪']).round(2)
                            enriched['阶段性结算金额'] = 0.0
                            enriched['裁定状态'] = '✔️ 单源水印打卡'
                            enriched['总工时'] = enriched['解析出勤天数'] * 8.0
                            enriched['平均工时'] = 8.0
                            enriched['_raw_punch_hours'] = [dict(r.get('__考勤明细__', {})) if isinstance(r.get('__考勤明细__'), dict) else {} for _, r in enriched.iterrows()]
                            enriched['应发工资'] = (enriched['最终核定天数'] * enriched['提取日薪'] + enriched['津贴']).round(2)
                            enriched['锁定'] = False

                            st.session_state['_payroll_working_df'] = enriched
                            st.session_state['_payroll_period'] = period_input
                            st.session_state.pop('_payroll_allocated_df', None)
                            st.success(f"✅ 解析成功！共识别到 {len(enriched)} 名务工人员。")
                            st.rerun()

        # Step 2: 考勤核对与阶段性结算工作台
        working_df = st.session_state.get('_payroll_working_df')
        if working_df is not None and not working_df.empty:
            st.markdown("---")
            with st.container(border=True):
                st.markdown("""
                <div class="step-indicator">
                    <span class="step-num">2</span>
                    <span>考勤智能核对与阶段性结算工作台（水印照片 VS 总包有效工天）</span>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("""
                <div class="hint-box" style="margin-bottom:12px;">
                💡 <b>考勤核对原则（极简两方比对：水印照片 VS 总包出勤统计表）</b>：<br>
                1. <b>真实出勤以水印照片为主</b>：防止有人没有上班也去打卡（防代打卡/防空刷）。<br>
                2. <b>总包系统考勤严格按【有效工天】算</b>：系统考勤天数直接取自总包出勤统计表中的“有效工天”。<br>
                3. <b>水印照片 > 系统打卡天数（出勤天数）</b>：以水印照片为主！在总包系统打卡记录容量内按有效工天走正常发放；<b>超出打卡记录的天数在系统中按【阶段性结算】做</b>，系统已为您特殊标明超出的天数与结算金额！<br>
                4. <b>水印照片 < 系统有效工天</b>：以水印照片为主核减！系统自动依据每日真实工时由长到短排序，<b>优先保留最长工时打 ✓，自动淘汰混打卡短工时</b>。
                </div>
                """, unsafe_allow_html=True)

                # 阶段性结算特别提醒横幅
                stage_settle_mask = (pd.to_numeric(working_df.get('阶段性结算天数', 0), errors='coerce').fillna(0) > 0)
                stage_settle_count = stage_settle_mask.sum()
                stage_settle_total_days = working_df.loc[stage_settle_mask, '阶段性结算天数'].sum() if stage_settle_count > 0 else 0.0
                stage_settle_total_amount = working_df.loc[stage_settle_mask, '阶段性结算金额'].sum() if stage_settle_count > 0 else 0.0

                if stage_settle_count > 0:
                    st.markdown(f"""
                    <div style="background:#FEF3C7; border:1px solid #F59E0B; border-radius:8px; padding:12px 16px; margin-bottom:14px; color:#92400E;">
                        <span style="font-size:16px; font-weight:bold;">🚨 阶段性结算特殊提醒</span><br>
                        共有 <b>{stage_settle_count}</b> 名工人<b>水印照片超出系统打卡天数</b>，累计多出 <b>{stage_settle_total_days:.1f}</b> 天，涉及金额 <b>¥{stage_settle_total_amount:,.2f}</b> 元。<br>
                        <span style="font-size:13px; color:#B45309;">※ 原因：总包系统考勤只能核减不能增加，多出的这几天在总包系统中完全没有打卡记录，无法做进正常考勤工资表，必须在总包系统中走【阶段性结算】做账补发。</span>
                    </div>
                    """, unsafe_allow_html=True)

                display_cols = [
                    '姓名', '分包/所属企业', '工种', '总工时', '平均工时',
                    '总包系统天数', '出勤天数', '水印照片天数', '阶段性结算天数', '正常打卡天数',
                    '最终核定天数', '裁定状态',
                    '提取日薪', '正常打卡工资', '阶段性结算金额', '津贴', '应发工资', '锁定'
                ]
                display_cols = [c for c in display_cols if c in working_df.columns]

                col_configs = {
                    '姓名': st.column_config.TextColumn('姓名', disabled=True),
                    '分包/所属企业': st.column_config.TextColumn('单位', disabled=True),
                    '工种': st.column_config.TextColumn('工种', disabled=True),
                    '总工时': st.column_config.NumberColumn('总工时(h)', disabled=True, format='%.1f'),
                    '平均工时': st.column_config.NumberColumn('日均工时(h)', disabled=True, format='%.1f'),
                    '总包系统天数': st.column_config.NumberColumn('系统有效工天', disabled=True, format='%.1f', help='按出勤统计表的有效工天核算'),
                    '出勤天数': st.column_config.NumberColumn('打卡天数', disabled=True, format='%.1f', help='总包打卡有记录的天数'),
                    '水印照片天数': st.column_config.NumberColumn('📸 水印照片天数', disabled=True, format='%.1f', help='现场水印照片真实核对天数，以此为主'),
                    '阶段性结算天数': st.column_config.NumberColumn('🚨 阶段性结算天数', disabled=True, format='%.1f', help='水印照片超出系统打卡天数的部分，无打卡流水记录，需在系统中走阶段性结算补发做账'),
                    '正常打卡天数': st.column_config.NumberColumn('正常打卡天数', disabled=True, format='%.1f', help='总包系统内走正常考勤发放的天数'),
                    '最终核定天数': st.column_config.NumberColumn(
                        '✏ 最终核定天数',
                        min_value=0.0, max_value=31.0, step=0.5, format='%.1f',
                        help='最终核定总出勤天数（默认直接以水印照片天数为主；若有现场特批断电故障抢工等可双击特批修改）'
                    ),
                    '裁定状态': st.column_config.TextColumn('裁定状态', disabled=True),
                    '提取日薪': st.column_config.NumberColumn('日薪标准', min_value=0.0, format='¥%.0f'),
                    '正常打卡工资': st.column_config.NumberColumn('正常打卡工资', disabled=True, format='¥%.2f', help='正常打卡天数 × 日薪'),
                    '阶段性结算金额': st.column_config.NumberColumn('⭐ 阶段性结算额', disabled=True, format='¥%.2f', help='阶段性结算天数 × 日薪'),
                    '津贴': st.column_config.NumberColumn('✏ 津贴/补助', min_value=0.0, step=50.0, format='¥%.2f', help='带班津贴、高温补贴、岗位补贴等'),
                    '应发工资': st.column_config.NumberColumn('应发工资', disabled=True, format='¥%.2f', help='正常打卡工资 + 阶段性结算金额 + 津贴'),
                    '锁定': st.column_config.CheckboxColumn('锁定额度 (不削减)', default=False),
                }

                col_flt1, col_flt2, col_flt3 = st.columns([2, 2, 1])
                with col_flt1:
                    filter_mode = st.radio(
                        "人员筛选视图",
                        ["全部人员", "仅看需走阶段性结算人员 (水印 > 打卡天数)", "仅看被核减人员 (水印 < 有效工天)"],
                        horizontal=True,
                        key='rad_filter_mode'
                    )
                with col_flt2:
                    hide_zero_att = st.checkbox("🔍 隐藏0考勤人员", value=True, key='chk_hide_zero_att')
                with col_flt3:
                    zero_att_count = sum(1 for _, r in working_df.iterrows() if float(r.get('最终核定天数', r.get('最终裁定工日', 0)) or 0) <= 0 and float(r.get('总包系统天数', 0) or 0) <= 0)
                    if zero_att_count > 0:
                        st.caption(f"包含 {zero_att_count} 名 0 考勤人员")

                editor_df = working_df.copy()
                if hide_zero_att:
                    v_mask = (pd.to_numeric(editor_df.get('最终核定天数', editor_df.get('最终裁定工日', 0)), errors='coerce').fillna(0) > 0) | \
                             (pd.to_numeric(editor_df.get('总包系统天数', 0), errors='coerce').fillna(0) > 0)
                    editor_df = editor_df[v_mask].copy()

                if "阶段性结算" in filter_mode:
                    editor_df = editor_df[pd.to_numeric(editor_df.get('阶段性结算天数', 0), errors='coerce').fillna(0) > 0].copy()
                elif "核减" in filter_mode:
                    editor_df = editor_df[
                        editor_df['裁定状态'].str.contains('核减', na=False)
                    ].copy()

                if editor_df.empty:
                    editor_df = working_df.copy()

                with st.form("working_df_form", border=False):
                    edited_working = st.data_editor(
                        editor_df[display_cols],
                        column_config=col_configs,
                        use_container_width=True,
                        num_rows='fixed',
                        hide_index=True,
                        key='working_data_editor'
                    )
                    submit_edits = st.form_submit_button("💾 保存核定工日与津贴并重新计算", type='primary')

                if submit_edits:
                    fin_map = {}
                    alw_map = {}
                    wage_map = {}
                    lock_map = {}
                    for _, r in edited_working.iterrows():
                        nm = str(r['姓名']).strip()
                        try:
                            fin_map[nm] = float(r.get('最终核定天数', r.get('最终裁定工日', 0.0)))
                        except (ValueError, TypeError):
                            pass
                        try:
                            alw_map[nm] = float(r.get('津贴', 0.0) or 0.0)
                        except (ValueError, TypeError):
                            pass
                        try:
                            wage_map[nm] = float(r.get('提取日薪', 300.0) or 300.0)
                        except (ValueError, TypeError):
                            pass
                        if '锁定' in r:
                            lock_map[nm] = bool(r['锁定'])

                    working_df_to_reconcile = working_df.copy()
                    for idx, row in working_df_to_reconcile.iterrows():
                        nm = str(row['姓名']).strip()
                        if nm in wage_map:
                            working_df_to_reconcile.at[idx, '提取日薪'] = wage_map[nm]
                        if nm in alw_map:
                            working_df_to_reconcile.at[idx, '津贴'] = alw_map[nm]
                        if nm in lock_map:
                            working_df_to_reconcile.at[idx, '锁定'] = lock_map[nm]

                    re_reconciled = reconcile_attendance(
                        working_df_to_reconcile,
                        final_days_map=fin_map,
                        allowance_map=alw_map
                    )
                    st.session_state['_payroll_working_df'] = re_reconciled
                    st.success("✅ 已重新保存核定数据，完成智能对账与阶段性结算计算！")
                    st.rerun()

                current_working_df = st.session_state.get('_payroll_working_df', working_df)

                total_workers = len(current_working_df)
                sys_total_days = current_working_df['总包系统天数'].sum() if '总包系统天数' in current_working_df.columns else 0.0
                watermark_total_days = current_working_df['水印照片天数'].sum() if '水印照片天数' in current_working_df.columns else 0.0
                final_total_days = current_working_df['最终核定天数'].sum() if '最终核定天数' in current_working_df.columns else 0.0
                stage_settle_total_days = current_working_df['阶段性结算天数'].sum() if '阶段性结算天数' in current_working_df.columns else 0.0
                total_allowance = current_working_df['津贴'].sum() if '津贴' in current_working_df.columns else 0.0
                orig_salary = ((current_working_df['最终核定天数'] * current_working_df['提取日薪']) + current_working_df.get('津贴', 0.0)).sum()

                st.markdown(f"""
                <div class="metric-container" style="margin-top:12px;">
                    <div class="metric-card"><div class="metric-title">核对总人数</div><div class="metric-value">{total_workers} 人</div></div>
                    <div class="metric-card"><div class="metric-title">系统有效工天合计</div><div class="metric-value">{sys_total_days:.1f} 天</div></div>
                    <div class="metric-card"><div class="metric-title">水印照片天数合计</div><div class="metric-value">{watermark_total_days:.1f} 天</div></div>
                    <div class="metric-card"><div class="metric-title">最终核定天数合计</div><div class="metric-value" style="color:#059669;">{final_total_days:.1f} 天</div></div>
                    <div class="metric-card"><div class="metric-title">阶段性结算天数合计</div><div class="metric-value" style="color:#D97706;">{stage_settle_total_days:.1f} 天</div></div>
                    <div class="metric-card"><div class="metric-title">全额应发总额</div><div class="metric-value" style="color:#2563EB;">¥{orig_salary:,.2f}</div></div>
                </div>
                """, unsafe_allow_html=True)

                # 展开查看工人工时明细与打钩日历抽屉
                with st.expander("🔍 展开查看工人工时明细与打钩日历抽屉"):
                    st.caption("您可以选择任意务工人员，查看其 1~31 号每天的具体打卡时长（小时）及系统依据工时由长到短的裁定结果。")
                    names_list = current_working_df['姓名'].tolist()
                    chosen_name = st.selectbox("选择要核查的务工人员", names_list, key='inspect_worker_name')

                    worker_row = current_working_df[current_working_df['姓名'] == chosen_name].iloc[0]
                    raw_punch_hours = worker_row.get('_raw_punch_hours')
                    if not isinstance(raw_punch_hours, dict):
                        raw_punch_hours = {}

                    selected_marks = worker_row.get('__考勤明细__')
                    if not isinstance(selected_marks, dict):
                        selected_marks = {}

                    rejected_foreman = worker_row.get('_rejected_foreman')
                    if not isinstance(rejected_foreman, dict):
                        rejected_foreman = {}

                    def _safe_float_val(v):
                        if v is None or v == '':
                            return 0.0
                        try:
                            return float(v)
                        except (ValueError, TypeError):
                            s = str(v).strip()
                            if '✓' in s or '√' in s or '打卡' in s or '出勤' in s:
                                return 8.0
                            return 0.0

                    days_calendar_rows = []
                    for d in range(1, 32):
                        h_num = 0.0
                        has_record = False
                        if d in raw_punch_hours:
                            h_num = _safe_float_val(raw_punch_hours[d])
                            has_record = True
                        elif d in rejected_foreman:
                            h_num = _safe_float_val(rejected_foreman[d])
                            has_record = True
                        elif d in selected_marks:
                            h_num = _safe_float_val(selected_marks[d])
                            has_record = True

                        if d in selected_marks:
                            mark = selected_marks[d]
                            if str(mark).strip() == '0.5':
                                status_tag = "✅ ✓ 选入考勤 (半天 0.5工日)"
                            else:
                                status_tag = "✅ ✓ 选入考勤 (有效工日)"
                        elif has_record and (h_num > 0 or d in rejected_foreman):
                            status_tag = "❌ 淘汰短工时 (混打卡已过滤)"
                        else:
                            status_tag = "— 无打卡"

                        hours_display = f"{h_num:.2f} h" if h_num > 0 else "-"
                        if h_num == 8.0 and d not in raw_punch_hours and d in selected_marks:
                            hours_display = "打卡 (标准工日)"

                        days_calendar_rows.append({
                            '日期': f"{d} 号",
                            '打卡工作时长(小时)': hours_display,
                            '裁定判定状态': status_tag
                        })

                    st.dataframe(pd.DataFrame(days_calendar_rows), use_container_width=True, hide_index=True)


            # Step 3: 资金额度反推与平账 (第二轮：比例反推与额度平齐)
            st.markdown("---")
            with st.container(border=True):
                st.markdown("""
                <div class="step-indicator">
                    <span class="step-num">3</span>
                    <span>资金额度反推与平账（第二轮：比例反推与截取相对短工时）</span>
                </div>
                """, unsafe_allow_html=True)

                col_b1, col_b2 = st.columns([2, 1])
                with col_b1:
                    target_budget = st.number_input(
                        "本次可用资金总额度 (元)",
                        min_value=0.0,
                        value=float(st.session_state.get('_target_budget', orig_salary)),
                        step=500.0,
                        key='input_target_budget'
                    )
                with col_b2:
                    strict_half = st.checkbox("严格规整为 0.5 天整倍数 (推荐工地实战)", value=True, key='check_strict_half')

                is_budget_sufficient = target_budget >= orig_salary - 0.01
                do_calc = False

                if is_budget_sufficient:
                    st.markdown(f"""
                    <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px; margin: 10px 0; color:#065F46;">
                        <b>🟢 资金额度充足（额度 ¥{target_budget:,.2f} ≥ 全额应发 ¥{orig_salary:,.2f}）</b><br>
                        当前资金完全可以覆盖所有工人工资，<b>无需进行工日反推缩减</b>。可直接点击下方按钮一键确认全额发放。
                    </div>
                    """, unsafe_allow_html=True)

                    c_btn_ok, c_btn_re = st.columns(2)
                    with c_btn_ok:
                        if st.button("💰 额度充足，确认全额发放（无需反推）", type='primary', key='btn_full_pay_confirm', use_container_width=True):
                            alloc_df = current_working_df.copy()
                            alloc_df['实发天数'] = alloc_df['最终核定天数']
                            alloc_df['实发工资'] = alloc_df['应发工资']
                            alloc_df['扣减天数'] = 0.0
                            alloc_df['欠发金额'] = 0.0
                            alloc_df['_rejected_details'] = [{} for _ in range(len(alloc_df))]
                            st.session_state['_payroll_allocated_df'] = alloc_df
                            st.session_state['_target_budget'] = target_budget
                            st.success("✅ 额度充足，已确认全员全额发放！无需扣减任何工日。")
                            st.rerun()
                    with c_btn_re:
                        do_calc = st.button(":material/tune: 仍按此额度强制重算", type='secondary', key='btn_calc_budget_forced', use_container_width=True)
                else:
                    do_calc = st.button(":material/tune: 执行额度反推与尾差调平", type='primary', key='btn_calc_budget', use_container_width=True)

                if do_calc:
                    with st.spinner("正在根据额度智能缩减天数并筛选次长工时……"):
                        st.session_state['_target_budget'] = target_budget
                        alloc_df = calculate_budget_allocation(
                            current_working_df,
                            total_budget=target_budget,
                            strict_half_days=strict_half
                        )

                        rejected_pool = {}
                        new_details_list = []
                        for _, row in alloc_df.iterrows():
                            wk_name = row['姓名']
                            raw_details = row.get('_raw_punch_hours')
                            if not isinstance(raw_details, dict) or not raw_details:
                                raw_details = row.get('__考勤明细__', {})
                                if not isinstance(raw_details, dict):
                                    raw_details = {}
                            actual_days = float(row['实发天数'])

                            selected_days, rejected_days = filter_top_working_days(raw_details, actual_days)
                            new_details_list.append(selected_days)
                            rejected_pool[wk_name] = rejected_days

                        alloc_df['__考勤明细__'] = new_details_list
                        alloc_df['扣减天数'] = (alloc_df['最终核定天数'] - alloc_df['实发天数']).round(2)
                        alloc_df['欠发金额'] = (alloc_df['应发工资'] - alloc_df['实发工资']).round(2)
                        alloc_df['_rejected_details'] = [rejected_pool.get(r['姓名'], {}) for _, r in alloc_df.iterrows()]

                        st.session_state['_payroll_allocated_df'] = alloc_df
                        st.success("✅ 额度反推分配完成！实发天数与金额已严格对齐。")
                        st.rerun()

                alloc_df = st.session_state.get('_payroll_allocated_df')
                if alloc_df is not None and not alloc_df.empty:
                    act_total = alloc_df['实发工资'].sum()
                    rem_diff = round(target_budget - act_total, 2)
                    deducted_count = (alloc_df['扣减天数'] > 0).sum()
                    deducted_amount = alloc_df['欠发金额'].sum()

                    st.markdown(f"""
                    <div class="metric-container" style="margin-top:12px;">
                        <div class="metric-card"><div class="metric-title">目标控制额度</div><div class="metric-value">¥{target_budget:,.2f}</div></div>
                        <div class="metric-card"><div class="metric-title">本次实发总额</div><div class="metric-value" style="color:#059669;">¥{act_total:,.2f}</div></div>
                        <div class="metric-card"><div class="metric-title">额度结余/微差</div><div class="metric-value">¥{rem_diff:,.2f}</div></div>
                        <div class="metric-card"><div class="metric-title">涉及延发人数</div><div class="metric-value" style="color:#dc2626;">{deducted_count} 人 (欠¥{deducted_amount:,.0f})</div></div>
                    </div>
                    """, unsafe_allow_html=True)

                    preview_cols = ['姓名', '分包/所属企业', '工种', '最终核定天数', '实发天数', '提取日薪', '津贴', '应发工资', '实发工资', '扣减天数', '欠发金额']
                    preview_cols = [c for c in preview_cols if c in alloc_df.columns]
                    st.dataframe(alloc_df[preview_cols], use_container_width=True, hide_index=True)

                    col_act1, col_act2 = st.columns(2)
                    with col_act1:
                        if st.button(":material/save: 将延发差额存入第二批储备池", type='secondary', use_container_width=True):
                            added_count = 0
                            for _, r in alloc_df[alloc_df['扣减天数'] > 0].iterrows():
                                add_to_backlog(
                                    period=st.session_state.get('_payroll_period', '当期'),
                                    company=r.get('分包/所属企业', company_input),
                                    worker_id=str(r.get('身份证号', '')),
                                    worker_name=r['姓名'],
                                    days_deducted=float(r['扣减天数']),
                                    amount_deducted=float(r['欠发金额']),
                                    rejected_days=r.get('_rejected_details', {}),
                                    daily_rate=float(r.get('提取日薪', 0))
                                )
                                added_count += 1
                            st.success(f"🎉 成功将 {added_count} 名工人的欠发工时与明细记录至【第二批工资储备池】！")

                    with col_act2:
                        export_df = alloc_df.copy()
                        export_df['最终核定天数'] = export_df['实发天数']
                        export_df['应发工资'] = export_df['实发工资']
                        export_df['日薪'] = export_df['提取日薪']
                        zip_bytes = build_all_exports_zip(export_df, include_ledger=False)
                        if zip_bytes:
                            st.download_button(
                                label=":material/archive: 一键下载标准考勤表与工资确认表 (ZIP)",
                                data=zip_bytes,
                                file_name=f"劳务考勤与工资确认表_{st.session_state.get('_payroll_period', '当期')}.zip",
                                mime="application/zip",
                                use_container_width=True,
                                type='primary'
                            )

    # ── Tab 2: 📋 第二批工资待发储备池 ─────────────────────────────────
    with tab_backlog:
        st.subheader("第二批工资待发储备池（欠薪工时精准结转）")
        st.caption("记录因资金额度受限扣减的务工人员工时与具体打卡日期，当下批额度充裕时一键调取生成补发表。")

        b_comp = st.selectbox("选择分包劳务单位", ["江苏旭之升建筑工程有限公司", "青海久昌建筑装饰工程有限公司"], key='select_backlog_comp')
        backlog_records = get_backlog_for_company(b_comp)

        if not backlog_records:
            st.info(f"✨ 当前【{b_comp}】暂无待发留存差额，所有人员工时均已全额结清！")
        else:
            for p_key, p_workers in backlog_records.items():
                with st.expander(f"📁 归档月份/周期：{p_key} (共 {len(p_workers)} 人待发)", expanded=True):
                    b_rows = []
                    for w_id, w_info in p_workers.items():
                        detail_str = "、".join([f"{d}号({h}h)" for d, h in w_info.get('扣减具体日期', {}).items()])
                        b_rows.append({
                            '序号': len(b_rows) + 1,
                            '姓名': w_info.get('姓名', ''),
                            '身份证号': w_info.get('身份证号', ''),
                            '日薪标准': w_info.get('日薪', 0.0),
                            '欠发天数': w_info.get('扣减天数', 0.0),
                            '欠发金额': w_info.get('欠额', 0.0),
                            '扣减具体日期明细': detail_str if detail_str else '工时比例折算',
                            '_key': w_id
                        })
                    b_df = pd.DataFrame(b_rows)
                    st.dataframe(b_df.drop(columns=['_key']), use_container_width=True, hide_index=True)

                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        wb_b2 = Workbook()
                        ws_b2 = wb_b2.active
                        ws_b2.title = f"{p_key}第二批补发工资表"
                        ws_b2.append(["序号", "姓名", "身份证号", "日薪", "补发天数", "补发金额", "对应扣减考勤日期"])
                        for r_item in b_rows:
                            ws_b2.append([
                                r_item['序号'], r_item['姓名'], r_item['身份证号'],
                                r_item['日薪标准'], r_item['欠发天数'], r_item['欠发金额'],
                                r_item['扣减具体日期明细']
                            ])
                        b2_buf = io.BytesIO()
                        wb_b2.save(b2_buf)
                        st.download_button(
                            label=f":material/download: 导出【{p_key}】第二批补发工资表 (Excel)",
                            data=b2_buf.getvalue(),
                            file_name=f"{b_comp}_{p_key}_第二批补发工资表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_b2_{p_key}",
                            use_container_width=True
                        )
                    with c_btn2:
                        if st.button(f":material/check_circle: 确认补发并核销【{p_key}】储备", key=f"clear_b2_{p_key}", use_container_width=True):
                            extract_from_backlog(b_comp, p_key, [r['_key'] for r in b_rows])
                            st.success(f"已核销【{p_key}】所有人员的第二批待发差额！")
                            st.rerun()

    # ── Tab 3: 📊 工资发放总台账 (多批次动态累计) ────────────────────────
    with tab_ledger:
        from modules.payroll_ledger_tab import render_payroll_ledger_tab
        render_payroll_ledger_tab()

