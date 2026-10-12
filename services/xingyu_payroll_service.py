# -*- coding: utf-8 -*-
"""
services/xingyu_payroll_service.py
上海兴豫建筑工程有限公司工资台账专用生成与填充服务。
符合总包/官方规范，包含：
1. 大标题与双层复合表头（单列垂直合并、月度与批次横向合并）
2. 基础人员列（序号、姓名、身份证号码、手机号码、进场时间）
3. 月度应发工资、月度已发工资明细
4. 10个具体批次流水（应发、已发、未发、班组长代付、约付时间）
5. 班组长代付月度明细
6. 年度汇总与退场状态（应发合计、已发合计、代付、未发工资、退场时间等）
7. 动态 Excel 公式联动（BS=SUM应发, BT=SUM已发, BV=BS-BT未发，底部整列SUM合计）
8. 冻结 F4 窗格，支持超大屏幕水平滚动而不丢失工人姓名
"""

import os
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def build_xingyu_payroll_workbook(
    company_name: str = "上海兴豫建筑工程有限公司",
    year: int = 2026,
    custom_rows: list = None,
    total_placeholder_rows: int = 78
) -> openpyxl.Workbook:
    """
    生成《公司工资台账》专业 Workbook
    :param company_name: 公司名称，用于大标题
    :param year: 年度，默认 2026
    :param custom_rows: 自定义填充数据，若为空则默认使用杨成示例并预留 total_placeholder_rows 行
    :param total_placeholder_rows: 预留数据行数，默认 78
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工资台账"
    ws.views.sheetView[0].showGridLines = True

    # 字体
    font_main_title = Font(name="微软雅黑", size=16, bold=True, color="1F4E79")
    font_h1 = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    font_h2 = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="微软雅黑", size=10, color="000000")
    font_total = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")

    # 填充色
    fill_title = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_h1_base = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_h1_plan = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_h1_paid = PatternFill(start_color="165B33", end_color="165B33", fill_type="solid")
    fill_h1_batch1 = PatternFill(start_color="41719C", end_color="41719C", fill_type="solid")
    fill_h1_batch2 = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    fill_h1_agent = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
    fill_h1_summary = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_h2 = PatternFill(start_color="333F48", end_color="333F48", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")

    # 边框
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    double_border_bottom = Side(border_style="double", color="1F4E79")
    thick_border_top = Side(border_style="thin", color="1F4E79")

    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thick_border_top, bottom=double_border_bottom)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    months = [f"{year}-01-01", f"{year}-02-01", f"{year}-03-01", f"{year}-04-01", f"{year}-05-01"]
    batch_dates = [
        f"{year}-03-06", f"{year}-04-07", f"{year}-05-09", f"{year}-06-10", f"{year}-07-12",
        f"{year}-08-13", f"{year}-09-14", f"{year}-10-16", f"{year}-11-17", f"{year}-12-19"
    ]
    batch_sub_headers = ["应发工资（元）", "已发工资（元）", "未发工资（元）", "分包班组长自行代付工资", "未付工资/约定支付时间"]

    columns_spec = []

    # 1. 基础信息列
    base_cols = [
        ("序号", "A", 8, "center"),
        ("姓名", "B", 12, "center"),
        ("身份证号码", "C", 22, "center"),
        ("手机号码", "D", 15, "center"),
        ("进场时间", "E", 14, "center"),
    ]
    for name, let, w, al in base_cols:
        columns_spec.append({
            "name": name,
            "h1": name,
            "h2": "",
            "is_merged_v": True,
            "fill": fill_h1_base,
            "width": w,
            "align": al,
            "type": "text"
        })

    # 2. 应发工资（元）月度明细 (F~J)
    for m in months:
        columns_spec.append({
            "name": f"应发_{m}",
            "h1": "应发工资（元）",
            "h2": m,
            "group": "应发工资（元）",
            "is_merged_v": False,
            "fill": fill_h1_plan,
            "width": 13,
            "align": "right",
            "type": "currency"
        })

    # 3. 已发工资（元）月度明细 (K~O)
    for m in months:
        columns_spec.append({
            "name": f"已发_{m}",
            "h1": "已发工资（元）",
            "h2": m,
            "group": "已发工资（元）",
            "is_merged_v": False,
            "fill": fill_h1_paid,
            "width": 13,
            "align": "right",
            "type": "currency"
        })

    # 4. 10个批次发薪流水 (P~BM)
    for b_idx, b_date in enumerate(batch_dates):
        b_fill = fill_h1_batch1 if b_idx % 2 == 0 else fill_h1_batch2
        for s_idx, sub_name in enumerate(batch_sub_headers):
            columns_spec.append({
                "name": f"{b_date}_{sub_name}",
                "h1": b_date,
                "h2": sub_name,
                "group": b_date,
                "is_merged_v": False,
                "fill": b_fill,
                "width": 16 if "时间" in sub_name or "代付" in sub_name else 14,
                "align": "center" if "时间" in sub_name else "right",
                "type": "text" if "时间" in sub_name else "currency"
            })

    # 5. 其中：班组长代付（元） (BN~BR)
    for m in months:
        columns_spec.append({
            "name": f"代付_{m}",
            "h1": "其中：班组长代付（元）",
            "h2": m,
            "group": "其中：班组长代付（元）",
            "is_merged_v": False,
            "fill": fill_h1_agent,
            "width": 13,
            "align": "right",
            "type": "currency"
        })

    # 6. 年度汇总与状态 (BS~BZ)
    summary_cols = [
        (f"{year}年应发工资（元）", 16, "right", "currency"),
        (f"{year}年已发工资（元）", 16, "right", "currency"),
        (f"{year}年分包班组长自行代付工资", 18, "right", "currency"),
        (f"{year}年未发工资（元）", 16, "right", "currency"),
        (f"{year}年未付工资/约定支付时间", 18, "center", "text"),
        ("人员退场时间", 14, "center", "text"),
        (f"{year}年未付工资\n约定支付时间", 16, "center", "text"),
        ("备注", 16, "left", "text")
    ]
    for s_name, w, al, tp in summary_cols:
        columns_spec.append({
            "name": s_name,
            "h1": s_name,
            "h2": "",
            "is_merged_v": True,
            "fill": fill_h1_summary,
            "width": w,
            "align": al,
            "type": tp
        })

    total_cols = len(columns_spec)
    last_col_letter = get_column_letter(total_cols)

    # ── 第 1 行：大标题 ──
    ws.row_dimensions[1].height = 42
    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws["A1"]
    title_cell.value = f"{company_name}工资台账"
    title_cell.font = font_main_title
    title_cell.alignment = align_center
    title_cell.fill = fill_title

    # ── 第 2 行 & 第 3 行：复合表头 ──
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 28

    for col_idx, col in enumerate(columns_spec, 1):
        c_letter = get_column_letter(col_idx)
        cell_h1 = ws.cell(row=2, column=col_idx)
        cell_h2 = ws.cell(row=3, column=col_idx)

        cell_h1.value = col["h1"]
        cell_h1.font = font_h1
        cell_h1.alignment = align_center
        cell_h1.fill = col["fill"]
        cell_h1.border = cell_border

        cell_h2.value = col["h2"]
        cell_h2.font = font_h2
        cell_h2.alignment = align_center
        cell_h2.fill = fill_h2
        cell_h2.border = cell_border

        ws.column_dimensions[c_letter].width = col["width"]

    # 垂直合并
    for col_idx, col in enumerate(columns_spec, 1):
        if col["is_merged_v"]:
            c_letter = get_column_letter(col_idx)
            ws.merge_cells(f"{c_letter}2:{c_letter}3")

    # 水平合并
    group_spans = []
    current_group = None
    start_c = None
    for col_idx, col in enumerate(columns_spec, 1):
        grp = col.get("group")
        if grp:
            if grp != current_group:
                if current_group and start_c:
                    group_spans.append((current_group, start_c, col_idx - 1))
                current_group = grp
                start_c = col_idx
        else:
            if current_group and start_c:
                group_spans.append((current_group, start_c, col_idx - 1))
                current_group = None
                start_c = None
    if current_group and start_c:
        group_spans.append((current_group, start_c, total_cols))

    for grp_name, s_col, e_col in group_spans:
        s_let = get_column_letter(s_col)
        e_let = get_column_letter(e_col)
        if s_col < e_col:
            ws.merge_cells(f"{s_let}2:{e_let}2")

    # ── 数据行填充 ──
    start_row = 4
    data_rows = custom_rows if custom_rows else []
    num_rows = max(len(data_rows), total_placeholder_rows)

    for idx in range(num_rows):
        curr_row = start_row + idx
        seq = idx + 1
        ws.row_dimensions[curr_row].height = 22
        r_data = data_rows[idx] if idx < len(data_rows) else None

        for col_idx, col in enumerate(columns_spec, 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = font_data
            cell.border = cell_border

            if seq % 2 == 0:
                cell.fill = fill_zebra

            col_type = col["type"]
            col_align = col["align"]
            if col_align == "center":
                cell.alignment = align_center
            elif col_align == "right":
                cell.alignment = align_right
            else:
                cell.alignment = align_left

            # 序号
            if col["name"] == "序号":
                cell.value = seq
                cell.number_format = "0"
                continue

            if col["name"] in ("身份证号码", "手机号码"):
                cell.number_format = "@"

            if col_type == "currency":
                cell.number_format = "#,##0.00"

            # 动态公式列
            if "应发工资（元）" in col["name"] and "年" in col["name"]:
                cell.value = f"=SUM(F{curr_row}:J{curr_row})"
                cell.number_format = "#,##0.00"
                continue

            if "已发工资（元）" in col["name"] and "年" in col["name"]:
                cell.value = f"=SUM(K{curr_row}:O{curr_row})"
                cell.number_format = "#,##0.00"
                continue

            if "未发工资（元）" in col["name"] and "年" in col["name"]:
                cell.value = f"=BS{curr_row}-BT{curr_row}"
                cell.number_format = "#,##0.00"
                continue

            # 若有自定义数据则填入
            if r_data:
                val = r_data.get(col["name"])
                if val is not None:
                    cell.value = val
            elif idx == 0:
                # 默认第一行保留杨成示例
                if col["name"] == "姓名":
                    cell.value = "杨成"
                elif col["name"] == "进场时间":
                    cell.value = "2025-02-19"
                elif col["name"] == f"应发_{year}-01-01":
                    cell.value = 5000.0
                elif col["name"] == f"应发_{year}-02-01":
                    cell.value = 0.0
                elif col["name"] == f"应发_{year}-03-01":
                    cell.value = 5000.0
                elif col["name"] == f"应发_{year}-04-01":
                    cell.value = 0.0
                elif col["name"] == f"应发_{year}-05-01":
                    cell.value = 5000.0
                elif col["name"] == f"已发_{year}-01-01":
                    cell.value = 10000.0
                elif col["name"] == f"已发_{year}-02-01":
                    cell.value = 4000.0
                elif col["name"] == f"已发_{year}-03-01":
                    cell.value = 6000.0
                elif col["name"] == f"已发_{year}-04-01":
                    cell.value = None
                elif col["name"] == f"已发_{year}-05-01":
                    cell.value = 6000.0
                elif col["name"] == f"代付_{year}-01-01":
                    cell.value = 10000.0
                elif col["name"] == f"代付_{year}-02-01":
                    cell.value = 4000.0
                elif col["name"] == f"代付_{year}-03-01":
                    cell.value = 6000.0
                elif col["name"] == f"代付_{year}-04-01":
                    cell.value = None
                elif col["name"] == f"代付_{year}-05-01":
                    cell.value = 6000.0
                elif col["name"] == f"{year}年分包班组长自行代付工资":
                    cell.value = 4000.0
                elif col["name"] == "人员退场时间":
                    cell.value = "2025-02-20"
                elif "未发工资" in col["name"] or "未付工资" in col["name"]:
                    cell.value = 0.0

    # ── 底部合计行 ──
    total_row = start_row + num_rows
    ws.row_dimensions[total_row].height = 28

    ws.merge_cells(f"A{total_row}:E{total_row}")
    cell_tot_label = ws[f"A{total_row}"]
    cell_tot_label.value = "合  计"
    cell_tot_label.font = font_total
    cell_tot_label.alignment = align_center
    cell_tot_label.fill = fill_total

    for c in range(1, 6):
        ws.cell(row=total_row, column=c).border = total_border
        ws.cell(row=total_row, column=c).fill = fill_total

    for col_idx in range(6, total_cols + 1):
        col = columns_spec[col_idx - 1]
        c_letter = get_column_letter(col_idx)
        cell = ws.cell(row=total_row, column=col_idx)
        cell.font = font_total
        cell.border = total_border
        cell.fill = fill_total

        if col["type"] == "currency":
            cell.value = f"=SUM({c_letter}{start_row}:{c_letter}{total_row - 1})"
            cell.number_format = "#,##0.00"
            cell.alignment = align_right
        else:
            cell.alignment = align_center

    # 冻结窗格
    ws.freeze_panes = "F4"

    return wb


def export_xingyu_excel(output_path: str = None, company_name: str = "上海兴豫建筑工程有限公司") -> str:
    """导出 Excel 文件到目标路径"""
    if not output_path:
        output_path = os.path.join(PROJECT_ROOT, f"{company_name}工资台账.xlsx")
    wb = build_xingyu_payroll_workbook(company_name=company_name)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path


def get_xingyu_excel_bytes(company_name: str = "上海兴豫建筑工程有限公司") -> bytes:
    """返回 Excel 内存二进制流，供 Streamlit 下载"""
    wb = build_xingyu_payroll_workbook(company_name=company_name)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
