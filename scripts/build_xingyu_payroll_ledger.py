# -*- coding: utf-8 -*-
"""
scripts/build_xingyu_payroll_ledger.py
根据用户模板规范生成《上海兴豫建筑工程有限公司工资台账.xlsx》
包含：
1. 完整两级表头与单元格合并（A1~BZ1 大标题，A2~BZ3 双层复合表头）
2. 基础人员信息列（序号、姓名、身份证号、手机号、进场时间）
3. 月度应发工资（5个月）、月度已发工资（5个月）
4. 10个发薪批次日期流水（每个批次5列）
5. 班组长代付月度明细（5个月）
6. 2026年度汇总列（应发合计、已发合计、代付合计、未发工资、约定支付时间、退场时间、备注）
7. 预置杨成示范行 + 2~78 序号预留行 + 底部全盘动态 SUM 合计行
8. 冻结前3行与前5列（F4冻结），便于横向与纵向对照查阅
9. 会计千分位格式与文本格式（身份证不截断、不转科学计数法）
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def build_xingyu_payroll_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工资台账"

    # 网格线可见
    ws.views.sheetView[0].showGridLines = True

    # 样式定义
    font_main_title = Font(name="微软雅黑", size=16, bold=True, color="1F4E79")
    font_h1 = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    font_h2 = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="微软雅黑", size=10, color="000000")
    font_total = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")
    font_sample = Font(name="微软雅黑", size=10, color="1E293B")

    # 填充色定义（专业工程风分组色）
    fill_title = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_h1_base = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")     # 深海军蓝（基础信息）
    fill_h1_plan = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")     # 稳重蓝（应发）
    fill_h1_paid = PatternFill(start_color="165B33", end_color="165B33", fill_type="solid")     # 稳重森林绿（已发）
    fill_h1_batch1 = PatternFill(start_color="41719C", end_color="41719C", fill_type="solid")   # 钢青蓝（流水奇数批次）
    fill_h1_batch2 = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")   # 天蓝色（流水偶数批次）
    fill_h1_agent = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")    # 暖棕橙（代付）
    fill_h1_summary = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")  # 藏青蓝（年度汇总）
    fill_h2 = PatternFill(start_color="333F48", end_color="333F48", fill_type="solid")          # 子表头深灰蓝
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")       # 合计行浅蓝
    fill_zebra = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")

    # 边框定义
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    double_border_bottom = Side(border_style="double", color="1F4E79")
    thick_border_top = Side(border_style="thin", color="1F4E79")

    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thick_border_top, bottom=double_border_bottom)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # 78 列定义
    # 0~4: 基础信息
    # 5~9: 应发工资（5个月）
    # 10~14: 已发工资（5个月）
    # 15~64: 10个批次流水（每个5列）
    # 65~69: 其中：班组长代付（5个月）
    # 70~77: 年度汇总及状态
    months = ["2026-01-01", "2026-02-01", "2026-03-01", "2026-04-01", "2026-05-01"]
    batch_dates = [
        "2026-03-06", "2026-04-07", "2026-05-09", "2026-06-10", "2026-07-12",
        "2026-08-13", "2026-09-14", "2026-10-16", "2026-11-17", "2026-12-19"
    ]
    batch_sub_headers = ["应发工资（元）", "已发工资（元）", "未发工资（元）", "分包班组长自行代付工资", "未付工资/约定支付时间"]

    # 构造列定义列表：[(col_idx, h1_title, h2_title, h1_group_name, fill_h1, col_type)]
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
        ("2026年应发工资（元）", 16, "right", "currency"),
        ("2026年已发工资（元）", 16, "right", "currency"),
        ("2026年分包班组长自行代付工资", 18, "right", "currency"),
        ("2026年未发工资（元）", 16, "right", "currency"),
        ("2026年未付工资/约定支付时间", 18, "center", "text"),
        ("人员退场时间", 14, "center", "text"),
        ("2026年未付工资\n约定支付时间", 16, "center", "text"),
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
    title_cell.value = "上海兴豫建筑工程有限公司工资台账"
    title_cell.font = font_main_title
    title_cell.alignment = align_center
    title_cell.fill = fill_title

    # ── 第 2 行 & 第 3 行：复合表头 ──
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 28

    # 填入表头文字与样式
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

        # 设置列宽
        ws.column_dimensions[c_letter].width = col["width"]

    # 合并表头单元格
    # 1. 垂直合并（单列垂直跨行2~3）
    for col_idx, col in enumerate(columns_spec, 1):
        if col["is_merged_v"]:
            c_letter = get_column_letter(col_idx)
            ws.merge_cells(f"{c_letter}2:{c_letter}3")

    # 2. 水平合并分组（应发工资 F2:J2, 已发工资 K2:O2, 各批次, 班组长代付 BN2:BR2）
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

    # ── 数据行填充（第 4 行开始） ──
    # 用户提供了 78 个序号的数据（包含第 1 行杨成的真实样板，以及 2~78 序号的预留行）
    # 杨成的样板数据
    # Col F~J: 5000, 0, 5000, 0, 5000
    # Col K~O: 10000, 4000, 6000, None, 6000
    # Col BN~BR: 10000, 4000, 6000, None, 6000
    # Col BS (应发合计) = SUM(F4:J4) = 15000
    # Col BT (已发合计) = SUM(K4:O4) = 26000
    # Col BU (班组长代付) = 4000
    # Col BV (未发工资) = BS4 - BT4 = -11000
    # Col BW = ""
    # Col BX (退场时间) = "2025-02-20"
    
    start_row = 4
    total_data_rows = 78

    for seq in range(1, total_data_rows + 1):
        curr_row = start_row + seq - 1
        ws.row_dimensions[curr_row].height = 22
        is_sample_row = (seq == 1)

        # 遍历每一列填充单元格
        for col_idx, col in enumerate(columns_spec, 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = font_data
            cell.border = cell_border

            # 斑马纹微底色
            if seq % 2 == 0:
                cell.fill = fill_zebra

            # 对齐与格式
            col_type = col["type"]
            col_align = col["align"]
            if col_align == "center":
                cell.alignment = align_center
            elif col_align == "right":
                cell.alignment = align_right
            else:
                cell.alignment = align_left

            # 序号列
            if col["name"] == "序号":
                cell.value = seq
                cell.number_format = "0"
                continue

            # 基础文本格式
            if col["name"] in ("身份证号码", "手机号码"):
                cell.number_format = "@"

            # 金额列默认数字格式
            if col_type == "currency":
                cell.number_format = "#,##0.00"

            # 动态公式列
            # 2026年应发工资（元）(BS)
            if col["name"] == "2026年应发工资（元）":
                # F ~ J 列之和
                cell.value = f"=SUM(F{curr_row}:J{curr_row})"
                cell.number_format = "#,##0.00"
                continue

            # 2026年已发工资（元）(BT)
            if col["name"] == "2026年已发工资（元）":
                # K ~ O 列之和
                cell.value = f"=SUM(K{curr_row}:O{curr_row})"
                cell.number_format = "#,##0.00"
                continue

            # 2026年未发工资（元）(BV)
            if col["name"] == "2026年未发工资（元）":
                # 应发合计 - 已发合计 = BS - BT
                cell.value = f"=BS{curr_row}-BT{curr_row}"
                cell.number_format = "#,##0.00"
                continue

            # 样本杨成数据
            if is_sample_row:
                if col["name"] == "姓名":
                    cell.value = "杨成"
                elif col["name"] == "进场时间":
                    cell.value = "2025-02-19"
                elif col["name"] == "应发_2026-01-01":
                    cell.value = 5000.0
                elif col["name"] == "应发_2026-02-01":
                    cell.value = 0.0
                elif col["name"] == "应发_2026-03-01":
                    cell.value = 5000.0
                elif col["name"] == "应发_2026-04-01":
                    cell.value = 0.0
                elif col["name"] == "应发_2026-05-01":
                    cell.value = 5000.0
                elif col["name"] == "已发_2026-01-01":
                    cell.value = 10000.0
                elif col["name"] == "已发_2026-02-01":
                    cell.value = 4000.0
                elif col["name"] == "已发_2026-03-01":
                    cell.value = 6000.0
                elif col["name"] == "已发_2026-04-01":
                    cell.value = None
                elif col["name"] == "已发_2026-05-01":
                    cell.value = 6000.0
                elif col["name"] == "代付_2026-01-01":
                    cell.value = 10000.0
                elif col["name"] == "代付_2026-02-01":
                    cell.value = 4000.0
                elif col["name"] == "代付_2026-03-01":
                    cell.value = 6000.0
                elif col["name"] == "代付_2026-04-01":
                    cell.value = None
                elif col["name"] == "代付_2026-05-01":
                    cell.value = 6000.0
                elif col["name"] == "2026年分包班组长自行代付工资":
                    cell.value = 4000.0
                elif col["name"] == "人员退场时间":
                    cell.value = "2025-02-20"
                elif "未发工资（元）" in col["name"] or "未付工资" in col["name"]:
                    # 流水中的未发工资/约定时间在模板中设为 0
                    cell.value = 0.0

    # ── 底部【合计行】（第 82 行） ──
    total_row = start_row + total_data_rows
    ws.row_dimensions[total_row].height = 28

    # A~E 列合并为“合计”
    ws.merge_cells(f"A{total_row}:E{total_row}")
    cell_tot_label = ws[f"A{total_row}"]
    cell_tot_label.value = "合  计"
    cell_tot_label.font = font_total
    cell_tot_label.alignment = align_center
    cell_tot_label.fill = fill_total

    for c in range(1, 6):
        ws.cell(row=total_row, column=c).border = total_border
        ws.cell(row=total_row, column=c).fill = fill_total

    # 金额列填充动态求和公式 =SUM(X4:X81)
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

    # ── 冻结窗格：固定前 3 行表头和前 5 列基础信息（F4 冻结） ──
    ws.freeze_panes = "F4"

    return wb


def export_xingyu_template_files():
    wb = build_xingyu_payroll_workbook()

    # 目标输出路径 1：项目根目录《上海兴豫建筑工程有限公司工资台账.xlsx》
    path1 = os.path.join(PROJECT_ROOT, "上海兴豫建筑工程有限公司工资台账.xlsx")
    wb.save(path1)
    print(f"成功保存至根目录: {path1}")

    # 目标输出路径 2：load-data/muban 模板目录
    muban_dir = os.path.join(PROJECT_ROOT, "load-data", "muban")
    os.makedirs(muban_dir, exist_ok=True)
    path2 = os.path.join(muban_dir, "上海兴豫建筑工程有限公司工资台账_标准模板.xlsx")
    wb.save(path2)
    print(f"成功保存至模板库: {path2}")

    return path1, path2


if __name__ == "__main__":
    p1, p2 = export_xingyu_template_files()
    print("Done!")
