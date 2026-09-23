"""
导出项目主表中退场人员的标准Excel报表
=======================================
功能：
1. 读取 data/master_state.json 主表全量数据
2. 筛选退场人员（当前状态为已退场/离场，共27人）以及含历史退场记录人员（共37人）
3. 按用户指定27个字段严格对齐输出
4. 设置规范的表头样式、列宽自适应、长数字（身份证、银行卡、电话等）强制为纯文本格式 @
"""

import json
from pathlib import Path
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
MASTER_STATE_FILE = BASE_DIR / "data" / "master_state.json"
OUTPUT_FILE = BASE_DIR / "退场人员信息表.xlsx"

COLUMNS = [
    "工人姓名",
    "性别",
    "出生日期",
    "年龄",
    "籍贯",
    "民族",
    "政治面貌",
    "工种",
    "家庭住址",
    "身份证号码",
    "发证机关",
    "有效起始日期",
    "有效截止日期",
    "开户银行",
    "开户支行",
    "工资卡号",
    "单位工资",
    "劳动合同登记",
    "劳动合同编号",
    "安全培训情况",
    "进场教育",
    "首次进场时间",
    "进场时间",
    "退场时间",
    "退场工资结清情况",
    "持上岗证情况",
    "联系电话",
]

TEXT_FORMAT_COLUMNS = {"身份证号码", "工资卡号", "劳动合同编号", "联系电话"}


def clean_str(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null", "<na>"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s


def derive_info_from_id(id_card: str):
    """如果主表中张学成等个别工人的出生年月、性别、年龄为空，从合规身份证提取兜底"""
    id_card = clean_str(id_card)
    birth_date, gender, age = "", "", ""
    if len(id_card) == 18 and id_card[:17].isdigit():
        birth_date = f"{id_card[6:10]}-{id_card[10:12]}-{id_card[12:14]}"
        gender = "男" if int(id_card[16]) % 2 != 0 else "女"
        current_year = datetime.datetime.now().year
        birth_year = int(id_card[6:10])
        age = str(current_year - birth_year)
    return birth_date, gender, age


def map_row(worker: dict) -> dict:
    cid = clean_str(worker.get("身份证号"))
    b_date_id, gender_id, age_id = derive_info_from_id(cid)

    name = clean_str(worker.get("姓名"))
    gender = clean_str(worker.get("性别")) or gender_id
    birth_date = clean_str(worker.get("出生日期")) or b_date_id
    age = clean_str(worker.get("年龄")) or age_id
    origin = clean_str(worker.get("籍贯"))
    nation = clean_str(worker.get("民族"))
    political = clean_str(worker.get("政治面貌"))
    job = clean_str(worker.get("工种"))
    address = clean_str(worker.get("详细地址") or worker.get("家庭住址"))
    bank = clean_str(worker.get("开户银行"))
    bank_card = clean_str(worker.get("工资卡号") or worker.get("银行卡号"))
    wage = clean_str(worker.get("结算单价/标准"))
    contract_status = clean_str(worker.get("合同签订状态"))
    contract_no = clean_str(worker.get("劳动合同编号"))
    entry_edu = clean_str(worker.get("进场告知书"))
    entry_date = clean_str(worker.get("进场日期"))
    exit_date = clean_str(worker.get("退场日期"))
    phone = clean_str(worker.get("手机号"))

    row_data = {
        "工人姓名": name,
        "性别": gender,
        "出生日期": birth_date,
        "年龄": age,
        "籍贯": origin,
        "民族": nation,
        "政治面貌": political,
        "工种": job,
        "家庭住址": address,
        "身份证号码": cid,
        "发证机关": "",
        "有效起始日期": "",
        "有效截止日期": "",
        "开户银行": bank,
        "开户支行": "",
        "工资卡号": bank_card,
        "单位工资": wage,
        "劳动合同登记": contract_status,
        "劳动合同编号": contract_no,
        "安全培训情况": "",
        "进场教育": entry_edu,
        "首次进场时间": entry_date,
        "进场时间": entry_date,
        "退场时间": exit_date,
        "退场工资结清情况": "",
        "持上岗证情况": "",
        "联系电话": phone,
    }
    return row_data


def style_worksheet(ws, title: str):
    ws.views.sheetView[0].showGridLines = True

    # 样式定义
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # 沉稳深蓝
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="微软雅黑", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    
    thin_border_side = Side(border_style="thin", color="D1D5DB")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # 写入表头
    ws.append(COLUMNS)
    ws.row_dimensions[1].height = 28

    col_name_to_idx = {col_name: i + 1 for i, col_name in enumerate(COLUMNS)}

    for col_idx in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = cell_border

    return col_name_to_idx, cell_font, center_align, left_align, cell_border, alt_fill


def build_excel_workbook(target_path: Path):
    if not MASTER_STATE_FILE.exists():
        raise FileNotFoundError(f"主表状态文件未找到: {MASTER_STATE_FILE}")

    with open(MASTER_STATE_FILE, "r", encoding="utf-8") as f:
        master_data = json.load(f)

    rows = master_data.get("rows", [])
    if not rows:
        raise ValueError("主表暂无数据")

    # 1. 明确退场人员：在场/进退场状态 为 '已退场' 或 '离场'
    offboarded_27 = [
        r for r in rows
        if clean_str(r.get("在场/进退场状态")) in ("已退场", "离场", "退场", "已离场")
    ]

    # 2. 全量含历史退场记录人员：退场日期不为空（含27人+10名退场后二次进场人员）
    offboarded_37 = [
        r for r in rows
        if clean_str(r.get("退场日期")) != ""
    ]

    wb = openpyxl.Workbook()
    # 默认工作表
    ws1 = wb.active
    ws1.title = "退场人员名单"

    # 填充 Sheet 1 (27人)
    col_map, cell_font, center_align, left_align, cell_border, alt_fill = style_worksheet(ws1, "退场人员名单")

    for row_idx, worker in enumerate(offboarded_27, start=2):
        row_dict = map_row(worker)
        ws1.row_dimensions[row_idx].height = 22
        is_even = (row_idx % 2 == 0)

        for col_idx, col_name in enumerate(COLUMNS, start=1):
            val = row_dict[col_name]
            cell = ws1.cell(row=row_idx, column=col_idx)

            if col_name in TEXT_FORMAT_COLUMNS:
                cell.number_format = "@"
                cell.value = str(val) if val else ""
            else:
                cell.value = val

            cell.font = cell_font
            cell.border = cell_border
            if is_even:
                cell.fill = alt_fill

            # 居中或靠左
            if col_name in ("工人姓名", "性别", "出生日期", "年龄", "民族", "政治面貌", "工种", "首次进场时间", "进场时间", "退场时间"):
                cell.alignment = center_align
            elif col_name in ("家庭住址", "开户银行", "单位工资", "劳动合同登记"):
                cell.alignment = left_align
            else:
                cell.alignment = center_align

    # 填充 Sheet 2 (含历史退场记录37人)
    ws2 = wb.create_sheet(title="含历史退场记录人员(37人)")
    col_map2, cell_font, center_align, left_align, cell_border, alt_fill = style_worksheet(ws2, "含历史退场记录人员(37人)")

    for row_idx, worker in enumerate(offboarded_37, start=2):
        row_dict = map_row(worker)
        ws2.row_dimensions[row_idx].height = 22
        is_even = (row_idx % 2 == 0)

        for col_idx, col_name in enumerate(COLUMNS, start=1):
            val = row_dict[col_name]
            cell = ws2.cell(row=row_idx, column=col_idx)

            if col_name in TEXT_FORMAT_COLUMNS:
                cell.number_format = "@"
                cell.value = str(val) if val else ""
            else:
                cell.value = val

            cell.font = cell_font
            cell.border = cell_border
            if is_even:
                cell.fill = alt_fill

            if col_name in ("工人姓名", "性别", "出生日期", "年龄", "民族", "政治面貌", "工种", "首次进场时间", "进场时间", "退场时间"):
                cell.alignment = center_align
            elif col_name in ("家庭住址", "开户银行", "单位工资", "劳动合同登记"):
                cell.alignment = left_align
            else:
                cell.alignment = center_align

    # 自动调整列宽
    for ws in [ws1, ws2]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val_str = str(cell.value or "")
                # 中文字符计算为2宽度，英文字符计算为1宽度
                w = sum(2 if ord(ch) > 127 else 1 for ch in val_str)
                if w > max_len:
                    max_len = w
            # 留出边距，最小宽度10
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    target_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target_path)
    print(f"成功输出Excel表格: {target_path}")
    print(f"  - Sheet1 [退场人员名单]: {len(offboarded_27)} 人")
    print(f"  - Sheet2 [含历史退场记录人员(37人)]: {len(offboarded_37)} 人")


if __name__ == "__main__":
    build_excel_workbook(OUTPUT_FILE)
