# -*- coding: utf-8 -*-
"""
services/payroll_ledger_service.py
工人工资动态台账服务（支持多分包劳务公司独立核算与全盘汇聚）：
1. 批次数据自动关联分包商（subcontractor）
2. 支持按分包商独立过滤、汇总，支持全项目跨分包合并
3. 生成支持 SUMIFS 动态公式联动的单分包专页与多分包合并 Excel
4. Windows 原生文件夹选择与本地保存
"""

import os
import re
import glob
import io
import pandas as pd
from datetime import datetime
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 项目数据根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BATCHES_DIR = os.path.join(PROJECT_ROOT, "data", "payroll_raw_batches")


def get_batches_dir() -> str:
    """获取并确保批次文件存放目录存在"""
    os.makedirs(BATCHES_DIR, exist_ok=True)
    return BATCHES_DIR


def parse_single_batch_text(text: str, default_company: str = "") -> dict:
    """
    智能解析单个官方工资表文本（CSV/制表符格式）
    自动提取分包商名称、班组、月份、单号、人员数据
    """
    lines = text.strip().splitlines()
    batch = {
        "title": "",
        "batch_short": "",
        "batch_code": "",
        "month_range": "",
        "subcontractor": default_company.strip(),
        "team": "",
        "records": [],
        "declared_total": 0.0,
        "calc_total": 0.0,
        "calc_days": 0.0
    }
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in re.split(r'[,，\t]', line)]
        
        # 1. 优先判断是否是人员数据行（以数字开头，且列数充足）
        if parts and parts[0].isdigit() and len(parts) >= 12:
            try:
                days_str = parts[11].replace("天", "").strip() if len(parts) > 11 else "0"
                days = float(days_str)
            except:
                days = 0.0
            
            try:
                workload = parts[12] if len(parts) > 12 else "0"
            except:
                workload = "0"
            
            try:
                expense = float(parts[15]) if len(parts) > 15 and parts[15] else 0.0
            except:
                expense = 0.0
            
            try:
                base_pay = float(parts[16]) if len(parts) > 16 and parts[16] else 0.0
            except:
                base_pay = 0.0
            
            try:
                allowance = float(parts[18]) if len(parts) > 18 and parts[18] else 0.0
            except:
                allowance = 0.0
            
            try:
                total_pay = float(parts[19]) if len(parts) > 19 and parts[19] else 0.0
            except:
                total_pay = base_pay + allowance
            
            team_val = parts[7] if len(parts) > 7 else batch["team"]
            record = {
                "seq": int(parts[0]),
                "name": parts[1] if len(parts) > 1 else "",
                "id_card": parts[2].strip() if len(parts) > 2 else "",
                "phone": parts[3].strip() if len(parts) > 3 else "",
                "bank_card": parts[4].strip() if len(parts) > 4 else "",
                "bank_name": parts[5] if len(parts) > 5 else "",
                "branch_name": parts[6] if len(parts) > 6 else "",
                "team": team_val,
                "status": parts[8] if len(parts) > 8 else "已进场",
                "rate_str": parts[9] if len(parts) > 9 else "",
                "job_type": parts[10] if len(parts) > 10 else "普工",
                "days": days,
                "workload": workload,
                "expense": expense,
                "base_pay": base_pay,
                "allowance": allowance,
                "total_pay": total_pay,
            }
            batch["records"].append(record)
            continue

        # 2. 表头元数据判断
        if "工资表" in line and not batch["title"]:
            batch["title"] = line.replace(",", " ").strip()
            m = re.search(r'(\d{4}年\d{2}月第\d+批)', batch["title"])
            if m:
                batch["batch_short"] = m.group(1)
            else:
                batch["batch_short"] = batch["title"].split()[0]
        elif line.startswith("工资单编号") or line.startswith("工资单编号："):
            parts_m = re.split(r'[:：]', line)
            if len(parts_m) > 1:
                batch["batch_code"] = parts_m[1].replace(",", "").strip()
        elif line.startswith("所属月份") or line.startswith("所属月份："):
            parts_m = re.split(r'[:：]', line)
            if len(parts_m) > 1:
                batch["month_range"] = parts_m[1].replace(",", "").strip()
        elif line.startswith("分包商名称") or line.startswith("分包商名称：") or line.startswith("分包单位") or line.startswith("劳务公司"):
            parts_m = re.split(r'[:：]', line)
            if len(parts_m) > 1:
                batch["subcontractor"] = parts_m[1].replace(",", "").strip()
        elif line.startswith("班组：") or line.startswith("班组:"):
            parts_m = re.split(r'[:：]', line)
            if len(parts_m) > 1:
                batch["team"] = parts_m[1].replace(",", "").strip()
        elif line.startswith("合计：") or line.startswith("合计"):
            parts_m = [p.strip() for p in re.split(r'[,，\t]', line)]
            for p in parts_m:
                try:
                    val = float(p)
                    if val > 0:
                        batch["declared_total"] = val
                except:
                    pass
    
    # 若表头未识别到分包商名称，尝试从班组名称或标题推断
    if not batch["subcontractor"]:
        context = f"{batch['title']} {batch['team']}"
        if "旭之升" in context:
            batch["subcontractor"] = "江苏旭之升建筑工程有限公司"
        elif "久昌" in context:
            batch["subcontractor"] = "青海久昌建筑装饰工程有限公司"
        elif default_company:
            batch["subcontractor"] = default_company.strip()
        else:
            batch["subcontractor"] = "默认分包劳务单位"

    batch["calc_total"] = sum(r["total_pay"] for r in batch["records"])
    batch["calc_days"] = sum(r["days"] for r in batch["records"])
    if not batch["batch_short"] and batch["title"]:
        batch["batch_short"] = batch["title"][:15]
    return batch


def excel_bytes_to_csv_text(file_bytes: bytes, filename: str = "") -> str:
    """
    将上传的官方 .xlsx 或 .xls 文件的所有单元格内容转为可解析的标准行文本
    完美兼容智慧护薪、中建三局、中建二局等系统的多表头与合并单元格导出格式
    """
    lines = []
    fn_lower = filename.lower()
    
    if fn_lower.endswith(".xls") and not fn_lower.endswith(".xlsx"):
        try:
            import pandas as pd
            xls_file = io.BytesIO(file_bytes)
            excel_dict = pd.read_excel(xls_file, sheet_name=None, header=None, dtype=str)
            for sname, df in excel_dict.items():
                if df.empty:
                    continue
                for _, row in df.iterrows():
                    row_vals = ["" if pd.isna(v) else str(v).replace(",", " ").strip() for v in row.values]
                    while row_vals and not row_vals[-1]:
                        row_vals.pop()
                    if any(row_vals):
                        lines.append(",".join(row_vals))
        except Exception as e:
            print(f"解析 .xls 失败: {e}")
    else:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                for row in ws.iter_rows(values_only=True):
                    row_vals = ["" if v is None else str(v).replace(",", " ").strip() for v in row]
                    while row_vals and not row_vals[-1]:
                        row_vals.pop()
                    if any(row_vals):
                        lines.append(",".join(row_vals))
        except Exception as e:
            print(f"解析 .xlsx 失败: {e}")
            
    return "\n".join(lines)


def parse_batch_file_content(content_str: str, default_company: str = "") -> list:
    """解析可能包含一个或多个批次的完整文本字符串"""
    splits = re.split(r'(?=科技文化中心-国际体育中心项目（专业足球场）\d{4}年\d{2}月第\d+批次工资表)', content_str)
    if len(splits) <= 1:
        splits = re.split(r'(?=\n.*?(?:20\d\d年\d\d月第\d+批次?工资表))', content_str)
    
    batches = []
    for sp in splits:
        sp = sp.strip()
        if not sp:
            continue
        b = parse_single_batch_text(sp, default_company=default_company)
        if b and len(b["records"]) > 0:
            batches.append(b)
    
    if not batches:
        b = parse_single_batch_text(content_str, default_company=default_company)
        if b and len(b["records"]) > 0:
            batches.append(b)
    return batches


def save_uploaded_batch(content_input, assigned_company: str = "", custom_name: str = None, filename: str = "") -> tuple:
    """
    保存上传的批次文件到 BATCHES_DIR（支持 Excel 二进制流、文本字符串）
    自动带上分包商前缀
    返回 (成功数量, 提示信息列表)
    """
    if isinstance(content_input, bytes):
        if filename.lower().endswith((".xlsx", ".xls")):
            content_str = excel_bytes_to_csv_text(content_input, filename=filename)
        else:
            try:
                content_str = content_input.decode("utf-8-sig")
            except:
                try:
                    content_str = content_input.decode("gbk")
                except:
                    content_str = content_input.decode("utf-8", errors="ignore")
    else:
        content_str = str(content_input)

    batches = parse_batch_file_content(content_str, default_company=assigned_company)
    if not batches:
        return 0, ["未识别到合规的官方工资表记录，请确认上传的 Excel 或文件格式（需包含姓名、身份证号、出勤、实发工资等列）"]
    
    saved_count = 0
    messages = []
    b_dir = get_batches_dir()
    
    for b in batches:
        comp = assigned_company.strip() if assigned_company.strip() else b["subcontractor"]
        b["subcontractor"] = comp
        b_name = b["batch_short"] or custom_name or f"批次_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        safe_comp = re.sub(r'[\/\\:\*\?"<>\|]', '_', comp)
        safe_name = re.sub(r'[\/\\:\*\?"<>\|]', '_', b_name)
        code_tag = f"_{b['batch_code'][-6:]}" if b["batch_code"] else ""
        
        file_name = f"{safe_comp}_{safe_name}{code_tag}.csv"
        file_path = os.path.join(b_dir, file_name)
        
        to_write = content_str
        if f"分包商名称：{comp}" not in to_write and f"分包商名称:{comp}" not in to_write:
            to_write = f"分包商名称：{comp}\n" + to_write

        with open(file_path, "w", encoding="utf-8-sig") as f:
            f.write(to_write)
        
        saved_count += 1
        messages.append(f"成功归档至【{comp}】：批次【{b_name}】（共 {len(b['records'])} 人，金额 ¥{b['calc_total']:,.2f}）")
    
    return saved_count, messages


def list_all_batch_files() -> list:
    """列出当前已存储的所有批次文件信息"""
    b_dir = get_batches_dir()
    files = (
        glob.glob(os.path.join(b_dir, "*.csv")) +
        glob.glob(os.path.join(b_dir, "*.txt")) +
        glob.glob(os.path.join(b_dir, "*.xlsx")) +
        glob.glob(os.path.join(b_dir, "*.xls"))
    )
    res = []
    for fp in files:
        fname = os.path.basename(fp)
        mtime = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M")
        size_kb = round(os.path.getsize(fp) / 1024, 1)
        res.append({
            "filename": fname,
            "filepath": fp,
            "mtime": mtime,
            "size": f"{size_kb} KB"
        })
    return res


def delete_batch_file(filename: str) -> bool:
    """删除指定的批次文件"""
    b_dir = get_batches_dir()
    fp = os.path.join(b_dir, filename)
    if os.path.exists(fp):
        os.remove(fp)
        return True
    return False


def get_all_companies() -> list:
    """扫描所有批次，返回当前存在的所有分包商列表"""
    data = load_all_payroll_data(company_filter=None)
    comps = set()
    for b in data["batches"]:
        if b.get("subcontractor"):
            comps.add(b["subcontractor"])
    # 结合项目常驻分包商兜底
    known_comps = ["江苏旭之升建筑工程有限公司", "青海久昌建筑装饰工程有限公司"]
    for kc in known_comps:
        comps.add(kc)
    return sorted(list(comps))


def load_all_payroll_data(company_filter: str = None) -> dict:
    """
    全量扫描并汇总批次数据：
    如果指定了 company_filter（非空且不是"全部劳务公司"），则只汇总该分包商的数据；
    否则汇总全部分包商的数据。
    """
    b_dir = get_batches_dir()
    files = (
        glob.glob(os.path.join(b_dir, "*.csv")) +
        glob.glob(os.path.join(b_dir, "*.txt")) +
        glob.glob(os.path.join(b_dir, "*.xlsx")) +
        glob.glob(os.path.join(b_dir, "*.xls"))
    )
    
    all_raw_batches = []
    for fp in files:
        content = ""
        if fp.lower().endswith((".xlsx", ".xls")):
            try:
                with open(fp, "rb") as f_b:
                    content = excel_bytes_to_csv_text(f_b.read(), filename=fp)
            except Exception as e:
                print(f"读取 Excel 批次失败 {fp}: {e}")
                content = ""
        else:
            encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030']
            for enc in encodings:
                try:
                    with open(fp, 'r', encoding=enc) as f:
                        content = f.read()
                    break
                except:
                    continue
        if content:
            b_list = parse_batch_file_content(content)
            all_raw_batches.extend(b_list)
    
    # 自动去重：依据官方工资单编号 batch_code 或 (分包商, 批次简称) 唯一去重，避免重复上传导致翻倍
    deduped_raw_batches = []
    seen_batch_keys = set()
    for b in all_raw_batches:
        b_key = b.get("batch_code", "").strip()
        if not b_key:
            b_key = f"{b.get('subcontractor', '')}_{b.get('batch_short', '')}"
        if b_key not in seen_batch_keys:
            seen_batch_keys.add(b_key)
            deduped_raw_batches.append(b)
    all_raw_batches = deduped_raw_batches

    # 按照月份和批次简称自然排序
    all_raw_batches.sort(key=lambda b: (b["subcontractor"], b["month_range"], b["batch_short"]))
    
    # 过滤分包商
    if company_filter and company_filter.strip() not in ("全部劳务公司", "全部分包商", "全部", ""):
        target_comp = company_filter.strip()
        all_batches = [b for b in all_raw_batches if b["subcontractor"] == target_comp]
    else:
        all_batches = all_raw_batches

    # 批次唯一简称
    unique_batch_shorts = []
    for b in all_batches:
        bs = b["batch_short"]
        # 如果是多分包合并视图，批次加上分包商简称前缀以防混淆
        if not company_filter or company_filter in ("全部劳务公司", "全部分包商", "全部"):
            comp_short = b["subcontractor"][:4] if b["subcontractor"] else ""
            bs = f"{comp_short}_{bs}"
        if bs in unique_batch_shorts:
            bs = f"{bs}({b['batch_code'][-4:] if b['batch_code'] else '新'})"
        unique_batch_shorts.append(bs)
        b["final_batch_name"] = bs

    workers = OrderedDict()
    all_flow_records = []

    for b in all_batches:
        b_name = b["final_batch_name"]
        b_comp = b["subcontractor"]
        for r in b["records"]:
            # 唯一工人键：在同一公司内按身份证号，在全盘时也是身份证号
            cid = r["id_card"]
            if cid not in workers:
                workers[cid] = {
                    "name": r["name"],
                    "id_card": cid,
                    "company": b_comp,
                    "phone": r["phone"],
                    "bank_card": r["bank_card"],
                    "bank_name": r["bank_name"],
                    "branch_name": r["branch_name"],
                    "job_type": r["job_type"],
                    "rate_str": r["rate_str"],
                    "latest_status": r["status"],
                    "batch_pays": {},
                    "batch_days": {},
                    "total_days": 0.0,
                    "grand_total": 0.0,
                }
            
            # 更新最新属性
            workers[cid]["company"] = b_comp
            workers[cid]["phone"] = r["phone"]
            workers[cid]["bank_card"] = r["bank_card"]
            if r["bank_name"]:
                workers[cid]["bank_name"] = r["bank_name"]
            if r["branch_name"]:
                workers[cid]["branch_name"] = r["branch_name"]
            workers[cid]["job_type"] = r["job_type"]
            workers[cid]["rate_str"] = r["rate_str"]
            workers[cid]["latest_status"] = r["status"]
            
            workers[cid]["batch_pays"][b_name] = workers[cid]["batch_pays"].get(b_name, 0.0) + r["total_pay"]
            workers[cid]["batch_days"][b_name] = workers[cid]["batch_days"].get(b_name, 0.0) + r["days"]
            
            workers[cid]["total_days"] += r["days"]
            workers[cid]["grand_total"] += r["total_pay"]

            all_flow_records.append({
                "所属劳务公司": b_comp,
                "发放批次": b_name,
                "工资单编号": b["batch_code"],
                "所属月份": b["month_range"],
                "序号": r["seq"],
                "姓名": r["name"],
                "身份证号": r["id_card"],
                "电话号码": r["phone"],
                "工种": r["job_type"],
                "进退场状态": r["status"],
                "工资标准": r["rate_str"].replace("基本工资：", ""),
                "实际出勤(天)": r["days"],
                "扣减支出": r["expense"],
                "实发基本工资": r["base_pay"],
                "津贴": r["allowance"],
                "实发总额(元)": r["total_pay"],
                "实名制银行卡号": r["bank_card"],
                "开户银行": r["bank_name"],
                "开户支行": r["branch_name"],
            })

    sorted_workers = sorted(workers.values(), key=lambda x: x["grand_total"], reverse=True)

    summary_rows = []
    for idx, w in enumerate(sorted_workers, 1):
        row = {
            "序号": idx,
            "姓名": w["name"],
            "所属劳务公司": w["company"],
            "身份证号": w["id_card"],
            "联系电话": w["phone"],
            "工种": w["job_type"],
            "状态": w["latest_status"],
            "日薪标准": w["rate_str"].replace("基本工资：", ""),
        }
        for bs in unique_batch_shorts:
            p = w["batch_pays"].get(bs, 0.0)
            d = w["batch_days"].get(bs, 0.0)
            row[f"{bs}_天数"] = d if d > 0 else "-"
            row[f"{bs}_实发"] = f"¥{p:,.2f}" if p > 0 else "-"
        
        row["累计出勤(天)"] = w["total_days"]
        row["累计实发总额"] = w["grand_total"]
        row["实名制银行卡号"] = w["bank_card"]
        row["开户银行及支行"] = f"{w['bank_name']}{(' - ' + w['branch_name']) if w['branch_name'] else ''}"
        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame()
    flow_df = pd.DataFrame(all_flow_records) if all_flow_records else pd.DataFrame()

    tot_pay = sum(w["grand_total"] for w in sorted_workers)
    tot_days = sum(w["total_days"] for w in sorted_workers)
    active_workers = [w for w in sorted_workers if w["latest_status"] == "已进场"]
    left_workers = [w for w in sorted_workers if w["latest_status"] == "已退场"]

    # 统计各个分包公司的概况
    comp_stats = {}
    for w in sorted_workers:
        c = w["company"]
        if c not in comp_stats:
            comp_stats[c] = {"worker_count": 0, "total_pay": 0.0, "total_days": 0.0, "batch_count": 0}
        comp_stats[c]["worker_count"] += 1
        comp_stats[c]["total_pay"] += w["grand_total"]
        comp_stats[c]["total_days"] += w["total_days"]
    for b in all_batches:
        c = b["subcontractor"]
        if c in comp_stats:
            comp_stats[c]["batch_count"] = comp_stats[c].get("batch_count", 0) + 1

    stats = {
        "batch_count": len(all_batches),
        "worker_count": len(sorted_workers),
        "flow_count": len(all_flow_records),
        "total_pay": tot_pay,
        "total_days": tot_days,
        "active_count": len(active_workers),
        "active_pay": sum(w["grand_total"] for w in active_workers),
        "left_count": len(left_workers),
        "left_pay": sum(w["grand_total"] for w in left_workers),
        "comp_stats": comp_stats
    }

    return {
        "company_scope": company_filter or "全部劳务公司",
        "batches": all_batches,
        "batch_shorts": unique_batch_shorts,
        "workers": workers,
        "sorted_workers": sorted_workers,
        "summary_df": summary_df,
        "flow_df": flow_df,
        "stats": stats
    }


def build_excel_workbook(data: dict) -> openpyxl.Workbook:
    """
    根据汇总数据构建高质感活表 Excel：
    自动根据是否为单公司还是多分包，组织 Sheet 架构
    """
    unique_batch_shorts = data["batch_shorts"]
    sorted_workers = data["sorted_workers"]
    all_batches = data["batches"]
    flow_records = data["flow_df"].to_dict("records") if not data["flow_df"].empty else []
    company_scope = data.get("company_scope", "全部劳务公司")

    wb = openpyxl.Workbook()

    font_title = Font(name="微软雅黑", size=15, bold=True, color="1F4E79")
    font_subtitle = Font(name="微软雅黑", size=10, italic=True, color="595959")
    font_header = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="微软雅黑", size=10, color="000000")
    font_total = Font(name="微软雅黑", size=10, bold=True, color="000000")
    font_status_active = Font(name="微软雅黑", size=10, color="276A3C", bold=True)
    font_status_left = Font(name="微软雅黑", size=10, color="A61C1C", bold=True)

    fill_primary = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sub_blue = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_accent_blue = PatternFill(start_color="41719C", end_color="41719C", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_status_active = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_status_left = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    double_border_bottom = Side(border_style="double", color="000000")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(border_style="thin", color="000000"), bottom=double_border_bottom, left=thin_border_side, right=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Sheet 1: 使用指南
    ws_guide = wb.active
    ws_guide.title = "使用指南(多分包如何追加)"
    ws_guide.views.sheetView[0].showGridLines = True
    
    ws_guide["A1"] = f"📘 工资发放动态台账 · 永久可追加使用指南（{company_scope}）"
    ws_guide["A1"].font = Font(name="微软雅黑", size=16, bold=True, color="1F4E79")
    ws_guide.row_dimensions[1].height = 40
    
    guide_lines = [
        ("【多分包架构原则】", "本台账支持工程项目部下辖各分包劳务公司（如江苏旭之升、青海久昌等）独立建账核算与全盘汇聚。"),
        ("", ""),
        ("方案 A：在系统界面一键追加新批次（最推荐）", ""),
        ("第 1 步：选择分包公司", "在系统【考勤对账与工资结算】->【工资发放总台账】顶部，切换到对应的劳务分包商。"),
        ("第 2 步：上传/粘贴批次", "将官方导出的工资表文本粘贴或上传，系统会自动归类至该劳务公司并归档入库。"),
        ("第 3 步：选文件夹导出", "点击【选择文件夹并导出】，生成的台账专属于该劳务公司（或全盘汇总），账目清晰绝对不串户！"),
        ("", ""),
        ("方案 B：纯 Excel 离线操作追加（无需系统）", ""),
        ("第 1 步：复制流水", "将官方导出的新工资表数据，直接追加粘贴到【流水明细库】表格的最下方。"),
        ("第 2 步：自动计入累计", "【全员工资总台账】里的『累计实发总额』和『累计出勤天数』公式自动生效，瞬间计算！"),
        ("第 3 步：横向加列", "在总台账表头输入新批次名，选中上一列公式向右一拖，立刻出数！"),
    ]
    for r_i, (colA, colB) in enumerate(guide_lines, 3):
        ws_guide.row_dimensions[r_i].height = 24
        c_a = ws_guide.cell(row=r_i, column=1, value=colA)
        c_b = ws_guide.cell(row=r_i, column=2, value=colB)
        if "【" in colA or "方案" in colA:
            c_a.font = Font(name="微软雅黑", size=11, bold=True, color="1F4E79")
            c_b.font = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")
        elif "第" in colA:
            c_a.font = Font(name="微软雅黑", size=10, bold=True, color="C00000")
            c_b.font = Font(name="微软雅黑", size=10, color="333333")
        else:
            c_a.font = Font(name="微软雅黑", size=10, color="555555")
            c_b.font = Font(name="微软雅黑", size=10, color="555555")
        c_a.alignment = Alignment(vertical="center")
        c_b.alignment = Alignment(vertical="center", wrap_text=True)

    ws_guide.column_dimensions["A"].width = 30
    ws_guide.column_dimensions["B"].width = 85

    # Sheet 2: 全员工资总台账 (SUMIFS 动态公式版)
    is_multi_comp = (company_scope in ("全部劳务公司", "全部分包商", "全部"))
    sheet_title = "全员工资总台账" if not is_multi_comp else "全项目各分包工资总台账"
    ws_main = wb.create_sheet(title=sheet_title)
    ws_main.views.sheetView[0].showGridLines = True

    # 包含公司列
    num_fixed_left = 8 if is_multi_comp else 7
    num_batch_cols = len(unique_batch_shorts) * 2
    tot_cols = num_fixed_left + num_batch_cols + 4

    ws_main.merge_cells(start_row=1, start_column=1, end_row=1, end_column=tot_cols)
    scope_title_str = f"【{company_scope}】" if not is_multi_comp else "【全项目各劳务分包大盘】"
    ws_main["A1"] = f"科技文化中心-国际体育中心项目（专业足球场）工人实发工资动态总台账 {scope_title_str}"
    ws_main["A1"].font = font_title
    ws_main["A1"].alignment = align_center
    ws_main.row_dimensions[1].height = 36

    ws_main.merge_cells(start_row=2, start_column=1, end_row=2, end_column=tot_cols)
    period_str = f"{all_batches[0]['month_range'][:7]} 至 {all_batches[-1]['month_range'][:7]}" if all_batches else "当期"
    ws_main["A2"] = f"劳务分包归属：{company_scope} | 已归档：{len(unique_batch_shorts)} 批次 | 覆盖周期：{period_str} | 底层公式完全联动"
    ws_main["A2"].font = font_subtitle
    ws_main["A2"].alignment = align_center
    ws_main.row_dimensions[2].height = 20

    ws_main.row_dimensions[3].height = 24
    ws_main.row_dimensions[4].height = 22

    if is_multi_comp:
        fixed_headers = [
            ("序号", 1), ("姓名", 2), ("所属劳务公司", 3), ("身份证号", 4),
            ("联系电话", 5), ("工种", 6), ("当前状态", 7), ("薪资标准", 8)
        ]
        col_id_idx = 4
        col_status_idx = 7
        cur_c = 9
    else:
        fixed_headers = [
            ("序号", 1), ("姓名", 2), ("身份证号", 3), ("联系电话", 4),
            ("工种", 5), ("当前状态", 6), ("最新薪资标准", 7)
        ]
        col_id_idx = 3
        col_status_idx = 6
        cur_c = 8

    for h, col_c in fixed_headers:
        ws_main.merge_cells(start_row=3, start_column=col_c, end_row=4, end_column=col_c)
        c3 = ws_main.cell(row=3, column=col_c, value=h)
        c3.font = font_header
        c3.fill = fill_primary
        c3.alignment = align_center
        c3.border = cell_border
        c4 = ws_main.cell(row=4, column=col_c)
        c4.fill = fill_primary
        c4.border = cell_border

    batch_col_map = {}
    for bs in unique_batch_shorts:
        ws_main.merge_cells(start_row=3, start_column=cur_c, end_row=3, end_column=cur_c + 1)
        c_top = ws_main.cell(row=3, column=cur_c, value=bs)
        c_top.font = font_header
        c_top.fill = fill_primary
        c_top.alignment = align_center
        c_top.border = cell_border
        ws_main.cell(row=3, column=cur_c + 1).fill = fill_primary
        ws_main.cell(row=3, column=cur_c + 1).border = cell_border
        
        c_sub1 = ws_main.cell(row=4, column=cur_c, value="出勤(天)")
        c_sub1.font = font_sub_header
        c_sub1.fill = fill_sub_blue
        c_sub1.alignment = align_center
        c_sub1.border = cell_border
        
        c_sub2 = ws_main.cell(row=4, column=cur_c + 1, value="实发(元)")
        c_sub2.font = font_sub_header
        c_sub2.fill = fill_sub_blue
        c_sub2.alignment = align_center
        c_sub2.border = cell_border
        
        batch_col_map[bs] = (cur_c, cur_c + 1)
        cur_c += 2

    right_headers = [
        ("累计总出勤(天)", cur_c),
        ("累计实发总额(元)", cur_c + 1),
        ("实名制银行卡号", cur_c + 2),
        ("开户银行及支行", cur_c + 3)
    ]
    col_tot_days = cur_c
    col_grand_pay = cur_c + 1
    col_card = cur_c + 2
    col_bank = cur_c + 3

    for h, col_c in right_headers:
        ws_main.merge_cells(start_row=3, start_column=col_c, end_row=4, end_column=col_c)
        c3 = ws_main.cell(row=3, column=col_c, value=h)
        c3.font = font_header
        c3.fill = fill_accent_blue if "累计" in h else fill_primary
        c3.alignment = align_center
        c3.border = cell_border
        c4 = ws_main.cell(row=4, column=col_c)
        c4.fill = fill_accent_blue if "累计" in h else fill_primary
        c4.border = cell_border

    start_r = 5
    for idx, w in enumerate(sorted_workers, 1):
        r_num = start_r + idx - 1
        ws_main.row_dimensions[r_num].height = 21
        
        col_ptr = 1
        ws_main.cell(row=r_num, column=col_ptr, value=idx).alignment = align_center
        col_ptr += 1
        ws_main.cell(row=r_num, column=col_ptr, value=w["name"]).alignment = align_center
        col_ptr += 1

        if is_multi_comp:
            ws_main.cell(row=r_num, column=col_ptr, value=w["company"]).alignment = align_center
            col_ptr += 1
        
        c_id = ws_main.cell(row=r_num, column=col_ptr, value=w["id_card"])
        c_id.alignment = align_center
        c_id.number_format = "@"
        col_id_cell_coord = f"{get_column_letter(col_ptr)}{r_num}"
        col_ptr += 1
        
        c_ph = ws_main.cell(row=r_num, column=col_ptr, value=w["phone"])
        c_ph.alignment = align_center
        c_ph.number_format = "@"
        col_ptr += 1
        
        ws_main.cell(row=r_num, column=col_ptr, value=w["job_type"]).alignment = align_center
        col_ptr += 1
        
        c_st = ws_main.cell(row=r_num, column=col_ptr, value=w["latest_status"])
        c_st.alignment = align_center
        if w["latest_status"] == "已进场":
            c_st.fill = fill_status_active
            c_st.font = font_status_active
        else:
            c_st.fill = fill_status_left
            c_st.font = font_status_left
        col_ptr += 1
            
        ws_main.cell(row=r_num, column=col_ptr, value=w["rate_str"].replace("基本工资：", "")).alignment = align_center
        col_ptr += 1

        # 公式联动流水库
        # 流水库对应列：B=发放批次, G=身份证号, L=实际出勤天数, P=实发总额(元)
        for bs in unique_batch_shorts:
            c_day, c_pay = batch_col_map[bs]
            batch_cell_ref = f"{get_column_letter(c_day)}$3"
            
            formula_day = f'=SUMIFS(流水明细库!$L:$L, 流水明细库!$G:$G, ${col_id_cell_coord}, 流水明细库!$B:$B, {batch_cell_ref})'
            cell_d = ws_main.cell(row=r_num, column=c_day, value=formula_day)
            cell_d.alignment = align_center
            
            formula_pay = f'=SUMIFS(流水明细库!$P:$P, 流水明细库!$G:$G, ${col_id_cell_coord}, 流水明细库!$B:$B, {batch_cell_ref})'
            cell_p = ws_main.cell(row=r_num, column=c_pay, value=formula_pay)
            cell_p.alignment = align_right
            cell_p.number_format = "#,##0.00"

        formula_tot_day = f'=SUMIFS(流水明细库!$L:$L, 流水明细库!$G:$G, ${col_id_cell_coord})'
        c_td = ws_main.cell(row=r_num, column=col_tot_days, value=formula_tot_day)
        c_td.alignment = align_center
        c_td.font = Font(name="微软雅黑", size=10, bold=True)

        formula_tot_pay = f'=SUMIFS(流水明细库!$P:$P, 流水明细库!$G:$G, ${col_id_cell_coord})'
        c_tp = ws_main.cell(row=r_num, column=col_grand_pay, value=formula_tot_pay)
        c_tp.alignment = align_right
        c_tp.number_format = "¥#,##0.00"
        c_tp.font = Font(name="微软雅黑", size=10, bold=True, color="C00000")

        c_card = ws_main.cell(row=r_num, column=col_card, value=w["bank_card"])
        c_card.alignment = align_center
        c_card.number_format = "@"
        
        bank_str = f"{w['bank_name']}{(' - ' + w['branch_name']) if w['branch_name'] else ''}"
        c_bank = ws_main.cell(row=r_num, column=col_bank, value=bank_str)
        c_bank.alignment = align_left

        for c_i in range(1, tot_cols + 1):
            cell = ws_main.cell(row=r_num, column=c_i)
            cell.border = cell_border
            if c_i != col_status_idx and idx % 2 == 0:
                cell.fill = fill_zebra
            if c_i not in (col_status_idx, col_tot_days, col_grand_pay):
                cell.font = font_data

    tot_row = start_r + len(sorted_workers)
    ws_main.row_dimensions[tot_row].height = 26
    ws_main.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=num_fixed_left)
    c_tot_lbl = ws_main.cell(row=tot_row, column=1, value=f"合计（共 {len(sorted_workers)} 人）：")
    c_tot_lbl.font = font_total
    c_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")
    
    for c_i in range(1, num_fixed_left + 1):
        ws_main.cell(row=tot_row, column=c_i).fill = fill_total
        ws_main.cell(row=tot_row, column=c_i).border = total_border

    for bs in unique_batch_shorts:
        c_d, c_p = batch_col_map[bs]
        col_letter_d = get_column_letter(c_d)
        col_letter_p = get_column_letter(c_p)
        
        c_sum_d = ws_main.cell(row=tot_row, column=c_d, value=f"=SUM({col_letter_d}{start_r}:{col_letter_d}{tot_row-1})")
        c_sum_d.alignment = align_center
        c_sum_d.font = font_total
        c_sum_d.fill = fill_total
        c_sum_d.border = total_border
        
        c_sum_p = ws_main.cell(row=tot_row, column=c_p, value=f"=SUM({col_letter_p}{start_r}:{col_letter_p}{tot_row-1})")
        c_sum_p.alignment = align_right
        c_sum_p.number_format = "#,##0.00"
        c_sum_p.font = font_total
        c_sum_p.fill = fill_total
        c_sum_p.border = total_border

    col_let_td = get_column_letter(col_tot_days)
    col_let_gp = get_column_letter(col_grand_pay)
    
    c_tot_td = ws_main.cell(row=tot_row, column=col_tot_days, value=f"=SUM({col_let_td}{start_r}:{col_let_td}{tot_row-1})")
    c_tot_td.alignment = align_center
    c_tot_td.font = font_total
    c_tot_td.fill = fill_total
    c_tot_td.border = total_border
    
    c_tot_gp = ws_main.cell(row=tot_row, column=col_grand_pay, value=f"=SUM({col_let_gp}{start_r}:{col_let_gp}{tot_row-1})")
    c_tot_gp.alignment = align_right
    c_tot_gp.number_format = "¥#,##0.00"
    c_tot_gp.font = font_total
    c_tot_gp.fill = fill_total
    c_tot_gp.border = total_border

    ws_main.cell(row=tot_row, column=col_card, value="").fill = fill_total
    ws_main.cell(row=tot_row, column=col_card).border = total_border
    ws_main.cell(row=tot_row, column=col_bank, value="").fill = fill_total
    ws_main.cell(row=tot_row, column=col_bank).border = total_border

    freeze_col_letter = get_column_letter(num_fixed_left + 1)
    ws_main.freeze_panes = f"{freeze_col_letter}5"

    ws_main.column_dimensions["A"].width = 6
    ws_main.column_dimensions["B"].width = 10
    if is_multi_comp:
        ws_main.column_dimensions["C"].width = 24
        ws_main.column_dimensions["D"].width = 22
        ws_main.column_dimensions["E"].width = 15
        ws_main.column_dimensions["F"].width = 12
        ws_main.column_dimensions["G"].width = 10
        ws_main.column_dimensions["H"].width = 13
    else:
        ws_main.column_dimensions["C"].width = 22
        ws_main.column_dimensions["D"].width = 15
        ws_main.column_dimensions["E"].width = 12
        ws_main.column_dimensions["F"].width = 10
        ws_main.column_dimensions["G"].width = 13
    for bs in unique_batch_shorts:
        c_d, c_p = batch_col_map[bs]
        ws_main.column_dimensions[get_column_letter(c_d)].width = 10
        ws_main.column_dimensions[get_column_letter(c_p)].width = 13
    ws_main.column_dimensions[get_column_letter(col_tot_days)].width = 14
    ws_main.column_dimensions[get_column_letter(col_grand_pay)].width = 17
    ws_main.column_dimensions[get_column_letter(col_card)].width = 24
    ws_main.column_dimensions[get_column_letter(col_bank)].width = 38

    # Sheet 3: 退场工人结算台账
    ws_off = wb.create_sheet(title="已退场工人结算台账")
    ws_off.views.sheetView[0].showGridLines = True
    
    ws_off.merge_cells("A1:L1")
    ws_off["A1"] = f"已办理退场工人工资结清与销账台账 ({company_scope})"
    ws_off["A1"].font = font_title
    ws_off["A1"].alignment = align_center
    ws_off.row_dimensions[1].height = 36
    
    off_headers = [
        "序号", "姓名", "所属分包公司", "身份证号", "联系电话", "工种", "离场状态",
        "日薪标准", "累计出勤(天)", "已结清总额(元)", "实名制银行卡号", "开户支行"
    ]
    ws_off.row_dimensions[2].height = 24
    for c_i, h in enumerate(off_headers, 1):
        cell = ws_off.cell(row=2, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_primary
        cell.alignment = align_center
        cell.border = cell_border

    off_workers = [w for w in sorted_workers if w["latest_status"] == "已退场"]
    for idx, w in enumerate(off_workers, 1):
        r_num = 2 + idx
        ws_off.row_dimensions[r_num].height = 20
        ws_off.cell(row=r_num, column=1, value=idx).alignment = align_center
        ws_off.cell(row=r_num, column=2, value=w["name"]).alignment = align_center
        ws_off.cell(row=r_num, column=3, value=w["company"]).alignment = align_center
        
        c_id = ws_off.cell(row=r_num, column=4, value=w["id_card"])
        c_id.alignment = align_center
        c_id.number_format = "@"
        
        c_ph = ws_off.cell(row=r_num, column=5, value=w["phone"])
        c_ph.alignment = align_center
        c_ph.number_format = "@"
        
        ws_off.cell(row=r_num, column=6, value=w["job_type"]).alignment = align_center
        
        c_st = ws_off.cell(row=r_num, column=7, value=w["latest_status"])
        c_st.alignment = align_center
        c_st.fill = fill_status_left
        c_st.font = font_status_left
        
        ws_off.cell(row=r_num, column=8, value=w["rate_str"].replace("基本工资：", "")).alignment = align_center
        
        # 对应流水明细库列：G=身份证, L=出勤天数, P=实发总额
        formula_td = f'=SUMIFS(流水明细库!$L:$L, 流水明细库!$G:$G, $D{r_num})'
        c_td = ws_off.cell(row=r_num, column=9, value=formula_td)
        c_td.alignment = align_center
        c_td.font = Font(name="微软雅黑", size=10, bold=True)
        
        formula_pay = f'=SUMIFS(流水明细库!$P:$P, 流水明细库!$G:$G, $D{r_num})'
        c_p = ws_off.cell(row=r_num, column=10, value=formula_pay)
        c_p.alignment = align_right
        c_p.number_format = "¥#,##0.00"
        c_p.font = Font(name="微软雅黑", size=10, bold=True, color="C00000")
        
        c_bc = ws_off.cell(row=r_num, column=11, value=w["bank_card"])
        c_bc.alignment = align_center
        c_bc.number_format = "@"
        
        bank_str = f"{w['bank_name']}{(' - ' + w['branch_name']) if w['branch_name'] else ''}"
        ws_off.cell(row=r_num, column=12, value=bank_str).alignment = align_left
        
        for c_i in range(1, 13):
            cell = ws_off.cell(row=r_num, column=c_i)
            cell.border = cell_border
            if c_i != 7 and idx % 2 == 0:
                cell.fill = fill_zebra
            if c_i not in (7, 9, 10):
                cell.font = font_data

    tot_off_r = len(off_workers) + 3
    ws_off.row_dimensions[tot_off_r].height = 24
    ws_off.merge_cells(start_row=tot_off_r, start_column=1, end_row=tot_off_r, end_column=8)
    c_tot_o = ws_off.cell(row=tot_off_r, column=1, value=f"退场人员合计（共 {len(off_workers)} 人）：")
    c_tot_o.font = font_total
    c_tot_o.alignment = Alignment(horizontal="right", vertical="center")
    
    ws_off.cell(row=tot_off_r, column=9, value=f"=SUM(I3:I{tot_off_r-1})").alignment = align_center
    c_o_pay = ws_off.cell(row=tot_off_r, column=10, value=f"=SUM(J3:J{tot_off_r-1})")
    c_o_pay.number_format = "¥#,##0.00"
    
    for c_i in range(1, 13):
        cell = ws_off.cell(row=tot_off_r, column=c_i)
        cell.fill = fill_total
        cell.font = font_total
        cell.border = total_border

    ws_off.freeze_panes = "H3"
    ws_off.column_dimensions["A"].width = 6
    ws_off.column_dimensions["B"].width = 10
    ws_off.column_dimensions["C"].width = 24
    ws_off.column_dimensions["D"].width = 22
    ws_off.column_dimensions["E"].width = 15
    ws_off.column_dimensions["F"].width = 12
    ws_off.column_dimensions["G"].width = 10
    ws_off.column_dimensions["H"].width = 13
    ws_off.column_dimensions["I"].width = 14
    ws_off.column_dimensions["J"].width = 18
    ws_off.column_dimensions["K"].width = 24
    ws_off.column_dimensions["L"].width = 38

    # Sheet 4: 流水明细库 (底层数据库，包含分包商列)
    ws_flow = wb.create_sheet(title="流水明细库")
    ws_flow.views.sheetView[0].showGridLines = True

    flow_headers = [
        "所属分包公司", "发放批次", "工资单编号", "所属月份", "原序号", "姓名", "身份证号", "电话号码",
        "工种", "进退场状态", "工资标准", "实际出勤(天)", "扣减支出", "实发基本工资",
        "津贴", "实发总额(元)", "实名制银行卡号", "开户银行", "开户支行"
    ]
    ws_flow.row_dimensions[1].height = 28
    for c_i, h in enumerate(flow_headers, 1):
        cell = ws_flow.cell(row=1, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_primary
        cell.alignment = align_center
        cell.border = cell_border

    for r_idx, rec in enumerate(flow_records, 2):
        ws_flow.row_dimensions[r_idx].height = 20
        ws_flow.cell(row=r_idx, column=1, value=rec.get("所属劳务公司", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=2, value=rec.get("发放批次", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=3, value=rec.get("工资单编号", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=4, value=rec.get("所属月份", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=5, value=rec.get("序号", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=6, value=rec.get("姓名", "")).alignment = align_center
        
        c_cid = ws_flow.cell(row=r_idx, column=7, value=rec.get("身份证号", ""))
        c_cid.alignment = align_center
        c_cid.number_format = "@"
        
        c_ph = ws_flow.cell(row=r_idx, column=8, value=rec.get("电话号码", ""))
        c_ph.alignment = align_center
        c_ph.number_format = "@"
        
        ws_flow.cell(row=r_idx, column=9, value=rec.get("工种", "")).alignment = align_center
        
        c_st = ws_flow.cell(row=r_idx, column=10, value=rec.get("进退场状态", ""))
        c_st.alignment = align_center
        if rec.get("进退场状态") == "已进场":
            c_st.fill = fill_status_active
            c_st.font = font_status_active
        else:
            c_st.fill = fill_status_left
            c_st.font = font_status_left

        ws_flow.cell(row=r_idx, column=11, value=rec.get("工资标准", "")).alignment = align_center
        ws_flow.cell(row=r_idx, column=12, value=rec.get("实际出勤(天)", 0)).alignment = align_center
        
        c_exp = ws_flow.cell(row=r_idx, column=13, value=rec.get("扣减支出", 0))
        c_exp.alignment = align_right
        c_exp.number_format = "#,##0.00"
        
        c_bp = ws_flow.cell(row=r_idx, column=14, value=rec.get("实发基本工资", 0))
        c_bp.alignment = align_right
        c_bp.number_format = "#,##0.00"
        
        c_al = ws_flow.cell(row=r_idx, column=15, value=rec.get("津贴", 0))
        c_al.alignment = align_right
        c_al.number_format = "#,##0.00"
        
        c_tot = ws_flow.cell(row=r_idx, column=16, value=rec.get("实发总额(元)", 0))
        c_tot.alignment = align_right
        c_tot.number_format = "¥#,##0.00"
        c_tot.font = Font(name="微软雅黑", size=10, bold=True)
        
        c_bc = ws_flow.cell(row=r_idx, column=17, value=rec.get("实名制银行卡号", ""))
        c_bc.alignment = align_center
        c_bc.number_format = "@"
        
        ws_flow.cell(row=r_idx, column=18, value=rec.get("开户银行", "")).alignment = align_left
        ws_flow.cell(row=r_idx, column=19, value=rec.get("开户支行", "")).alignment = align_left

        for c_i in range(1, 20):
            cell = ws_flow.cell(row=r_idx, column=c_i)
            cell.border = cell_border
            if c_i != 10 and r_idx % 2 == 1:
                cell.fill = fill_zebra
            if c_i not in (10, 16):
                cell.font = font_data

    tot_flow_r = len(flow_records) + 2
    ws_flow.row_dimensions[tot_flow_r].height = 24
    ws_flow.merge_cells(start_row=tot_flow_r, start_column=1, end_row=tot_flow_r, end_column=11)
    c_tf = ws_flow.cell(row=tot_flow_r, column=1, value="流水全量合计：")
    c_tf.font = font_total
    c_tf.alignment = Alignment(horizontal="right", vertical="center")
    
    ws_flow.cell(row=tot_flow_r, column=12, value=f"=SUM(L2:L{tot_flow_r-1})").alignment = align_center
    ws_flow.cell(row=tot_flow_r, column=13, value=f"=SUM(M2:M{tot_flow_r-1})").number_format = "#,##0.00"
    ws_flow.cell(row=tot_flow_r, column=14, value=f"=SUM(N2:N{tot_flow_r-1})").number_format = "#,##0.00"
    ws_flow.cell(row=tot_flow_r, column=15, value=f"=SUM(O2:O{tot_flow_r-1})").number_format = "#,##0.00"
    c_s_tot = ws_flow.cell(row=tot_flow_r, column=16, value=f"=SUM(P2:P{tot_flow_r-1})")
    c_s_tot.number_format = "¥#,##0.00"

    for c_i in range(1, 20):
        cell = ws_flow.cell(row=tot_flow_r, column=c_i)
        cell.border = total_border
        cell.fill = fill_total
        cell.font = font_total

    ws_flow.freeze_panes = "G2"
    for c_i in range(1, 20):
        ws_flow.column_dimensions[get_column_letter(c_i)].width = 16
    ws_flow.column_dimensions["A"].width = 24
    ws_flow.column_dimensions["B"].width = 16
    ws_flow.column_dimensions["C"].width = 22
    ws_flow.column_dimensions["D"].width = 24
    ws_flow.column_dimensions["F"].width = 10
    ws_flow.column_dimensions["G"].width = 22
    ws_flow.column_dimensions["H"].width = 15
    ws_flow.column_dimensions["Q"].width = 24
    ws_flow.column_dimensions["R"].width = 26
    ws_flow.column_dimensions["S"].width = 30

    return wb


def export_ledger_to_folder(target_folder: str, filename: str = None, company_scope: str = None) -> tuple:
    """
    生成并保存指定分包商（或全盘）的 Excel 到用户指定的文件夹中
    """
    if not target_folder or not os.path.exists(target_folder):
        return False, "", f"目标文件夹不存在：{target_folder}"
    
    data = load_all_payroll_data(company_filter=company_scope)
    if not data["sorted_workers"]:
        return False, "", f"【{company_scope or '所选范围'}】当前没有任何批次数据，无法生成台账"
    
    if not filename:
        comp_tag = company_scope if company_scope and company_scope != "全部劳务公司" else "全部分包大盘"
        filename = f"科技文化中心项目_{comp_tag}_工人工资发放动态总台账.xlsx"
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
        
    full_path = os.path.join(target_folder, filename)
    try:
        wb = build_excel_workbook(data)
        wb.save(full_path)
        return True, full_path, ""
    except Exception as e:
        return False, "", f"保存文件失败：{str(e)}"


def get_excel_bytes(company_scope: str = None) -> bytes:
    """获取指定分包商（或全盘）台账 Excel 二进制流"""
    data = load_all_payroll_data(company_filter=company_scope)
    wb = build_excel_workbook(data)
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def open_folder_in_explorer(folder_path: str):
    """在 Windows 资源管理器中打开指定文件夹"""
    if os.path.exists(folder_path):
        os.system(f'explorer "{os.path.normpath(folder_path)}"')


def ask_directory_gui(initial_dir: str = "") -> str:
    """弹出 Windows 原生文件夹选择框"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = PROJECT_ROOT
        selected = filedialog.askdirectory(initialdir=initial_dir, title="请选择工资台账保存文件夹")
        root.destroy()
        return selected or ""
    except Exception as e:
        print(f"弹窗选择文件夹失败: {e}")
        return ""
