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
import json
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


def _load_master_personnel_map() -> dict:
    """从本地主数据库/状态中加载工人进退场时间及档案信息"""
    master_path = os.path.join(PROJECT_ROOT, "data", "master_state.json")
    if not os.path.exists(master_path):
        return {}
    try:
        with open(master_path, "r", encoding="utf-8") as f:
            d = json.load(f)
        mapping = {}
        for r in d.get("rows", []):
            cid = str(r.get("身份证号", "")).strip()
            if cid:
                mapping[cid] = {
                    "entry_date": str(r.get("进场日期", "") or "").strip(),
                    "exit_date": str(r.get("退场日期", "") or "").strip(),
                    "phone": str(r.get("手机号", "") or "").strip(),
                    "name": str(r.get("姓名", "") or "").strip(),
                    "status": str(r.get("在场/进退场状态", "") or "").strip()
                }
        return mapping
    except Exception as e:
        print(f"读取 master_state 失败: {e}")
        return {}


def build_excel_workbook(data: dict) -> openpyxl.Workbook:
    """
    根据汇总数据构建符合中建二局标准规范的【中建二局安装工程有限公司工资台账】Excel：
    1. 主表【工资台账】：严格按照 28 列两级表头与动态公式架构组织
       - 基础信息：序号、姓名、身份证号码、手机号码、进场时间 (A~E)
       - 应发工资（元）：2026年6月 ~ 2026年10月 (F~J)
       - 已发工资（元）：2026年6月 ~ 2026年10月 (K~O)
       - 其中：班组长代付（元）：2026年6月 ~ 2026年10月 (P~T)
       - 年度汇总：2026年应发工资、已发工资、班组长代付、未发工资、约付时间、退场时间、备注 (U~AB)
    2. 动态公式联动：
       - U列 (应发合计) = SUM(F{row}:J{row})
       - V列 (已发合计) = SUM(K{row}:O{row})
       - W列 (代付合计) = SUM(P{row}:T{row})
       - X列 (未发差额) = U{row}-V{row}
       - 底部合计行：整列 SUM
    3. 冻结 F4 窗格，固定表头与工人姓名/身份证
    4. 附表【流水明细库】：保留各批次原始流水，支持穿透核验
    """
    company_scope = data.get("company_scope", "全部劳务公司")
    sorted_workers = data.get("sorted_workers", [])
    flow_records = data["flow_df"].to_dict("records") if not data["flow_df"].empty else []
    master_map = _load_master_personnel_map()

    wb = openpyxl.Workbook()

    # 1. 确定台账大标题抬头
    if company_scope and company_scope not in ("全部劳务公司", "全部分包商", "全部", ""):
        comp_title = f"中建二局安装工程有限公司工资台账（{company_scope}）"
    else:
        comp_title = "中建二局安装工程有限公司工资台账"

    # 2. 基础样式定义
    font_main_title = Font(name="微软雅黑", size=16, bold=True, color="1F4E79")
    font_header = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="微软雅黑", size=10, color="000000")
    font_total = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")

    fill_title = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_h1_base = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")     # 海军蓝（基础信息）
    fill_h1_plan = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")     # 稳重蓝（应发）
    fill_h1_paid = PatternFill(start_color="165B33", end_color="165B33", fill_type="solid")     # 森林绿（已发）
    fill_h1_agent = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")    # 暖棕橙（代付）
    fill_h1_summary = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")  # 藏青（年度汇总）
    fill_h2 = PatternFill(start_color="333F48", end_color="333F48", fill_type="solid")          # 子表头深灰蓝
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")       # 合计浅蓝
    fill_zebra = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    double_border_bottom = Side(border_style="double", color="1F4E79")
    thick_border_top = Side(border_style="thin", color="1F4E79")

    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thick_border_top, bottom=double_border_bottom)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Sheet 1: 主表【工资台账】
    ws_main = wb.active
    ws_main.title = "工资台账"
    ws_main.views.sheetView[0].showGridLines = True

    # 3. 28 列规格定义
    month_labels = ["2026年6月", "2026年7月", "2026年8月", "2026年9月", "2026年10月"]
    month_keys = ["06", "07", "08", "09", "10"]

    columns_spec = []

    # 1. 基础信息列 (A~E)
    base_cols = [
        ("序号", 8, "center", "text"),
        ("姓名", 12, "center", "text"),
        ("身份证号码", 22, "center", "text"),
        ("手机号码", 15, "center", "text"),
        ("进场时间", 14, "center", "text"),
    ]
    for name, w, al, tp in base_cols:
        columns_spec.append({
            "key": name,
            "h1": name,
            "h2": "",
            "is_merged_v": True,
            "fill": fill_h1_base,
            "width": w,
            "align": al,
            "type": tp
        })

    # 2. 应发工资（元） (F~J)
    for idx_m, m_label in enumerate(month_labels):
        m_key = month_keys[idx_m]
        columns_spec.append({
            "key": f"应发_{m_key}",
            "h1": "应发工资（元）",
            "h2": m_label,
            "group": "应发工资（元）",
            "is_merged_v": False,
            "fill": fill_h1_plan,
            "width": 14,
            "align": "right",
            "type": "currency"
        })

    # 3. 已发工资（元） (K~O)
    for idx_m, m_label in enumerate(month_labels):
        m_key = month_keys[idx_m]
        columns_spec.append({
            "key": f"已发_{m_key}",
            "h1": "已发工资（元）",
            "h2": m_label,
            "group": "已发工资（元）",
            "is_merged_v": False,
            "fill": fill_h1_paid,
            "width": 14,
            "align": "right",
            "type": "currency"
        })

    # 4. 其中：班组长代付（元） (P~T)
    for idx_m, m_label in enumerate(month_labels):
        m_key = month_keys[idx_m]
        columns_spec.append({
            "key": f"代付_{m_key}",
            "h1": "其中：班组长代付（元）",
            "h2": m_label,
            "group": "其中：班组长代付（元）",
            "is_merged_v": False,
            "fill": fill_h1_agent,
            "width": 14,
            "align": "right",
            "type": "currency"
        })

    # 5. 年度汇总与状态 (U~AB)
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
            "key": s_name,
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
    ws_main.row_dimensions[1].height = 42
    ws_main.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws_main["A1"]
    title_cell.value = comp_title
    title_cell.font = font_main_title
    title_cell.alignment = align_center
    title_cell.fill = fill_title

    # ── 第 2 行 & 第 3 行：双层表头 ──
    ws_main.row_dimensions[2].height = 26
    ws_main.row_dimensions[3].height = 26

    for col_idx, col in enumerate(columns_spec, 1):
        c_letter = get_column_letter(col_idx)
        cell_h1 = ws_main.cell(row=2, column=col_idx)
        cell_h2 = ws_main.cell(row=3, column=col_idx)

        cell_h1.value = col["h1"]
        cell_h1.font = font_header
        cell_h1.alignment = align_center
        cell_h1.fill = col["fill"]
        cell_h1.border = cell_border

        cell_h2.value = col["h2"]
        cell_h2.font = font_sub_header
        cell_h2.alignment = align_center
        cell_h2.fill = fill_h2
        cell_h2.border = cell_border

        ws_main.column_dimensions[c_letter].width = col["width"]

    # 垂直合并（A~E 和 U~AB 列垂直合并 2~3 行）
    for col_idx, col in enumerate(columns_spec, 1):
        if col["is_merged_v"]:
            c_letter = get_column_letter(col_idx)
            ws_main.merge_cells(f"{c_letter}2:{c_letter}3")

    # 水平合并（F2:J2 应发工资、K2:O2 已发工资、P2:T2 班组长代付）
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
            ws_main.merge_cells(f"{s_let}2:{e_let}2")

    # ── 数据行填充 ──
    start_row = 4
    num_data_rows = max(len(sorted_workers), 78)

    for idx in range(num_data_rows):
        curr_row = start_row + idx
        seq = idx + 1
        ws_main.row_dimensions[curr_row].height = 22
        w = sorted_workers[idx] if idx < len(sorted_workers) else None

        # 统计工人各月份（06, 07, 08, 09, 10）的发放总额
        worker_batch_pays = w.get("batch_pays", {}) if w else {}
        worker_month_pays = {}
        for b_name, p_val in worker_batch_pays.items():
            m_found = re.search(r'(\d{2})月', b_name)
            if m_found:
                mm = m_found.group(1)
                worker_month_pays[mm] = worker_month_pays.get(mm, 0.0) + p_val

        cid = w["id_card"] if w else ""
        m_info = master_map.get(cid, {}) if cid else {}

        for col_idx, col in enumerate(columns_spec, 1):
            cell = ws_main.cell(row=curr_row, column=col_idx)
            cell.font = font_data
            cell.border = cell_border

            if seq % 2 == 0:
                cell.fill = fill_zebra

            col_align = col["align"]
            if col_align == "center":
                cell.alignment = align_center
            elif col_align == "right":
                cell.alignment = align_right
            else:
                cell.alignment = align_left

            col_key = col["key"]

            # 序号
            if col_key == "序号":
                cell.value = seq
                cell.number_format = "0"
                continue

            # 文本格式保护
            if col_key in ("身份证号码", "手机号码"):
                cell.number_format = "@"

            # 金额列格式
            if col["type"] == "currency":
                cell.number_format = "#,##0.00"

            # 动态公式列
            # U列: 2026年应发工资 = SUM(F{row}:J{row})
            if col_key == "2026年应发工资（元）":
                cell.value = f"=SUM(F{curr_row}:J{curr_row})"
                continue

            # V列: 2026年已发工资 = SUM(K{row}:O{row})
            if col_key == "2026年已发工资（元）":
                cell.value = f"=SUM(K{curr_row}:O{curr_row})"
                continue

            # W列: 2026年分包班组长自行代付工资 = SUM(P{row}:T{row})
            if col_key == "2026年分包班组长自行代付工资":
                cell.value = f"=SUM(P{curr_row}:T{curr_row})"
                continue

            # X列: 2026年未发工资 = U{row}-V{row} (应发 - 已发)
            if col_key == "2026年未发工资（元）":
                cell.value = f"=U{curr_row}-V{curr_row}"
                continue

            # 工人数据填充
            if w:
                if col_key == "姓名":
                    cell.value = w["name"]
                elif col_key == "身份证号码":
                    cell.value = w["id_card"]
                elif col_key == "手机号码":
                    cell.value = w.get("phone", "") or m_info.get("phone", "")
                elif col_key == "进场时间":
                    cell.value = m_info.get("entry_date", "")
                elif col_key == "人员退场时间":
                    cell.value = m_info.get("exit_date", "")
                elif col_key.startswith("已发_"):
                    mm = col_key.replace("已发_", "")
                    p = worker_month_pays.get(mm)
                    if p is not None:
                        cell.value = p
                elif col_key.startswith("应发_"):
                    mm = col_key.replace("应发_", "")
                    p = worker_month_pays.get(mm)
                    if p is not None:
                        cell.value = p

    # ── 底部【合计行】 ──
    total_row = start_row + num_data_rows
    ws_main.row_dimensions[total_row].height = 28

    # A~E 列合并为“合  计”
    ws_main.merge_cells(f"A{total_row}:E{total_row}")
    cell_tot_label = ws_main[f"A{total_row}"]
    cell_tot_label.value = "合  计"
    cell_tot_label.font = font_total
    cell_tot_label.alignment = align_center
    cell_tot_label.fill = fill_total

    for c in range(1, 6):
        ws_main.cell(row=total_row, column=c).border = total_border
        ws_main.cell(row=total_row, column=c).fill = fill_total

    for col_idx in range(6, total_cols + 1):
        col = columns_spec[col_idx - 1]
        c_letter = get_column_letter(col_idx)
        cell = ws_main.cell(row=total_row, column=col_idx)
        cell.font = font_total
        cell.border = total_border
        cell.fill = fill_total

        if col["type"] == "currency":
            cell.value = f"=SUM({c_letter}{start_row}:{c_letter}{total_row - 1})"
            cell.number_format = "#,##0.00"
            cell.alignment = align_right
        else:
            cell.alignment = align_center

    # 冻结窗格：固定前 3 行与前 5 列
    ws_main.freeze_panes = "F4"

    # Sheet 2: 附表【流水明细库】
    if flow_records:
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
            cell.fill = fill_h1_base
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
            ws_flow.cell(row=r_idx, column=10, value=rec.get("进退场状态", ""))
            ws_flow.cell(row=r_idx, column=10).alignment = align_center
            if rec.get("进退场状态") == "已进场":
                ws_flow.cell(row=r_idx, column=10).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                ws_flow.cell(row=r_idx, column=10).font = Font(name="微软雅黑", size=10, color="276A3C", bold=True)
            else:
                ws_flow.cell(row=r_idx, column=10).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                ws_flow.cell(row=r_idx, column=10).font = Font(name="微软雅黑", size=10, color="A61C1C", bold=True)

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
            c_tot.number_format = "#,##0.00"
            c_tot.font = Font(name="微软雅黑", size=10, bold=True)
            
            c_bc = ws_flow.cell(row=r_idx, column=17, value=rec.get("实名制银行卡号", ""))
            c_bc.alignment = align_center
            c_bc.number_format = "@"
            
            ws_flow.cell(row=r_idx, column=18, value=rec.get("开户银行", "")).alignment = align_left
            ws_flow.cell(row=r_idx, column=19, value=rec.get("开户支行", "")).alignment = align_left

            for c_i in range(1, 20):
                cell = ws_flow.cell(row=r_idx, column=c_i)
                cell.border = cell_border
                if r_idx % 2 == 1:
                    cell.fill = fill_zebra

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
        ws_flow.cell(row=tot_flow_r, column=16, value=f"=SUM(P2:P{tot_flow_r-1})").number_format = "#,##0.00"

        for c_i in range(1, 20):
            cell = ws_flow.cell(row=tot_flow_r, column=c_i)
            cell.border = total_border
            cell.fill = fill_total
            cell.font = font_total

        ws_flow.freeze_panes = "G2"

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
