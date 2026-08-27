# services/task_engine.py
from datetime import date
from typing import Any, Dict, List, Tuple
from modules.onboarding_pipeline import PAPER_ITEMS as ON_PAPER, SYSTEM_ITEMS as ON_SYSTEM, ACCESS_ITEMS as ON_ACCESS

# Offboarding constants to be synced with offboarding_service
OFF_STEPS = [
    "1. 工人小灵光发起",
    "2. 班组长确认",
    "3. 劳资员（我）确认",
    "4. 提交财务发放",
    "5. 财务发放完成",
]
OFF_PAPER = ["收纸质离场结算单", "归档身份证+结算单照片", "结清证明上传"]
OFF_SYSTEM = ["处理离场月报", "更新花名册", "更新签到表"]

def _is_all_true(d: dict, keys: list) -> bool:
    return all(d.get(k, False) for k in keys)

def get_onboarding_status(worker_id: str, data: dict, master_ids: set, master_names: set) -> dict:
    """计算单个进场人员的下一步状态和异常。"""
    info = data.get("info", {})
    paper = data.get("paper", {})
    system = data.get("system", {})
    access = data.get("access", {})
    status = data.get("status", "active")
    
    name = str(info.get("姓名", "未知")).strip()
    team = str(info.get("班组", "未知")).strip()
    sid = str(info.get("身份证号", "")).strip()
    
    in_master = (sid and sid in master_ids) or (f"{name}_{team}" in master_names) or (name in master_names)

    # Calculate completed vs total items for general progress
    total_items = len(ON_PAPER) + len(ON_SYSTEM) + len(ON_ACCESS) + 1 # +1 for master sync
    completed_items = sum(1 for v in paper.values() if v) + sum(1 for v in system.values() if v) + sum(1 for v in access.values() if v)
    if in_master:
        completed_items += 1
    
    if status == "completed" or completed_items == total_items:
        return {"worker_id": worker_id, "name": name, "team": team, "type": "onboarding", "category": "completed", "status": "正常在场", "action": "无", "anomaly": None}

    anomaly = None
    if in_master and not _is_all_true(paper, ON_PAPER):
        anomaly = "官方数据已同步，但纸质资料仍未补齐"
    elif in_master and not _is_all_true(access, ON_ACCESS):
        anomaly = "官方数据已同步，但门禁/小灵光手续未全部完成"
    elif _is_all_true(system, ON_SYSTEM) and not in_master:
        anomaly = "本地台账已更新，但未执行官方数据导出同步"

    category = ""
    action = ""
    status_label = "进场办理中"

    # Rule Engine: Give ONLY ONE next step
    if not _is_all_true(paper, ON_PAPER):
        missing = [k for k in ON_PAPER if not paper.get(k, False)]
        category = "red"
        action = f"补齐资料: {missing[0]}"
    elif not access.get("门禁录入完成", False):
        category = "orange"
        action = "等待总包录入门禁"
    elif not access.get("百工聚合同/告知书签署及加卡", False):
        category = "orange"
        action = "等待工人百工聚签合同及加卡"
    elif not access.get("智慧护薪合同发起及工人/班组长确认", False):
        category = "orange"
        action = "等待工人/班组长智慧护薪确认"
    elif not in_master:
        category = "yellow"
        action = "导出三局和智慧护薪数据导入系统"
        status_label = "等待官方数据同步"
    elif not _is_all_true(system, ON_SYSTEM):
        missing = [k for k in ON_SYSTEM if not system.get(k, False)]
        category = "red"
        action = f"同步本地台账: {missing[0]}"
        status_label = "台账同步中"
    else:
        category = "completed"
        action = "无"
        status_label = "正常在场"

    return {
        "worker_id": worker_id, "name": name, "team": team, 
        "type": "onboarding", "category": category, 
        "status": status_label, "action": action, "anomaly": anomaly
    }

def get_offboarding_status(worker_id: str, data: dict) -> dict:
    info = data.get("info", {})
    steps = data.get("steps", {})
    paper = data.get("paper", {})
    system = data.get("system", {})
    status_flag = data.get("status", "active")
    
    name = str(info.get("姓名", "未知")).strip()
    team = str(info.get("班组", "未知")).strip()

    total_items = len(OFF_STEPS) + len(OFF_PAPER) + len(OFF_SYSTEM)
    completed_items = sum(1 for v in steps.values() if v) + sum(1 for v in paper.values() if v) + sum(1 for v in system.values() if v)
    
    if status_flag == "completed" or completed_items == total_items:
        return {"worker_id": worker_id, "name": name, "team": team, "type": "offboarding", "category": "completed", "status": "历史归档", "action": "无", "anomaly": None}

    anomaly = None
    if steps.get("5. 财务发放完成", False) and not paper.get("收纸质离场结算单", False):
        anomaly = "工资已发，但尚未收到纸质离场结算单！"
    elif steps.get("5. 财务发放完成", False) and not steps.get("3. 劳资员（我）确认", False):
        anomaly = "工资已发，但我尚未在小灵光确认"

    category = ""
    action = ""
    status_label = "离场办理中"

    # Ordered rules
    if not steps.get("1. 工人小灵光发起", False):
        category = "orange"
        action = "等待工人小灵光发起离场"
    elif not paper.get("收纸质离场结算单", False):
        category = "red"
        action = "收纸质离场结算单"
    elif not paper.get("归档身份证+结算单照片", False):
        category = "red"
        action = "归档身份证+离场结算单照片"
    elif not steps.get("2. 班组长确认", False):
        category = "orange"
        action = "等待班组长小灵光确认"
    elif not steps.get("3. 劳资员（我）确认", False):
        category = "red"
        action = "需要在小灵光进行劳资员确认"
    elif not steps.get("4. 提交财务发放", False):
        category = "red"
        action = "提交财务发放工资"
    elif not steps.get("5. 财务发放完成", False):
        category = "orange"
        action = "等待财务打款完成"
    elif not paper.get("结清证明上传", False):
        category = "red"
        action = "上传结清证明"
        status_label = "等待结清证明"
    elif not system.get("处理离场月报", False):
        category = "red"
        action = "处理离场月报"
        status_label = "等待台账同步"
    elif not system.get("更新花名册", False):
        category = "red"
        action = "更新花名册"
        status_label = "等待台账同步"
    elif not system.get("更新签到表", False):
        category = "red"
        action = "更新签到表"
        status_label = "等待台账同步"
    else:
        category = "completed"
        action = "无"
        status_label = "历史归档"

    return {
        "worker_id": worker_id, "name": name, "team": team, 
        "type": "offboarding", "category": category, 
        "status": status_label, "action": action, "anomaly": anomaly
    }

def get_all_tasks(onboarding_data: dict, offboarding_data: dict, master_df) -> Tuple[list, list, list, list]:
    tasks_red = []
    tasks_orange = []
    tasks_yellow = []
    anomalies = []

    # Get master info
    master_ids = set()
    master_names = set()
    if master_df is not None and not master_df.empty:
        id_col = next((c for c in ["身份证号", "id_card"] if c in master_df.columns), None)
        if id_col:
            master_ids = {str(x).strip() for x in master_df[id_col].dropna() if str(x).strip()}
        name_col = next((c for c in ["姓名", "name"] if c in master_df.columns), None)
        team_col = next((c for c in ["班组", "team", "工种", "job_type"] if c in master_df.columns), None)
        if name_col:
            for _, r in master_df.iterrows():
                nm = str(r.get(name_col, "")).strip()
                tm = str(r.get(team_col, "")).strip() if team_col else ""
                if nm:
                    master_names.add(nm)
                    if tm:
                        master_names.add(f"{nm}_{tm}")

    for worker_id, data in onboarding_data.items():
        if data.get("status") == "completed":
            continue
            
        task = get_onboarding_status(worker_id, data, master_ids, master_names)
        if task["anomaly"]:
            anomalies.append(task)
            
        if task["category"] == "red":
            tasks_red.append(task)
        elif task["category"] == "orange":
            tasks_orange.append(task)
        elif task["category"] == "yellow":
            tasks_yellow.append(task)

    for worker_id, data in offboarding_data.items():
        if data.get("status") == "completed":
            continue
            
        task = get_offboarding_status(worker_id, data)
        if task["anomaly"]:
            anomalies.append(task)
            
        if task["category"] == "red":
            tasks_red.append(task)
        elif task["category"] == "orange":
            tasks_orange.append(task)
        elif task["category"] == "yellow":
            tasks_yellow.append(task)

    return tasks_red, tasks_orange, tasks_yellow, anomalies
