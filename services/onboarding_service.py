from __future__ import annotations

from datetime import date, datetime
import json
from pathlib import Path
from typing import Any, Dict, Union

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_ONBOARDING_FILE = BASE_DIR / "data" / "onboarding_data.json"

PAPER_ITEMS = [
    "体检单",
    "三级教育",
    "承诺书",
    "岗前培训",
    "签到表按手印(2张)",
    "花名册",
    "劳动合同(纸质)",
    "进场告知书(纸质)",
]
SYSTEM_ITEMS = ["更新花名册", "更新月更报表", "更新签到表"]
ACCESS_ITEMS = [
    "门禁录入完成",
    "百工聚合同/告知书签署及加卡",
    "智慧护薪合同发起及工人/班组长确认",
]
TOTAL_ITEMS = len(PAPER_ITEMS) + len(SYSTEM_ITEMS) + len(ACCESS_ITEMS)


class OnboardingService:
    """Service layer handling worker onboarding pipeline business logic."""

    def __init__(self, file_path: Union[str, Path] = DEFAULT_ONBOARDING_FILE):
        self.file_path = Path(file_path)

    def get_records(self) -> Dict[str, Dict[str, Any]]:
        """Read and return all onboarding records from JSON file or session_state."""
        try:
            import streamlit as st
            if hasattr(st, "session_state") and "onboarding_data" in st.session_state:
                if isinstance(st.session_state.onboarding_data, dict):
                    return st.session_state.onboarding_data
        except Exception:
            pass
        if not self.file_path.exists():
            return {}
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def save_records(self, records: Dict[str, Dict[str, Any]]) -> None:
        """Save onboarding records dictionary to JSON file."""
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

    def get_progress(self, worker_data: Dict[str, Any]) -> tuple[int, int]:
        """Calculate completed items count and total items count for a record."""
        completed = 0
        completed += sum(1 for v in worker_data.get("paper", {}).values() if v)
        completed += sum(1 for v in worker_data.get("system", {}).values() if v)
        completed += sum(1 for v in worker_data.get("access", {}).values() if v)
        return completed, TOTAL_ITEMS

    def get_pending_workers(self) -> Dict[str, Dict[str, Any]]:
        """Get onboarding records for workers who are not marked as completed."""
        records = self.get_records()
        pending = {}
        for record_id, record_data in records.items():
            if record_data.get("status") != "completed":
                pending[record_id] = record_data
        return pending

    def create_onboarding(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a new worker onboarding record."""
        if not isinstance(data, dict):
            raise ValueError("Data must be a dictionary")

        info = data.get("info", data)
        name = str(info.get("姓名") or info.get("name") or "").strip()
        team = (
            str(info.get("班组") or info.get("team") or "").strip()
            or "待分配班组"
        )

        sid = str(info.get("身份证号") or info.get("id_card") or "").strip()
        phone = str(info.get("手机号") or info.get("phone") or "").strip()

        if not name:
            raise ValueError(
                "Worker name is required for creating onboarding record"
            )

        records = self.get_records()
        record_id = None
        
        if sid:
            for k, v in records.items():
                if v.get("info", {}).get("身份证号") == sid:
                    record_id = k
                    break
        if not record_id and phone:
            for k, v in records.items():
                i = v.get("info", {})
                if i.get("姓名") == name and i.get("手机号") == phone:
                    record_id = k
                    break
        if not record_id:
            for k, v in records.items():
                i = v.get("info", {})
                if i.get("姓名") == name and i.get("班组") == team:
                    record_id = k
                    break
        
        if not record_id:
            if sid:
                record_id = sid
            elif phone:
                record_id = f"{name}_{phone}"
            else:
                record_id = f"{name}_{team}"

        if record_id not in records:
            new_record = {
                "info": {
                    "姓名": name,
                    "班组": team,
                    "身份证号": str(
                        info.get("身份证号") or info.get("id_card") or ""
                    ).strip(),
                    "手机号": str(
                        info.get("手机号") or info.get("phone") or ""
                    ).strip(),
                    "工种": str(
                        info.get("工种") or info.get("job_type") or ""
                    ).strip(),
                    "银行卡号": str(
                        info.get("银行卡号") or info.get("bank_card") or ""
                    ).strip(),
                    "进场日期": str(
                        info.get("进场日期") or data.get("created_at") or date.today()
                    ),
                },
                "paper": {
                    k: data.get("paper", {}).get(k, False) for k in PAPER_ITEMS
                },
                "system": {
                    k: data.get("system", {}).get(k, False)
                    for k in SYSTEM_ITEMS
                },
                "access": {
                    k: data.get("access", {}).get(k, False)
                    for k in ACCESS_ITEMS
                },
                "created_at": str(data.get("created_at") or date.today()),
            }
            records[record_id] = new_record
            self.save_records(records)
            return new_record

        return records[record_id]

    def mark_completed(self, record_id: str, is_completed: bool = True) -> None:
        """Mark an onboarding record as completed/active without deleting it."""
        records = self.get_records()
        if record_id in records:
            records[record_id]["status"] = "completed" if is_completed else "active"
            if is_completed:
                records[record_id]["completed_at"] = str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            self.save_records(records)

    def complete_onboarding(self, record_id: str) -> None:
        """Complete onboarding record (alias for mark_completed)."""
        self.mark_completed(record_id, True)

    def get_onboarding_df(self) -> pd.DataFrame:
        """
        将进场流程中（进场流水线）的人员记录转换为标准 DataFrame。
        """
        records = self.get_records()
        if not records:
            return pd.DataFrame()
        rows = []
        for worker_id, data in records.items():
            if not isinstance(data, dict):
                continue
            info = data.get("info", {})
            name = str(info.get("姓名") or info.get("name") or "").strip()
            team = str(info.get("班组") or info.get("team") or "").strip() or "待分配班组"
            id_card = str(info.get("身份证号") or info.get("id_card") or "").strip()
            phone = str(info.get("手机号") or info.get("phone") or "").strip()
            job_type = str(info.get("工种") or info.get("job_type") or "").strip()
            bank_card = str(info.get("银行卡号") or info.get("bank_card") or "").strip()
            entry_date = str(info.get("进场日期") or data.get("created_at") or date.today())

            row = {
                "姓名": name,
                "班组": team,
                "身份证号": id_card,
                "手机号": phone,
                "工种": job_type,
                "银行卡号": bank_card,
                "进场日期": entry_date,
                "进场时间": entry_date,
                "在场/进退场状态": "进场手续中",
            }
            rows.append(row)
        return pd.DataFrame(rows)

    def merge_with_master(self, master_df: pd.DataFrame | None) -> pd.DataFrame:
        """
        将主表与进场流程中活跃人员合并（只要添加到进场流程中即视为进场）。
        自动去重，防止主表已有人员重复叠加。
        """
        onboarding_df = self.get_onboarding_df()
        if master_df is None or master_df.empty:
            return onboarding_df
        if onboarding_df.empty:
            return master_df.copy()

        # 收集主表中已有的唯一标识
        master_ids = set()
        master_names = set()
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

        # 找出未在主表中的进场流水线人员并追加
        new_rows = []
        for _, r in onboarding_df.iterrows():
            sid = str(r.get("身份证号", "")).strip()
            nm = str(r.get("姓名", "")).strip()
            tm = str(r.get("班组", "")).strip()
            combo = f"{nm}_{tm}"

            if sid and sid in master_ids:
                continue
            if not sid and (combo in master_names or (not tm and nm in master_names)):
                continue
            new_rows.append(r.to_dict())

        if not new_rows:
            return master_df.copy()

        return pd.concat([master_df, pd.DataFrame(new_rows)], ignore_index=True)

