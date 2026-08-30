from __future__ import annotations

from datetime import date, datetime
import json
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple, Union

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OFFBOARDING_FILE = BASE_DIR / "data" / "offboarding_data.json"
DEFAULT_HISTORY_FILE = BASE_DIR / "data" / "offboarding_history.json"

OFFBOARDING_STEPS = [
    "1. 工人在百工聚发起离场",
    "2. 班组长确认",
    "3. 劳资员（我）确认",
    "4. 财务发放完成",
]
OFF_PAPER = ["收纸质离场结算单", "归档身份证+结算单照片", "结清证明上传"]
OFF_SYSTEM = ["处理离场月报", "更新花名册", "更新签到表"]
TOTAL_ITEMS = len(OFFBOARDING_STEPS) + len(OFF_PAPER) + len(OFF_SYSTEM)


class OffboardingService:
    """Service layer handling offboarding pipeline business logic."""

    def __init__(
        self,
        file_path: Union[str, Path] = DEFAULT_OFFBOARDING_FILE,
        history_path: Union[str, Path] = DEFAULT_HISTORY_FILE,
    ):
        self.file_path = Path(file_path)
        self.history_path = Path(history_path)

    def get_records(self) -> Dict[str, Dict[str, Any]]:
        """Read and return all active offboarding records."""
        try:
            import streamlit as st
            if hasattr(st, "session_state") and "offboarding_data" in st.session_state:
                if isinstance(st.session_state.offboarding_data, dict):
                    return st.session_state.offboarding_data
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
        """Save offboarding records dictionary to JSON file."""
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

    def get_progress(self, worker_data: Dict[str, Any]) -> tuple[int, int]:
        """Calculate completed steps count and total count."""
        steps = worker_data.get("steps", {})
        paper = worker_data.get("paper", {})
        system = worker_data.get("system", {})
        completed = sum(1 for v in steps.values() if v) + sum(1 for v in paper.values() if v) + sum(1 for v in system.values() if v)
        return completed, TOTAL_ITEMS

    def get_pending_workers(self) -> Dict[str, Dict[str, Any]]:
        """Get offboarding records for workers who have not completed all steps."""
        records = self.get_records()
        pending = {}
        for record_id, record_data in records.items():
            if record_data.get("status") != "completed":
                pending[record_id] = record_data
        return pending

    def create_offboarding(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a new worker offboarding record."""
        if not isinstance(data, dict):
            raise ValueError("Data must be a dictionary")

        info = data.get("info", data)
        name = str(info.get("姓名") or info.get("name") or "").strip()
        team = str(info.get("班组") or info.get("team") or "").strip() or "待分配班组"
        source_id = str(info.get("身份证号") or info.get("id_card") or "").strip()
        phone = str(info.get("手机号") or info.get("phone") or "").strip()

        if not name:
            raise ValueError("Worker name is required for offboarding")

        records = self.get_records()
        record_id = None
        
        if source_id:
            for k, v in records.items():
                if v.get("info", {}).get("身份证号") == source_id:
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
            if source_id:
                record_id = source_id
            elif phone:
                record_id = f"{name}_{phone}"
            else:
                record_id = f"{name}_{team}"

        if record_id not in records:
            new_record = {
                "info": {
                    "姓名": name,
                    "班组": team,
                    "身份证号": source_id,
                    "手机号": phone,
                    "工种": str(info.get("工种") or info.get("job_type") or "").strip(),
                    "离场日期": str(info.get("离场日期") or date.today()),
                },
                "steps": {k: data.get("steps", {}).get(k, False) for k in OFFBOARDING_STEPS},
                "paper": {k: data.get("paper", {}).get(k, False) for k in OFF_PAPER},
                "system": {k: data.get("system", {}).get(k, False) for k in OFF_SYSTEM},
                "created_at": str(data.get("created_at") or date.today()),
            }
            records[record_id] = new_record
            self.save_records(records)
            return new_record

        return records[record_id]

    def mark_completed(self, record_id: str, is_completed: bool = True) -> None:
        """Mark an offboarding record as completed/active without deleting it."""
        records = self.get_records()
        if record_id in records:
            records[record_id]["status"] = "completed" if is_completed else "active"
            if is_completed:
                records[record_id]["completed_at"] = str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                self.archive_offboarding(record_id, records[record_id])
            self.save_records(records)

    # Remove update_status

    def load_history(self) -> List[Dict[str, Any]]:
        """Load offboarding history records."""
        if not self.history_path.exists():
            return []
        try:
            with open(self.history_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            return []

    def archive_offboarding(self, worker_id: str, data: Dict[str, Any]) -> None:
        """Archive completed offboarding worker record."""
        history = self.load_history()
        info = data.get("info", {})
        record = {
            "姓名": info.get("姓名", ""),
            "班组": info.get("班组", ""),
            "身份证号": info.get("身份证号", ""),
            "离场日期": str(date.today()),
        }
        key = record["身份证号"] or f"{record['姓名']}_{record['班组']}"
        existing_keys = {
            r.get("身份证号") or f"{r.get('姓名','')}_{r.get('班组','')}"
            for r in history
        }
        if key not in existing_keys:
            history.append(record)
            self.history_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.history_path, "w", encoding="utf-8") as f:
                json.dump(history, f, ensure_ascii=False, indent=2)

    def get_offboarded_identifiers(self, include_active: bool = True) -> Tuple[Set[str], Set[str]]:
        """
        获取已离场/离场结算中的人员唯一标识。
        :param include_active: 若为 True，则包含正在办理离场结算板块中的人员（即只要加入离场板块即扣减在场）。
        :return: (left_ids: 身份证集合, left_names: 姓名/姓名_班组集合)
        """
        left_ids: Set[str] = set()
        left_names: Set[str] = set()

        # 1. 历史归档已办结离场人员
        for rec in self.load_history():
            sid = str(rec.get("身份证号") or "").strip()
            nm = str(rec.get("姓名") or "").strip()
            tm = str(rec.get("班组") or "").strip()
            if sid:
                left_ids.add(sid)
            if nm:
                left_names.add(nm)
                if tm:
                    left_names.add(f"{nm}_{tm}")

        # 2. 正在离场结算中（活跃离场板块）的人员
        if include_active:
            for rec in self.get_records().values():
                info = rec.get("info", {}) if isinstance(rec, dict) else {}
                sid = str(info.get("身份证号") or info.get("id_card") or "").strip()
                nm = str(info.get("姓名") or info.get("name") or "").strip()
                tm = str(info.get("班组") or info.get("team") or "").strip()
                if sid:
                    left_ids.add(sid)
                if nm:
                    left_names.add(nm)
                    if tm:
                        left_names.add(f"{nm}_{tm}")

        return left_ids, left_names

    def filter_onsite_df(self, df: pd.DataFrame, include_active: bool = True) -> pd.DataFrame:
        """
        过滤人员 DataFrame，返回当前在场人员名单（自动剔除已归档和正在离场结算板块中的人员）。
        """
        if df is None or df.empty:
            return pd.DataFrame() if df is None else df.copy()

        left_ids, left_names = self.get_offboarded_identifiers(include_active=include_active)
        if not left_ids and not left_names:
            return df.copy()

        mask_out = pd.Series(False, index=df.index)

        # 优先通过身份证号比对
        id_col = next((c for c in ["身份证号", "id_card"] if c in df.columns), None)
        if id_col and left_ids:
            clean_ids = df[id_col].astype(str).str.strip()
            mask_out |= clean_ids.isin(left_ids)

        # 次选通过姓名/班组比对（针对未填身份证或主表无身份证的情况）
        name_col = next((c for c in ["姓名", "name"] if c in df.columns), None)
        team_col = next((c for c in ["班组", "team", "工种", "job_type"] if c in df.columns), None)

        if name_col and left_names:
            names = df[name_col].astype(str).str.strip()
            if team_col:
                teams = df[team_col].astype(str).str.strip()
                combo = names + "_" + teams
                mask_out |= combo.isin(left_names)

            # 单纯姓名匹配（主要处理班组为空/待分配班组的情况）
            pure_left_names = {k for k in left_names if "_" not in k}
            if pure_left_names:
                if team_col:
                    team_empty = df[team_col].astype(str).str.strip().isin(["", "nan", "None", "待分配班组", "未分组", "未知"])
                    mask_out |= (names.isin(pure_left_names) & team_empty)
                else:
                    mask_out |= names.isin(pure_left_names)

        return df[~mask_out].copy()

