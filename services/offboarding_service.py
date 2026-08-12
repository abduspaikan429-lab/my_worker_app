from __future__ import annotations

from datetime import date
import json
from pathlib import Path
from typing import Any, Dict, List, Union

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_OFFBOARDING_FILE = BASE_DIR / "data" / "offboarding_data.json"
DEFAULT_HISTORY_FILE = BASE_DIR / "data" / "offboarding_history.json"

OFFBOARDING_STEPS = [
    "1. 工人小灵光发起",
    "2. 班组长确认",
    "3. 劳资员确认",
    "4. 劳资员提交财务发放",
    "5. 财务发放完成",
]
TOTAL_ITEMS = len(OFFBOARDING_STEPS)


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
        completed = sum(1 for v in steps.values() if v)
        return completed, TOTAL_ITEMS

    def get_pending_workers(self) -> Dict[str, Dict[str, Any]]:
        """Get offboarding records for workers who have not completed all steps."""
        records = self.get_records()
        pending = {}
        for record_id, record_data in records.items():
            completed, total = self.get_progress(record_data)
            if completed < total:
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

        if not name:
            raise ValueError("Worker name is required for offboarding")

        record_id = f"{name}_{team}"
        records = self.get_records()

        if record_id not in records:
            new_record = {
                "info": {
                    "姓名": name,
                    "班组": team,
                    "身份证号": source_id,
                    "手机号": str(info.get("手机号") or info.get("phone") or "").strip(),
                    "工种": str(info.get("工种") or info.get("job_type") or "").strip(),
                },
                "steps": {k: data.get("steps", {}).get(k, False) for k in OFFBOARDING_STEPS},
            }
            records[record_id] = new_record
            self.save_records(records)
            return new_record

        return records[record_id]

    def update_status(self, record_id: str, step_updates: Dict[str, bool]) -> Dict[str, Any]:
        """Update step completion status for offboarding record."""
        records = self.get_records()
        if record_id not in records:
            raise ValueError(f"Offboarding record '{record_id}' not found")

        record = records[record_id]
        if "steps" not in record:
            record["steps"] = {k: False for k in OFFBOARDING_STEPS}

        record["steps"].update(step_updates)
        self.save_records(records)

        # Auto-archive if 100% complete
        completed, total = self.get_progress(record)
        if completed == total:
            self.archive_offboarding(record_id, record)

        return record

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
