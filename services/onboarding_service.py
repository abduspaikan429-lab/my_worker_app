from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Union

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
        """Read and return all onboarding records from JSON file."""
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
        """Get onboarding records for workers who are not 100% complete."""
        records = self.get_records()
        pending = {}
        for record_id, record_data in records.items():
            completed, total = self.get_progress(record_data)
            if completed < total:
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

        if not name:
            raise ValueError(
                "Worker name is required for creating onboarding record"
            )

        record_id = f"{name}_{team}"
        records = self.get_records()

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
            }
            records[record_id] = new_record
            self.save_records(records)
            return new_record

        return records[record_id]

    def complete_onboarding(self, record_id: str) -> None:
        """Mark all onboarding items for record_id as 100% complete."""
        records = self.get_records()
        if record_id in records:
            record = records[record_id]
            record["paper"] = {k: True for k in PAPER_ITEMS}
            record["system"] = {k: True for k in SYSTEM_ITEMS}
            record["access"] = {k: True for k in ACCESS_ITEMS}
            self.save_records(records)
