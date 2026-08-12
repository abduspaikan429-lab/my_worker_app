from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

from repositories.worker_repository import WorkerRepository

ALIAS_MAP = {
    "name": "姓名",
    "id_card": "身份证号",
    "phone": "手机号",
    "bank_card": "工资卡号",
    "job_type": "工种",
    "status": "在场/进退场状态",
    "entry_date": "进场日期",
    "exit_date": "退场日期",
    "company_name": "分包/所属企业",
    "team_name": "班组",
}

REVERSE_ALIAS_MAP = {v: k for k, v in ALIAS_MAP.items()}


def _sync_worker_aliases(worker: Dict[str, Any]) -> Dict[str, Any]:
    """Sync English and Chinese alias fields in worker dictionary."""
    for eng_key, chi_key in ALIAS_MAP.items():
        if eng_key in worker and chi_key not in worker:
            worker[chi_key] = worker[eng_key]
        elif chi_key in worker and eng_key not in worker:
            worker[eng_key] = worker[chi_key]
    return worker


class WorkerService:
    """Service layer for Worker business logic and repository interactions."""

    def __init__(self, repository: Optional[WorkerRepository] = None):
        if repository is None:
            from core.repository_factory import get_worker_repository
            repository = get_worker_repository()
        self.repository = repository

    def get_workers(self) -> List[Dict[str, Any]]:
        """Retrieve all workers."""
        return self.repository.get_all_workers()

    def get_worker_by_id_card(self, id_card: str) -> Optional[Dict[str, Any]]:
        """Query worker by ID card number."""
        if not id_card:
            return None
        return self.repository.get_worker_by_id_card(id_card)

    def create_worker(self, worker: Dict[str, Any]) -> Dict[str, Any]:
        """Validate and create a new worker record."""
        if not isinstance(worker, dict):
            raise ValueError("Worker data must be a dictionary")

        id_card = worker.get("id_card") or worker.get("身份证号")
        name = worker.get("name") or worker.get("姓名")

        if not id_card or not str(id_card).strip():
            raise ValueError("id_card (身份证号) is required for creating worker")
        if not name or not str(name).strip():
            raise ValueError("name (姓名) is required for creating worker")

        id_card = str(id_card).strip()
        existing = self.get_worker_by_id_card(id_card)
        if existing:
            raise ValueError(f"Worker with id_card '{id_card}' already exists")

        _sync_worker_aliases(worker)

        workers = self.get_workers()
        workers.append(worker)
        self.repository.save_workers(workers)
        return worker

    def update_worker(
        self, worker_key: Any, data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Validate and update an existing worker record."""
        if not isinstance(data, dict):
            raise ValueError("Update data must be a dictionary")

        key_str = str(worker_key).strip()
        workers = self.get_workers()

        target_index = -1
        for idx, w in enumerate(workers):
            if isinstance(w, dict):
                card = str(w.get("id_card") or w.get("身份证号") or "").strip()
                w_id = str(w.get("id") or "").strip()
                if card == key_str or w_id == key_str:
                    target_index = idx
                    break

        if target_index == -1:
            raise ValueError(f"Worker '{worker_key}' not found")

        target = workers[target_index]
        for k, v in data.items():
            target[k] = v
            if k in ALIAS_MAP:
                target[ALIAS_MAP[k]] = v
            if k in REVERSE_ALIAS_MAP:
                target[REVERSE_ALIAS_MAP[k]] = v

        self.repository.save_workers(workers)
        return target

    def save_workers(self, workers: List[Dict[str, Any]]) -> None:
        """Batch save workers list."""
        self.repository.save_workers(workers)
