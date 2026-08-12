from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from repositories.sqlite_worker_repository import SqliteWorkerRepository
from repositories.worker_repository import JsonWorkerRepository, WorkerRepository
from services.sync_log_service import DEFAULT_LOG_PATH, write_sync_log
from services.worker_compare_service import WorkerCompareService
from services.worker_sync_service import WorkerSyncService


class ShadowWorkerRepository(WorkerRepository):
    """Shadow Worker Repository implementation.

    Primary storage: JsonWorkerRepository (all reads & primary writes)
    Shadow storage: SqliteWorkerRepository (background sync mirror & verification)
    """

    def __init__(
        self,
        json_repo: Optional[JsonWorkerRepository] = None,
        sqlite_repo: Optional[SqliteWorkerRepository] = None,
        log_path: Union[str, Path] = DEFAULT_LOG_PATH,
    ):
        self.primary_repo = json_repo or JsonWorkerRepository()
        self.shadow_repo = sqlite_repo or SqliteWorkerRepository()
        self.log_path = Path(log_path)

    def get_all_workers(self) -> List[Dict[str, Any]]:
        """Always retrieve workers from primary JSON repository."""
        return self.primary_repo.get_all_workers()

    def get_worker_by_id_card(self, id_card: str) -> Optional[Dict[str, Any]]:
        """Always search worker from primary JSON repository."""
        return self.primary_repo.get_worker_by_id_card(id_card)

    def save_workers(self, workers: List[Dict[str, Any]]) -> None:
        """Save workers to primary JSON repository first, then shadow-sync to SQLite."""
        # 1. Primary save to JSON (Must succeed or raise exception)
        self.primary_repo.save_workers(workers)

        # 2. Shadow sync to SQLite & verify (Errors logged without interrupting JSON save)
        try:
            sync_res = WorkerSyncService.sync_workers(workers, self.shadow_repo)
            sqlite_workers = self.shadow_repo.get_all_workers()
            compare_res = WorkerCompareService.compare_workers(workers, sqlite_workers)

            if sync_res.get("success") and compare_res.get("result") == "PASS":
                msg = f"[INFO] Shadow sync SUCCESS: JSON count = {sync_res['json_count']}, SQLite count = {sync_res['sqlite_count']}"
                write_sync_log(msg, log_path=self.log_path)
            else:
                msg = f"[WARN] Shadow sync FAIL: JSON count = {sync_res['json_count']}, SQLite count = {sync_res['sqlite_count']}, Compare = {compare_res['result']}"
                write_sync_log(msg, log_path=self.log_path)
        except Exception as e:
            error_msg = f"[ERROR] Shadow SQLite sync exception: {str(e)}"
            write_sync_log(error_msg, log_path=self.log_path)
