from __future__ import annotations

from typing import Any, Dict, List

from repositories.sqlite_worker_repository import SqliteWorkerRepository


class WorkerSyncService:
    """Service to synchronize JSON worker data to SQLite mirror database."""

    @staticmethod
    def sync_workers(
        json_workers: List[Dict[str, Any]],
        sqlite_repository: SqliteWorkerRepository,
    ) -> Dict[str, Any]:
        """Clear SQLite workers table, write all JSON workers, and verify counts."""

        # Count unique valid JSON workers (by non-empty id_card)
        valid_json_cards = set()
        for w in json_workers:
            if isinstance(w, dict):
                card = (w.get("id_card") or w.get("身份证号") or "").strip()
                if card:
                    valid_json_cards.add(card)

        expected_json_count = len(valid_json_cards)

        # 1. Clear SQLite workers table
        sqlite_repository.clear_all_workers()

        # 2. Write JSON workers to SQLite
        sqlite_repository.save_workers(json_workers)

        # 3. Read back workers from SQLite and compare count
        sqlite_workers = sqlite_repository.get_all_workers()
        sqlite_count = len(sqlite_workers)

        success = (expected_json_count == sqlite_count)

        return {
            "json_count": expected_json_count,
            "sqlite_count": sqlite_count,
            "success": success,
        }
