from __future__ import annotations

from typing import Optional

from config import settings
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from repositories.worker_repository import JsonWorkerRepository, WorkerRepository


def get_worker_repository(
    storage_type: Optional[str] = None,
) -> WorkerRepository:
    """Factory function to instantiate the configured WorkerRepository implementation.

    Args:
        storage_type: Optional storage mode override ('json' or 'sqlite').
                     If None, reads settings.WORKER_STORAGE.

    Returns:
        WorkerRepository: An instance of JsonWorkerRepository or SqliteWorkerRepository.

    Raises:
        ValueError: If storage mode is unsupported.
    """
    mode = storage_type or getattr(settings, "WORKER_STORAGE", "json")
    if not isinstance(mode, str):
        raise ValueError(f"Invalid WORKER_STORAGE configuration type: {type(mode)}")

    clean_mode = mode.lower().strip()

    if clean_mode == "json":
        json_path = getattr(settings, "JSON_STATE_FILE", "data/master_state.json")
        return JsonWorkerRepository(file_path=json_path)
    elif clean_mode == "sqlite":
        db_path = getattr(settings, "SQLITE_DB_PATH", "data/worker.db")
        return SqliteWorkerRepository(db_path=db_path)
    else:
        raise ValueError(
            f"Unsupported WORKER_STORAGE mode: '{mode}'. Allowed modes are 'json' or 'sqlite'."
        )
