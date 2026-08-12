from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from db.database import init_db
from repositories.worker_repository import JsonWorkerRepository
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from services.worker_compare_service import WorkerCompareService
from services.worker_sync_service import WorkerSyncService


def run_sync() -> None:
    json_path = BASE_DIR / "data" / "master_state.json"
    sqlite_path = BASE_DIR / "data" / "worker.db"

    # 1. Initialize SQLite
    init_db(sqlite_path)

    # 2. Read JSON
    json_repo = JsonWorkerRepository(file_path=json_path)
    json_workers = json_repo.get_all_workers()

    # 3. Sync to data/worker.db
    sqlite_repo = SqliteWorkerRepository(db_path=sqlite_path)
    sync_res = WorkerSyncService.sync_workers(json_workers, sqlite_repo)

    # 4. Compare using WorkerCompareService
    compare_res = WorkerCompareService.compare_workers(
        json_workers, sqlite_repo.get_all_workers()
    )

    compare_status = "SUCCESS" if (sync_res["success"] and compare_res["result"] == "PASS") else "FAILED"

    # 5. Print report
    print("\nWorker Sync Report")
    print("==================")
    print("\nJSON workers:")
    print(sync_res["json_count"])
    print("\nSQLite workers:")
    print(sync_res["sqlite_count"])
    print("\nCompare:")
    print(compare_status)


if __name__ == "__main__":
    run_sync()
