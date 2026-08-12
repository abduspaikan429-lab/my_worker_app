from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from repositories.worker_repository import JsonWorkerRepository
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from services.worker_compare_service import WorkerCompareService


def run_comparison() -> None:
    json_path = BASE_DIR / "data" / "master_state.json"
    sqlite_path = BASE_DIR / "data" / "worker.db"

    json_repo = JsonWorkerRepository(file_path=json_path)
    sqlite_repo = SqliteWorkerRepository(db_path=sqlite_path)

    json_workers = json_repo.get_all_workers()
    sqlite_workers = sqlite_repo.get_all_workers()

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    print("\nJSON vs SQLite Worker Comparison Report")
    print("=======================================")
    print(f"JSON Repository File: {json_path.relative_to(BASE_DIR)}")
    print(f"SQLite Database File: {sqlite_path.relative_to(BASE_DIR)}")
    print(f"JSON Worker Count: {report['json_count']}")
    print(f"SQLite Worker Count: {report['sqlite_count']}")
    print(f"Missing in SQLite Count: {len(report['missing_in_sqlite'])}")
    if report['missing_in_sqlite']:
        print(f"  Missing ID Cards: {report['missing_in_sqlite']}")
    print(f"Missing in JSON Count: {len(report['missing_in_json'])}")
    if report['missing_in_json']:
        print(f"  Extra ID Cards: {report['missing_in_json']}")
    print(f"Different Name Count: {len(report['different_name'])}")
    if report['different_name']:
        print(f"  Different Name Details: {report['different_name']}")
    print(f"Overall Comparison Result: {report['result']}")


if __name__ == "__main__":
    run_comparison()
