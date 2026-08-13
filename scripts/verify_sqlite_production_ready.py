from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from db.database import get_connection
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from repositories.worker_repository import JsonWorkerRepository
from services.worker_compare_service import WorkerCompareService


def verify_production_ready() -> bool:
    db_path = BASE_DIR / "data" / "worker.db"
    json_path = BASE_DIR / "data" / "master_state.json"

    print("\nSQLite Production Readiness Verification")
    print("========================================")

    # 1. Check data/worker.db exists
    if not db_path.exists():
        print("1. SQLite DB File Existence: FAIL (data/worker.db missing)")
        print("\nResult: NOT READY")
        return False
    print("1. SQLite DB File Existence: PASS")

    # 2. Check workers table exists
    conn = get_connection(db_path)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='workers';")
        if not cursor.fetchone():
            print("2. Workers Table Existence: FAIL ('workers' table missing)")
            print("\nResult: NOT READY")
            return False
    finally:
        conn.close()
    print("2. Workers Table Existence: PASS")

    # 3, 4, 5. Read workers from JSON & SQLite, compare counts, id_cards, names
    json_repo = JsonWorkerRepository(file_path=json_path)
    sqlite_repo = SqliteWorkerRepository(db_path=db_path)

    json_workers = json_repo.get_all_workers()
    sqlite_workers = sqlite_repo.get_all_workers()

    compare_res = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    json_count = compare_res["json_count"]
    sqlite_count = compare_res["sqlite_count"]

    count_pass = (json_count == sqlite_count) and (json_count > 0)
    id_card_pass = (len(compare_res["missing_in_sqlite"]) == 0) and (len(compare_res["missing_in_json"]) == 0)
    name_pass = (len(compare_res["different_name"]) == 0)

    print(f"3. Worker Count Match: {'PASS' if count_pass else 'FAIL'} (JSON: {json_count}, SQLite: {sqlite_count})")
    print(f"4. ID Card Consistency: {'PASS' if id_card_pass else 'FAIL'}")
    print(f"5. Name Consistency: {'PASS' if name_pass else 'FAIL'}")

    is_ready = count_pass and id_card_pass and name_pass and compare_res["result"] == "PASS"

    print(f"\nResult: {'READY' if is_ready else 'NOT READY'}")
    return is_ready


if __name__ == "__main__":
    verify_production_ready()
