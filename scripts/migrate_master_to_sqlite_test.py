import json
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from repositories.sqlite_worker_repository import SqliteWorkerRepository


def run_migration_test() -> None:
    json_path = BASE_DIR / "data" / "master_state.json"
    test_db_path = BASE_DIR / "data" / "test_worker_migration.db"

    # Remove test DB if already exists to ensure fresh test
    if test_db_path.exists():
        try:
            test_db_path.unlink()
        except Exception:
            pass

    if not json_path.exists():
        print(f"Error: {json_path} does not exist!")
        print("\nMigration Test Summary")
        print("======================")
        print("Database: data/test_worker_migration.db")
        print("JSON worker count: 0")
        print("SQLite worker count: 0")
        print("Duplicate ID card count: 0")
        print("Migration result: FAIL")
        return

    # Read JSON
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    rows = data.get("rows", [])
    json_worker_count = len(rows)

    # Check ID card duplicates in JSON
    id_cards = [r.get("身份证号") for r in rows if isinstance(r, dict) and r.get("身份证号")]
    unique_id_cards = set(id_cards)
    duplicate_count = len(id_cards) - len(unique_id_cards)

    # Initialize SqliteWorkerRepository with test_worker_migration.db
    repo = SqliteWorkerRepository(db_path=test_db_path)

    # Write workers into SQLite
    repo.save_workers(rows)

    # Read back workers from SQLite
    sqlite_workers = repo.get_all_workers()
    sqlite_worker_count = len(sqlite_workers)

    # Migration PASS criteria:
    # sqlite_worker_count matches number of valid unique workers inserted from JSON
    pass_migration = (sqlite_worker_count == len(unique_id_cards)) and (sqlite_worker_count > 0)

    print("\nMigration Test Summary")
    print("======================")
    print(f"Database: {test_db_path.relative_to(BASE_DIR)}")
    print(f"JSON worker count: {json_worker_count}")
    print(f"SQLite worker count: {sqlite_worker_count}")
    print(f"Duplicate ID card count: {duplicate_count}")
    print(f"Migration result: {'PASS' if pass_migration else 'FAIL'}")


if __name__ == "__main__":
    run_migration_test()
