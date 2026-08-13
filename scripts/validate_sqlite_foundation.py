import json
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from db.database import DEFAULT_DB_PATH, get_connection, init_db


def validate() -> None:
    # 1. Initialize SQLite
    init_db()

    conn = get_connection()
    cursor = conn.cursor()

    # 2. Check tables exist
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = {row[0] for row in cursor.fetchall()}

    companies_ok = "companies" in tables
    teams_ok = "teams" in tables
    workers_ok = "workers" in tables

    # 5. Output SQLite table structure
    print("--- SQLite Table Structures ---")
    for table_name in ["companies", "teams", "workers"]:
        if table_name in tables:
            print(f"\nTable: {table_name}")
            cursor.execute(f"PRAGMA table_info({table_name});")
            columns = cursor.fetchall()
            for col in columns:
                # col schema: (cid, name, type, notnull, dflt_value, pk)
                col_name = col[1]
                col_type = col[2]
                notnull_str = "NOT NULL" if col[3] else "NULL"
                default_str = f"DEFAULT {col[4]}" if col[4] is not None else ""
                pk_str = "PRIMARY KEY" if col[5] else ""
                details = " ".join(filter(None, [col_name, col_type, notnull_str, default_str, pk_str]))
                print(f"  - {details}")
        else:
            print(f"\nTable: {table_name} (MISSING)")

    # 3 & 4. Check data/master_state.json & count valid workers
    master_json_path = BASE_DIR / "data" / "master_state.json"
    json_status = "NOT FOUND"
    worker_count = 0

    if master_json_path.exists():
        json_status = "FOUND"
        try:
            with open(master_json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                rows = data.get("rows", [])
                worker_count = len([r for r in rows if isinstance(r, dict) and r.get("身份证号")])
        except Exception as e:
            print(f"Error reading master_state.json: {e}")

    conn.close()

    overall_pass = companies_ok and teams_ok and workers_ok

    # 7. Final Formatted Output
    print("\nSQLite foundation validation")
    print("============================")
    print("Database: data/worker.db")
    print(f"companies table: {'PASS' if companies_ok else 'FAIL'}")
    print(f"teams table: {'PASS' if teams_ok else 'FAIL'}")
    print(f"workers table: {'PASS' if workers_ok else 'FAIL'}")
    print(f"master_state.json: {json_status}")
    print(f"master worker count: {worker_count}")
    print("JSON migration: NOT RUN")
    print(f"Overall: {'PASS' if overall_pass else 'FAIL'}")


if __name__ == "__main__":
    validate()
