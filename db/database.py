import sqlite3
from pathlib import Path
from typing import Union

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DB_PATH = BASE_DIR / "data" / "worker.db"
DEFAULT_SCHEMA_PATH = Path(__file__).resolve().parent / "schema.sql"


def get_connection(db_path: Union[str, Path] = DEFAULT_DB_PATH) -> sqlite3.Connection:
    """Connect to SQLite database and apply default configurations."""
    path_obj = Path(db_path)
    if path_obj.parent:
        path_obj.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(path_obj))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA busy_timeout = 5000;")
    return conn


def init_db(
    db_path: Union[str, Path] = DEFAULT_DB_PATH,
    schema_path: Union[str, Path] = DEFAULT_SCHEMA_PATH,
) -> None:
    """Initialize database tables using schema.sql without destroying existing data."""
    schema_file = Path(schema_path)
    if not schema_file.exists():
        raise FileNotFoundError(f"Schema file not found: {schema_file}")

    with open(schema_file, "r", encoding="utf-8") as f:
        schema_sql = f.read()

    conn = get_connection(db_path)
    try:
        conn.executescript(schema_sql)
        conn.commit()
    finally:
        conn.close()
