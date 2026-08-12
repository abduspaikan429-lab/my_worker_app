import sqlite3
from pathlib import Path
import pytest

from db.database import get_connection, init_db


@pytest.fixture
def temp_db(tmp_path: Path):
    """Fixture providing a temporary SQLite database connection."""
    db_file = tmp_path / "test_worker.db"
    init_db(db_path=db_file)
    conn = get_connection(db_path=db_file)
    yield conn
    conn.close()


def test_create_temp_db_and_tables(temp_db: sqlite3.Connection):
    """1. Can create temporary SQLite DB & 2. Three tables exist."""
    cursor = temp_db.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = {row[0] for row in cursor.fetchall()}
    assert "companies" in tables
    assert "teams" in tables
    assert "workers" in tables


def test_foreign_keys_enabled(temp_db: sqlite3.Connection):
    """3. foreign_keys is enabled."""
    cursor = temp_db.cursor()
    cursor.execute("PRAGMA foreign_keys;")
    res = cursor.fetchone()
    assert res[0] == 1


def test_teams_company_id_foreign_key(temp_db: sqlite3.Connection):
    """4. teams.company_id foreign key works."""
    cursor = temp_db.cursor()
    with pytest.raises(sqlite3.IntegrityError):
        cursor.execute(
            """
            INSERT INTO teams (company_id, name, created_at, updated_at)
            VALUES (999, 'Non-existent Company Team', '2026-08-12T00:00:00', '2026-08-12T00:00:00')
            """
        )


def test_companies_name_unique_constraint(temp_db: sqlite3.Connection):
    """5. companies.name unique constraint works."""
    cursor = temp_db.cursor()
    cursor.execute(
        """
        INSERT INTO companies (name, created_at, updated_at)
        VALUES ('Unique Company', '2026-08-12T00:00:00', '2026-08-12T00:00:00')
        """
    )
    temp_db.commit()

    with pytest.raises(sqlite3.IntegrityError):
        cursor.execute(
            """
            INSERT INTO companies (name, created_at, updated_at)
            VALUES ('Unique Company', '2026-08-12T00:00:00', '2026-08-12T00:00:00')
            """
        )


def test_workers_id_card_unique_constraint(temp_db: sqlite3.Connection):
    """6. workers.id_card unique constraint works."""
    cursor = temp_db.cursor()
    cursor.execute(
        """
        INSERT INTO workers (name, id_card, created_at, updated_at)
        VALUES ('Worker A', '110101199001011234', '2026-08-12T00:00:00', '2026-08-12T00:00:00')
        """
    )
    temp_db.commit()

    with pytest.raises(sqlite3.IntegrityError):
        cursor.execute(
            """
            INSERT INTO workers (name, id_card, created_at, updated_at)
            VALUES ('Worker B', '110101199001011234', '2026-08-12T00:00:00', '2026-08-12T00:00:00')
            """
        )
