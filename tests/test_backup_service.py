from pathlib import Path
import pytest

from services.backup_service import BackupService


@pytest.fixture
def temp_backup_setup(tmp_path: Path):
    """Fixture providing isolated temporary backup directory and test data files."""
    backup_dir = tmp_path / "data" / "backups"
    db_file = tmp_path / "data" / "test_worker.db"
    json_file = tmp_path / "data" / "test_master.json"

    db_file.parent.mkdir(parents=True, exist_ok=True)
    db_file.write_text("DUMMY DB CONTENT", encoding="utf-8")
    json_file.write_text('{"rows": []}', encoding="utf-8")

    service = BackupService(backup_dir=backup_dir)
    return {
        "service": service,
        "backup_dir": backup_dir,
        "db_file": db_file,
        "json_file": json_file,
    }


def test_backup_database_success(temp_backup_setup):
    """1. Test database backup success."""
    service: BackupService = temp_backup_setup["service"]
    db_file: Path = temp_backup_setup["db_file"]

    backup_path = service.backup_database(db_file)
    assert backup_path.exists()
    assert "worker_" in backup_path.name
    assert backup_path.read_text(encoding="utf-8") == "DUMMY DB CONTENT"


def test_backup_json_success(temp_backup_setup):
    """2. Test JSON backup success."""
    service: BackupService = temp_backup_setup["service"]
    json_file: Path = temp_backup_setup["json_file"]

    backup_path = service.backup_json(json_file)
    assert backup_path.exists()
    assert "master_" in backup_path.name
    assert backup_path.read_text(encoding="utf-8") == '{"rows": []}'


def test_list_backups(temp_backup_setup):
    """3. Test list_backups returns dictionary of database and json backup paths."""
    service: BackupService = temp_backup_setup["service"]
    db_file: Path = temp_backup_setup["db_file"]
    json_file: Path = temp_backup_setup["json_file"]

    service.backup_database(db_file)
    service.backup_json(json_file)

    backups = service.list_backups()
    assert len(backups["db_backups"]) == 1
    assert len(backups["json_backups"]) == 1


def test_cleanup_old_backups(temp_backup_setup):
    """4. Test cleanup_old_backups removes oldest files when limit exceeded."""
    service: BackupService = temp_backup_setup["service"]
    backup_dir: Path = temp_backup_setup["backup_dir"]
    backup_dir.mkdir(parents=True, exist_ok=True)

    # Create 5 dummy DB backups and 5 dummy JSON backups
    for i in range(5):
        (backup_dir / f"worker_2026010{i}_000000.db").write_text(f"DB {i}")
        (backup_dir / f"master_2026010{i}_000000.json").write_text(f"JSON {i}")

    # Cleanup with max_backups = 3
    del_res = service.cleanup_old_backups(max_backups=3)
    assert del_res["db_deleted"] == 2
    assert del_res["json_deleted"] == 2

    remaining = service.list_backups()
    assert len(remaining["db_backups"]) == 3
    assert len(remaining["json_backups"]) == 3


def test_restore_database_success(temp_backup_setup):
    """5. Test restore_database creates pre-restore backup and restores database file."""
    service: BackupService = temp_backup_setup["service"]
    db_file: Path = temp_backup_setup["db_file"]

    # Initial backup
    backup_path = service.backup_database(db_file)

    # Modify database file
    db_file.write_text("MODIFIED DB CONTENT", encoding="utf-8")

    # Restore from backup
    pre_restore = service.restore_database(backup_path, target_db_path=db_file)

    assert pre_restore is not None
    assert pre_restore.exists()
    # Content of target db should be restored to original
    assert db_file.read_text(encoding="utf-8") == "DUMMY DB CONTENT"
