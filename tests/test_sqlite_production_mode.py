from pathlib import Path
from unittest.mock import patch
import pytest

from config import settings
from core.repository_factory import get_worker_repository
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from services.worker_backup_service import backup_workers_to_json
from services.worker_compare_service import WorkerCompareService


def test_sqlite_mode_returns_sqlite_repository(tmp_path: Path):
    """1. Test WORKER_STORAGE='sqlite' returns SqliteWorkerRepository."""
    test_db = tmp_path / "test_prod.db"
    with patch.object(settings, "WORKER_STORAGE", "sqlite"), patch.object(
        settings, "SQLITE_DB_PATH", str(test_db)
    ):
        repo = get_worker_repository()
        assert isinstance(repo, SqliteWorkerRepository)
        assert repo.db_path == test_db


def test_backup_workers_to_json_success(tmp_path: Path):
    """2. Test backup_workers_to_json creates formatted backup JSON."""
    backup_file = tmp_path / "master_state_backup.json"
    sample_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234", "手机号": "13800138000"},
        {"姓名": "李四", "身份证号": "110101199505054321", "手机号": "13900139000"},
    ]

    backup_workers_to_json(sample_workers, backup_path=backup_file)

    assert backup_file.exists()
    content = backup_file.read_text(encoding="utf-8")
    assert '"rows": [' in content
    assert '"张三"' in content
    assert '"李四"' in content


def test_compare_service_validation_success():
    """3. Test compare service validates matching JSON and SQLite workers."""
    json_workers = [
        {"姓名": "王五", "身份证号": "220202198808088888", "手机号": "13700137000"}
    ]
    sqlite_workers = [
        {"name": "王五", "id_card": "220202198808088888", "phone": "13700137000"}
    ]

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)
    assert report["json_count"] == 1
    assert report["sqlite_count"] == 1
    assert report["missing_in_sqlite"] == []
    assert report["missing_in_json"] == []
    assert report["different_name"] == []
    assert report["result"] == "PASS"
