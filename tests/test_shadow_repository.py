from pathlib import Path
from unittest.mock import patch
import pytest

from repositories.shadow_worker_repository import ShadowWorkerRepository
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from repositories.worker_repository import JsonWorkerRepository


def test_shadow_repo_successful_sync_flow(tmp_path: Path):
    """1. Test JSON save success, 2. SQLite sync success, 3. Compare success."""
    json_file = tmp_path / "data" / "master_state.json"
    sqlite_db = tmp_path / "data" / "worker.db"
    log_file = tmp_path / "logs" / "sync.log"

    json_repo = JsonWorkerRepository(file_path=json_file)
    sqlite_repo = SqliteWorkerRepository(db_path=sqlite_db)
    shadow_repo = ShadowWorkerRepository(
        json_repo=json_repo, sqlite_repo=sqlite_repo, log_path=log_file
    )

    sample_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234", "手机号": "13800138000"},
        {"姓名": "李四", "身份证号": "110101199505054321", "手机号": "13900139000"},
    ]

    # Perform save
    shadow_repo.save_workers(sample_workers)

    # 1. JSON save success
    json_workers = shadow_repo.get_all_workers()
    assert len(json_workers) == 2
    assert json_workers[0]["姓名"] == "张三"

    # 2. SQLite sync success
    sqlite_workers = sqlite_repo.get_all_workers()
    assert len(sqlite_workers) == 2

    # 3. Compare success & log written
    assert log_file.exists()
    log_text = log_file.read_text(encoding="utf-8")
    assert "[INFO] Shadow sync SUCCESS" in log_text


def test_shadow_repo_sqlite_exception_fallback(tmp_path: Path):
    """4. Test when SQLite raises an exception, JSON save still succeeds and error log is recorded."""
    json_file = tmp_path / "data" / "master_state.json"
    sqlite_db = tmp_path / "data" / "worker.db"
    log_file = tmp_path / "logs" / "sync.log"

    json_repo = JsonWorkerRepository(file_path=json_file)
    broken_sqlite_repo = SqliteWorkerRepository(db_path=sqlite_db)

    # Mock clear_all_workers on broken_sqlite_repo to raise RuntimeError
    with patch.object(
        broken_sqlite_repo,
        "clear_all_workers",
        side_effect=RuntimeError("SQLite Database Disk Error"),
    ):
        shadow_repo = ShadowWorkerRepository(
            json_repo=json_repo,
            sqlite_repo=broken_sqlite_repo,
            log_path=log_file,
        )

        sample_workers = [
            {"姓名": "王五", "身份证号": "220202198808088888", "手机号": "13700137000"}
        ]

        # Call save_workers - must NOT raise exception
        shadow_repo.save_workers(sample_workers)

        # Primary JSON repo still saved the data
        saved_workers = shadow_repo.get_all_workers()
        assert len(saved_workers) == 1
        assert saved_workers[0]["姓名"] == "王五"

        # Log file recorded error message
        assert log_file.exists()
        log_text = log_file.read_text(encoding="utf-8")
        assert "[ERROR] Shadow SQLite sync exception" in log_text
        assert "SQLite Database Disk Error" in log_text


def test_shadow_repo_id_card_search(tmp_path: Path):
    """Test get_worker_by_id_card queries primary JSON repository."""
    json_file = tmp_path / "data" / "master_state.json"
    sqlite_db = tmp_path / "data" / "worker.db"

    json_repo = JsonWorkerRepository(file_path=json_file)
    sqlite_repo = SqliteWorkerRepository(db_path=sqlite_db)
    shadow_repo = ShadowWorkerRepository(
        json_repo=json_repo, sqlite_repo=sqlite_repo
    )

    sample_workers = [
        {"姓名": "赵六", "身份证号": "330303200001015555"}
    ]
    shadow_repo.save_workers(sample_workers)

    worker = shadow_repo.get_worker_by_id_card("330303200001015555")
    assert worker is not None
    assert worker["姓名"] == "赵六"
