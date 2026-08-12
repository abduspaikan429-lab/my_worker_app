from pathlib import Path
import pytest

from repositories.sqlite_worker_repository import SqliteWorkerRepository
from services.worker_sync_service import WorkerSyncService


@pytest.fixture
def temp_sqlite_repo(tmp_path: Path):
    db_file = tmp_path / "test_sync.db"
    return SqliteWorkerRepository(db_path=db_file)


def test_sync_empty_sqlite(temp_sqlite_repo: SqliteWorkerRepository):
    """1. Test syncing empty JSON workers into empty SQLite database."""
    res = WorkerSyncService.sync_workers([], temp_sqlite_repo)
    assert res["json_count"] == 0
    assert res["sqlite_count"] == 0
    assert res["success"] is True
    assert temp_sqlite_repo.get_all_workers() == []


def test_sync_five_workers(temp_sqlite_repo: SqliteWorkerRepository):
    """2. Test syncing 5 workers successfully, 3. verifying matching count, and 4. checking id_card query."""
    sample_workers = [
        {"姓名": "Worker 1", "身份证号": "110101199001010001", "手机号": "13800000001"},
        {"姓名": "Worker 2", "身份证号": "110101199001010002", "手机号": "13800000002"},
        {"姓名": "Worker 3", "身份证号": "110101199001010003", "手机号": "13800000003"},
        {"姓名": "Worker 4", "身份证号": "110101199001010004", "手机号": "13800000004"},
        {"姓名": "Worker 5", "身份证号": "110101199001010005", "手机号": "13800000005"},
    ]

    res = WorkerSyncService.sync_workers(sample_workers, temp_sqlite_repo)

    # 3. Check counts match
    assert res["json_count"] == 5
    assert res["sqlite_count"] == 5
    assert res["success"] is True

    # 4. Check id_card query works
    worker_3 = temp_sqlite_repo.get_worker_by_id_card("110101199001010003")
    assert worker_3 is not None
    assert worker_3["姓名"] == "Worker 3"
    assert worker_3["手机号"] == "13800000003"


def test_sync_clears_previous_data(temp_sqlite_repo: SqliteWorkerRepository):
    """Test that sync clears previous SQLite data before saving new batch."""
    old_workers = [
        {"姓名": "Old 1", "身份证号": "990101199001010001"},
        {"姓名": "Old 2", "身份证号": "990101199001010002"},
    ]
    WorkerSyncService.sync_workers(old_workers, temp_sqlite_repo)
    assert len(temp_sqlite_repo.get_all_workers()) == 2

    new_workers = [
        {"姓名": "New 1", "身份证号": "110101199001010001"},
    ]
    res = WorkerSyncService.sync_workers(new_workers, temp_sqlite_repo)
    assert res["json_count"] == 1
    assert res["sqlite_count"] == 1
    assert temp_sqlite_repo.get_worker_by_id_card("990101199001010001") is None
    assert temp_sqlite_repo.get_worker_by_id_card("110101199001010001") is not None
