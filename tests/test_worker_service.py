from pathlib import Path
import pytest

from repositories.sqlite_worker_repository import SqliteWorkerRepository
from services.worker_service import WorkerService


@pytest.fixture
def temp_worker_service(tmp_path: Path):
    """Fixture providing WorkerService backed by an isolated temporary database."""
    test_db = tmp_path / "test_service.db"
    repo = SqliteWorkerRepository(db_path=test_db)
    service = WorkerService(repository=repo)
    return service


def test_service_get_workers_empty(temp_worker_service: WorkerService):
    """1. Test get_workers returns empty list initially."""
    assert temp_worker_service.get_workers() == []


def test_service_create_worker(temp_worker_service: WorkerService):
    """2. Test create_worker creates a new worker record."""
    new_worker = {
        "姓名": "张三",
        "身份证号": "110101199001011234",
        "手机号": "13800138000",
        "工种": "电工",
    }

    created = temp_worker_service.create_worker(new_worker)
    assert created["姓名"] == "张三"

    workers = temp_worker_service.get_workers()
    assert len(workers) == 1
    assert workers[0]["姓名"] == "张三"


def test_service_create_worker_duplicate_id_card(temp_worker_service: WorkerService):
    """Test create_worker with duplicate id_card raises ValueError."""
    worker = {"姓名": "张三", "身份证号": "110101199001011234"}
    temp_worker_service.create_worker(worker)

    duplicate_worker = {"姓名": "张三 (重名)", "身份证号": "110101199001011234"}
    with pytest.raises(ValueError) as exc_info:
        temp_worker_service.create_worker(duplicate_worker)
    assert "already exists" in str(exc_info.value)


def test_service_update_worker(temp_worker_service: WorkerService):
    """3. Test update_worker updates existing worker record."""
    worker = {"姓名": "李四", "身份证号": "110101199505054321", "手机号": "13900139000"}
    temp_worker_service.create_worker(worker)

    updated = temp_worker_service.update_worker(
        "110101199505054321", {"手机号": "13999999999", "工种": "木工"}
    )

    assert updated["手机号"] == "13999999999"
    assert updated["工种"] == "木工"

    fetched = temp_worker_service.get_worker_by_id_card("110101199505054321")
    assert fetched["手机号"] == "13999999999"
    assert fetched["工种"] == "木工"


def test_service_get_worker_by_id_card(temp_worker_service: WorkerService):
    """4. Test get_worker_by_id_card queries worker by id_card."""
    worker = {"姓名": "王五", "身份证号": "220202198808088888"}
    temp_worker_service.create_worker(worker)

    found = temp_worker_service.get_worker_by_id_card("220202198808088888")
    assert found is not None
    assert found["姓名"] == "王五"

    not_found = temp_worker_service.get_worker_by_id_card("999999999999999999")
    assert not_found is None
