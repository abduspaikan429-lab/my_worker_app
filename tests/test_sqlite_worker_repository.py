from pathlib import Path
import pytest

from repositories.sqlite_worker_repository import SqliteWorkerRepository


@pytest.fixture
def temp_sqlite_repo(tmp_path: Path):
    """Fixture providing an isolated SqliteWorkerRepository with a temporary database."""
    db_file = tmp_path / "temp_sqlite_worker.db"
    repo = SqliteWorkerRepository(db_path=db_file)
    return repo


def test_sqlite_repo_init_and_empty(temp_sqlite_repo: SqliteWorkerRepository):
    """1. Temporary SQLite DB creation and empty query test."""
    workers = temp_sqlite_repo.get_all_workers()
    assert workers == []
    assert temp_sqlite_repo.get_worker_by_id_card("110101199001011234") is None


def test_sqlite_repo_save_and_get_all(temp_sqlite_repo: SqliteWorkerRepository):
    """2. save_workers & 3. get_all_workers test."""
    sample_workers = [
        {
            "姓名": "张三",
            "身份证号": "110101199001011234",
            "手机号": "13800138000",
            "工种": "电焊工",
            "分包/所属企业": "测试建筑公司",
            "班组": "测试班组A",
        },
        {
            "姓名": "李四",
            "身份证号": "110101199505054321",
            "手机号": "13900139000",
            "工种": "木工",
            "分包/所属企业": "测试建筑公司",
            "班组": "测试班组B",
        },
    ]

    temp_sqlite_repo.save_workers(sample_workers)

    workers = temp_sqlite_repo.get_all_workers()
    assert len(workers) == 2
    names = {w["姓名"] for w in workers}
    assert "张三" in names
    assert "李四" in names


def test_sqlite_repo_get_worker_by_id_card(temp_sqlite_repo: SqliteWorkerRepository):
    """4. get_worker_by_id_card test."""
    sample_workers = [
        {
            "姓名": "王五",
            "身份证号": "220202198808088888",
            "手机号": "13700137000",
            "工种": "水暖工",
        }
    ]
    temp_sqlite_repo.save_workers(sample_workers)

    # Search existing worker
    worker = temp_sqlite_repo.get_worker_by_id_card("220202198808088888")
    assert worker is not None
    assert worker["姓名"] == "王五"
    assert worker["手机号"] == "13700137000"

    # Search non-existent worker
    assert temp_sqlite_repo.get_worker_by_id_card("999999999999999999") is None


def test_sqlite_repo_data_consistency(temp_sqlite_repo: SqliteWorkerRepository):
    """5. Data consistency test across save, update and field verification."""
    original_worker = {
        "序号": "001",
        "姓名": "赵六",
        "身份证号": "330303200001015555",
        "手机号": "13600136000",
        "工资卡号": "6222021234567890",
        "工种": "架子工",
        "在场/进退场状态": "已进场",
        "进场日期": "2026-08-01",
        "分包/所属企业": "东方建设工程有限公司",
        "班组": "架子工一班",
    }

    temp_sqlite_repo.save_workers([original_worker])

    saved = temp_sqlite_repo.get_worker_by_id_card("330303200001015555")
    assert saved is not None
    assert saved["姓名"] == original_worker["姓名"]
    assert saved["身份证号"] == original_worker["身份证号"]
    assert saved["手机号"] == original_worker["手机号"]
    assert saved["工资卡号"] == original_worker["工资卡号"]
    assert saved["工种"] == original_worker["工种"]
    assert saved["分包/所属企业"] == original_worker["分包/所属企业"]
    assert saved["班组"] == original_worker["班组"]

    # Test update (upsert)
    updated_worker = dict(original_worker)
    updated_worker["手机号"] = "13699999999"
    updated_worker["工种"] = "高级架子工"

    temp_sqlite_repo.save_workers([updated_worker])

    all_workers = temp_sqlite_repo.get_all_workers()
    assert len(all_workers) == 1  # No duplicate record created

    re_fetched = temp_sqlite_repo.get_worker_by_id_card("330303200001015555")
    assert re_fetched["手机号"] == "13699999999"
    assert re_fetched["工种"] == "高级架子工"
