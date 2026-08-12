from pathlib import Path
import pytest

from repositories.worker_repository import JsonWorkerRepository, WorkerRepository


def test_base_worker_repository_raises_not_implemented():
    """Verify base WorkerRepository methods raise NotImplementedError."""
    repo = WorkerRepository()
    with pytest.raises(NotImplementedError):
        repo.get_all_workers()

    with pytest.raises(NotImplementedError):
        repo.save_workers([])

    with pytest.raises(NotImplementedError):
        repo.get_worker_by_id_card("110101199001011234")


def test_json_worker_repository_non_existent_file(tmp_path: Path):
    """Verify reading non-existent file returns empty list and None for lookup."""
    test_file = tmp_path / "non_existent.json"
    repo = JsonWorkerRepository(file_path=test_file)

    assert repo.get_all_workers() == []
    assert repo.get_worker_by_id_card("110101199001011234") is None


def test_json_worker_repository_save_and_read(tmp_path: Path):
    """Verify save_workers writes formatted JSON and get_all_workers reads it back correctly."""
    test_file = tmp_path / "test_master_state.json"
    repo = JsonWorkerRepository(file_path=test_file)

    sample_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234", "手机号": "13800138000"},
        {"姓名": "李四", "身份证号": "110101199505054321", "手机号": "13900139000"},
    ]

    repo.save_workers(sample_workers)
    assert test_file.exists()

    # Check file formatting
    content = test_file.read_text(encoding="utf-8")
    assert '"rows": [' in content
    assert '"张三"' in content

    # Read back workers
    read_workers = repo.get_all_workers()
    assert len(read_workers) == 2
    assert read_workers[0]["姓名"] == "张三"
    assert read_workers[1]["姓名"] == "李四"


def test_json_worker_repository_get_worker_by_id_card(tmp_path: Path):
    """Verify get_worker_by_id_card searches correctly and returns None when missing."""
    test_file = tmp_path / "test_master_state.json"
    repo = JsonWorkerRepository(file_path=test_file)

    sample_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234", "手机号": "13800138000"},
        {"姓名": "王五", "id_card": "220202198808088888", "手机号": "13700137000"},
    ]
    repo.save_workers(sample_workers)

    # Match by Chinese key '身份证号'
    worker1 = repo.get_worker_by_id_card("110101199001011234")
    assert worker1 is not None
    assert worker1["姓名"] == "张三"

    # Match by English key 'id_card'
    worker2 = repo.get_worker_by_id_card("220202198808088888")
    assert worker2 is not None
    assert worker2["姓名"] == "王五"

    # Non-existent search
    worker_none = repo.get_worker_by_id_card("999999999999999999")
    assert worker_none is None


def test_real_master_state_unmodified():
    """Verify real data/master_state.json is untouched by tests."""
    real_file = Path("data/master_state.json")
    if real_file.exists():
        stat_before = real_file.stat().st_mtime
        repo = JsonWorkerRepository(file_path=real_file)
        workers = repo.get_all_workers()
        assert isinstance(workers, list)
        stat_after = real_file.stat().st_mtime
        assert stat_before == stat_after
