import pytest
from services.worker_compare_service import WorkerCompareService


def test_compare_exact_match():
    """1. Test exact match between JSON and SQLite records returns PASS."""
    json_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234"},
        {"姓名": "李四", "身份证号": "110101199505054321"},
    ]
    sqlite_workers = [
        {"name": "张三", "id_card": "110101199001011234"},
        {"name": "李四", "id_card": "110101199505054321"},
    ]

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    assert report["json_count"] == 2
    assert report["sqlite_count"] == 2
    assert report["missing_in_sqlite"] == []
    assert report["missing_in_json"] == []
    assert report["different_name"] == []
    assert report["result"] == "PASS"


def test_compare_sqlite_missing_one():
    """2. Test SQLite missing one worker returns FAIL."""
    json_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234"},
        {"姓名": "李四", "身份证号": "110101199505054321"},
    ]
    sqlite_workers = [
        {"name": "张三", "id_card": "110101199001011234"},
    ]

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    assert report["json_count"] == 2
    assert report["sqlite_count"] == 1
    assert report["missing_in_sqlite"] == ["110101199505054321"]
    assert report["missing_in_json"] == []
    assert report["result"] == "FAIL"


def test_compare_different_name():
    """3. Test worker with same id_card but different name returns FAIL."""
    json_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234"},
    ]
    sqlite_workers = [
        {"name": "张三丰", "id_card": "110101199001011234"},
    ]

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    assert report["json_count"] == 1
    assert report["sqlite_count"] == 1
    assert report["missing_in_sqlite"] == []
    assert report["missing_in_json"] == []
    assert len(report["different_name"]) == 1
    assert report["different_name"][0]["id_card"] == "110101199001011234"
    assert report["different_name"][0]["json_name"] == "张三"
    assert report["different_name"][0]["sqlite_name"] == "张三丰"
    assert report["result"] == "FAIL"


def test_compare_json_missing_one():
    """Test JSON missing one worker returns FAIL."""
    json_workers = [
        {"姓名": "张三", "身份证号": "110101199001011234"},
    ]
    sqlite_workers = [
        {"name": "张三", "id_card": "110101199001011234"},
        {"name": "李四", "id_card": "110101199505054321"},
    ]

    report = WorkerCompareService.compare_workers(json_workers, sqlite_workers)

    assert report["json_count"] == 1
    assert report["sqlite_count"] == 2
    assert report["missing_in_json"] == ["110101199505054321"]
    assert report["result"] == "FAIL"
