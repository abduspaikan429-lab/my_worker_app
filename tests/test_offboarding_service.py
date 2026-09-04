from pathlib import Path
import pytest

from services.offboarding_service import OffboardingService


@pytest.fixture
def temp_offboarding_service(tmp_path: Path):
    """Fixture providing OffboardingService backed by temporary test JSON files."""
    test_data = tmp_path / "test_offboarding_data.json"
    test_history = tmp_path / "test_offboarding_history.json"
    service = OffboardingService(file_path=test_data, history_path=test_history)
    return service


def test_create_offboarding(temp_offboarding_service: OffboardingService):
    """1. Test creating offboarding record."""
    data = {
        "info": {
            "姓名": "张三",
            "班组": "架子工组",
            "身份证号": "110101199001011234",
            "手机号": "13800138000",
        }
    }
    record = temp_offboarding_service.create_offboarding(data)

    assert record["info"]["姓名"] == "张三"
    assert record["info"]["班组"] == "架子工组"
    assert "steps" in record


def test_get_records_and_pending_workers(
    temp_offboarding_service: OffboardingService,
):
    """2. Test querying records and pending workers."""
    data = {"info": {"姓名": "李四", "班组": "木工组", "身份证号": "123456"}}
    temp_offboarding_service.create_offboarding(data)

    records = temp_offboarding_service.get_records()
    assert len(records) == 1
    assert "123456" in records

    pending = temp_offboarding_service.get_pending_workers()
    assert len(pending) == 1
    assert "123456" in pending


def test_update_status_and_archiving(
    temp_offboarding_service: OffboardingService,
):
    """3. Test updating status and auto-archiving when 100% complete."""
    data = {"info": {"姓名": "王五", "班组": "水暖组", "身份证号": "220202198808088888"}}
    rec = temp_offboarding_service.create_offboarding(data)
    record_id = "220202198808088888"

    step_updates = {k: True for k in rec.get("steps", {})}
    records = temp_offboarding_service.get_records()
    records[record_id]["paper"] = {k: True for k in records[record_id].get("paper", {})}
    records[record_id]["system"] = {k: True for k in records[record_id].get("system", {})}
    temp_offboarding_service.save_records(records)

    temp_offboarding_service.update_status(record_id, step_updates)

    # History file should contain archived record
    history = temp_offboarding_service.load_history()
    assert len(history) == 1
    assert history[0]["姓名"] == "王五"


def test_immediate_offboarding_deduction_from_onsite(
    temp_offboarding_service: OffboardingService,
):
    """4. Test that adding a worker to offboarding immediately excludes them from on-site headcount."""
    import pandas as pd

    master_df = pd.DataFrame([
        {"姓名": "张三", "班组": "架子工组", "身份证号": "110101199001011234"},
        {"姓名": "李四", "班组": "木工组", "身份证号": "110101199505054321"},
        {"姓名": "王五", "班组": "水暖组", "身份证号": "220202198808088888"},
    ])

    # Initially, all 3 workers are on-site
    onsite = temp_offboarding_service.filter_onsite_df(master_df)
    assert len(onsite) == 3

    # Add 李四 to offboarding (steps not completed yet)
    temp_offboarding_service.create_offboarding({
        "info": {"姓名": "李四", "班组": "木工组", "身份证号": "110101199505054321"}
    })

    # Verification 1: Identifiers include 李四
    left_ids, left_names = temp_offboarding_service.get_offboarded_identifiers(include_active=True)
    assert "110101199505054321" in left_ids

    # Verification 2: filter_onsite_df immediately deducts 李四 from on-site
    onsite_after_add = temp_offboarding_service.filter_onsite_df(master_df)
    assert len(onsite_after_add) == 2
    assert "李四" not in onsite_after_add["姓名"].values
    assert set(onsite_after_add["姓名"].values) == {"张三", "王五"}

    # Complete all steps for 李四 to auto-archive
    record_id = "110101199505054321"
    records = temp_offboarding_service.get_records()
    records[record_id]["paper"] = {k: True for k in records[record_id].get("paper", {})}
    records[record_id]["system"] = {k: True for k in records[record_id].get("system", {})}
    step_updates = {k: True for k in records[record_id].get("steps", {})}
    temp_offboarding_service.save_records(records)
    temp_offboarding_service.update_status(record_id, step_updates)

    history = temp_offboarding_service.load_history()
    assert any(h["姓名"] == "李四" for h in history)


def test_revoke_offboarding_restores_onsite(
    temp_offboarding_service: OffboardingService,
):
    """5. Test that revoking offboarding restores worker to on-site list."""
    import pandas as pd

    master_df = pd.DataFrame([
        {"姓名": "张三", "班组": "架子工组", "身份证号": "110101199001011234"},
        {"姓名": "李四", "班组": "木工组", "身份证号": "110101199505054321"},
    ])

    # Add 张三 to offboarding
    temp_offboarding_service.create_offboarding({
        "info": {"姓名": "张三", "班组": "架子工组", "身份证号": "110101199001011234"}
    })
    onsite = temp_offboarding_service.filter_onsite_df(master_df)
    assert len(onsite) == 1
    assert onsite.iloc[0]["姓名"] == "李四"

    # Revoke 张三 from offboarding (delete from records)
    records = temp_offboarding_service.get_records()
    for k in list(records.keys()):
        if records[k].get("info", {}).get("姓名") == "张三":
            del records[k]
    temp_offboarding_service.save_records(records)

    # Verification: 张三 is restored to on-site
    onsite_restored = temp_offboarding_service.filter_onsite_df(master_df)
    assert len(onsite_restored) == 2
    assert set(onsite_restored["姓名"].values) == {"张三", "李四"}

