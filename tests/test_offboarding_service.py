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
    data = {"info": {"姓名": "李四", "班组": "木工组"}}
    temp_offboarding_service.create_offboarding(data)

    records = temp_offboarding_service.get_records()
    assert len(records) == 1
    assert "李四_木工组" in records

    pending = temp_offboarding_service.get_pending_workers()
    assert len(pending) == 1
    assert "李四_木工组" in pending


def test_update_status_and_archiving(
    temp_offboarding_service: OffboardingService,
):
    """3. Test updating status and auto-archiving when 100% complete."""
    data = {"info": {"姓名": "王五", "班组": "水暖组", "身份证号": "220202198808088888"}}
    temp_offboarding_service.create_onboarding = temp_offboarding_service.create_offboarding
    temp_offboarding_service.create_offboarding(data)

    record_id = "王五_水暖组"
    step_updates = {
        "1. 工人小灵光发起": True,
        "2. 班组长确认": True,
        "3. 劳资员确认": True,
        "4. 劳资员提交财务发放": True,
        "5. 财务发放完成": True,
    }

    temp_offboarding_service.update_status(record_id, step_updates)

    # Completed worker should no longer be pending
    pending = temp_offboarding_service.get_pending_workers()
    assert record_id not in pending

    # History file should contain archived record
    history = temp_offboarding_service.load_history()
    assert len(history) == 1
    assert history[0]["姓名"] == "王五"
