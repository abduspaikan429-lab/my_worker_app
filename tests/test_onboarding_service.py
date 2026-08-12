from pathlib import Path
import pytest

from services.onboarding_service import OnboardingService


@pytest.fixture
def temp_onboarding_service(tmp_path: Path):
    """Fixture providing OnboardingService backed by an isolated temporary JSON file."""
    test_json = tmp_path / "test_onboarding_data.json"
    service = OnboardingService(file_path=test_json)
    return service


def test_create_onboarding(temp_onboarding_service: OnboardingService):
    """1. Test creating a new onboarding record."""
    data = {
        "info": {
            "姓名": "张三",
            "班组": "木工组",
            "身份证号": "110101199001011234",
            "手机号": "13800138000",
        }
    }
    record = temp_onboarding_service.create_onboarding(data)

    assert record["info"]["姓名"] == "张三"
    assert record["info"]["班组"] == "木工组"
    assert "paper" in record
    assert "system" in record
    assert "access" in record


def test_get_records_and_pending_workers(
    temp_onboarding_service: OnboardingService,
):
    """2. Test querying all records and pending workers."""
    data = {"info": {"姓名": "李四", "班组": "电工组"}}
    temp_onboarding_service.create_onboarding(data)

    all_records = temp_onboarding_service.get_records()
    assert len(all_records) == 1
    assert "李四_电工组" in all_records

    pending = temp_onboarding_service.get_pending_workers()
    assert len(pending) == 1
    assert "李四_电工组" in pending


def test_complete_onboarding(temp_onboarding_service: OnboardingService):
    """3. Test marking onboarding complete removes record from pending list."""
    data = {"info": {"姓名": "王五", "班组": "架子工组"}}
    temp_onboarding_service.create_onboarding(data)

    record_id = "王五_架子工组"
    assert record_id in temp_onboarding_service.get_pending_workers()

    # Complete onboarding
    temp_onboarding_service.complete_onboarding(record_id)

    # All records still contains it
    all_records = temp_onboarding_service.get_records()
    assert record_id in all_records

    # Pending workers filter no longer contains completed worker
    pending = temp_onboarding_service.get_pending_workers()
    assert record_id not in pending


def test_data_persistence_and_formatting(
    temp_onboarding_service: OnboardingService,
):
    """4. Test data is correctly formatted and persisted to JSON file."""
    data = {"info": {"姓名": "赵六", "班组": "水暖组"}}
    temp_onboarding_service.create_onboarding(data)

    test_file = temp_onboarding_service.file_path
    assert test_file.exists()

    content = test_file.read_text(encoding="utf-8")
    assert '"赵六_水暖组"' in content
    assert '"姓名": "赵六"' in content
