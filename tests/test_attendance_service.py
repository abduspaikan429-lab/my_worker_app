import io
import pandas as pd
import pytest

from services.attendance_service import AttendanceService


def test_attendance_state_parsing():
    """Test attendance_state parses normal and abnormal attendance values correctly."""
    assert AttendanceService.attendance_state("√") == "有考勤"
    assert AttendanceService.attendance_state("✓") == "有考勤"
    assert AttendanceService.attendance_state("08:00-18:00") == "有考勤"
    assert AttendanceService.attendance_state("") == "缺勤/无记录"
    assert AttendanceService.attendance_state("--") == "缺勤/无记录"
    assert AttendanceService.attendance_state("请假") == "待确认"


def test_clean_identity():
    """Test clean_identity removes spaces and trailing .0."""
    assert AttendanceService.clean_identity(" 110101199001011234 ") == "110101199001011234"
    assert AttendanceService.clean_identity("110101199001011234.0") == "110101199001011234"
    assert AttendanceService.is_valid_identity("110101199001011234") is True
    assert AttendanceService.is_valid_identity("123") is False


def test_day_number_parsing():
    """Test day_number parses columns like '1', '01日', '15号' correctly."""
    assert AttendanceService.day_number("1") == 1
    assert AttendanceService.day_number("01日") == 1
    assert AttendanceService.day_number("15号") == 15
    assert AttendanceService.day_number("31") == 31
    assert AttendanceService.day_number("32") is None
    assert AttendanceService.day_number("姓名") is None


def test_extract_daily_source():
    """Test extract_daily_source extracts daily attendance values per worker."""
    df = pd.DataFrame([
        {"姓名": "张三", "身份证号": "110101199001011234", "1日": "√", "2日": "√"},
        {"姓名": "李四", "身份证号": "110101199505054321", "1日": "", "2日": "√"},
    ])
    res = AttendanceService.extract_daily_source(df, "三局")

    assert len(res) == 2
    assert "110101199001011234" in res
    assert res["110101199001011234"]["days"][1] == "√"
    assert res["110101199001011234"]["days"][2] == "√"
