import pandas as pd
import pytest

from services.payroll_service import PayrollService


def test_parse_daily_rate():
    """Test parse_daily_rate extracts numbers from strings correctly."""
    assert PayrollService.parse_daily_rate("基本工资：450元/天") == 450.0
    assert PayrollService.parse_daily_rate("400") == 400.0
    assert PayrollService.parse_daily_rate("500元") == 500.0
    assert PayrollService.parse_daily_rate("") == 0.0
    assert PayrollService.parse_daily_rate(None) == 0.0


def test_calculate_payroll():
    """Test calculate_payroll calculates rates and total wages."""
    df = pd.DataFrame([
        {"姓名": "张三", "身份证号": "110101199001011234", "考勤天数": 20, "结算单价/标准": "基本工资：450元/天"},
        {"姓名": "李四", "身份证号": "110101199505054321", "考勤天数": 15, "结算单价/标准": "400"},
    ])

    res_df = PayrollService.calculate_payroll(df)

    assert "日薪单价" in res_df.columns
    assert "应发工资" in res_df.columns
    assert res_df.iloc[0]["日薪单价"] == 450.0
    assert res_df.iloc[0]["应发工资"] == 9000.0
    assert res_df.iloc[1]["日薪单价"] == 400.0
    assert res_df.iloc[1]["应发工资"] == 6000.0


def test_summarize_payroll():
    """Test summarize_payroll calculates correct totals."""
    df = pd.DataFrame([
        {"姓名": "张三", "考勤天数": 20, "应发工资": 9000.0},
        {"姓名": "李四", "考勤天数": 15, "应发工资": 6000.0},
    ])

    summary = PayrollService.summarize_payroll(df)
    assert summary["total_workers"] == 2
    assert summary["total_days"] == 35.0
    assert summary["total_wages"] == 15000.0


def test_payroll_edge_cases():
    """Test edge cases with missing rates, zero days, or invalid rate strings."""
    df = pd.DataFrame([
        {"姓名": "王五", "身份证号": "220202198808088888", "考勤天数": 0, "结算单价/标准": "无"},
    ])

    res_df = PayrollService.calculate_payroll(df)
    assert res_df.iloc[0]["日薪单价"] == 0.0
    assert res_df.iloc[0]["应发工资"] == 0.0

    summary = PayrollService.summarize_payroll(res_df)
    assert summary["total_workers"] == 1
    assert summary["total_days"] == 0.0
    assert summary["total_wages"] == 0.0
