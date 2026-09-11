# tests/test_personnel_data_service.py
import pytest
from services.personnel_data_service import PersonnelDataService, get_province_from_id, calculate_age_from_id, clean_job_title


def test_id_helper_functions():
    # 籍贯解析测试
    assert get_province_from_id("410225198909026650") == "河南"
    assert get_province_from_id("51102319740811611X") == "四川"
    assert get_province_from_id("320101199001011234") == "江苏"
    assert get_province_from_id("999999199001011234") == "其他"
    assert get_province_from_id("") == "其他"

    # 年龄计算测试
    age = calculate_age_from_id("410225198909026650", base_year=2026)
    assert age == 37
    assert calculate_age_from_id("invalid_id") is None

    # 工种清洗测试
    assert clean_job_title("其它") == "其他"
    assert clean_job_title("nan") == "普工"
    assert clean_job_title("电焊工") == "电焊工"


def test_personnel_data_service_load():
    service = PersonnelDataService()
    status = service.get_data_status()
    assert "roster" in status
    assert "inflow" in status
    assert "outflow" in status

    data = service.load_all_data(force_reload=True)
    assert "df_roster" in data
    assert "unique_roster" in data
    assert "df_inflow" in data
    assert "df_outflow" in data
    assert "monthly_summary" in data
    assert "demographics" in data
    assert "compliance" in data

    # 验证去重花名册总人数为 72 人
    unique_roster = data["unique_roster"]
    assert len(unique_roster) == 72

    # 验证月度流动数据与官方【合计】基准表完全对齐
    summary = data["monthly_summary"]
    assert summary["6月"]["in_total"] == 34
    assert summary["7月"]["in_total"] == 24
    assert summary["8月"]["in_total"] == 13
    assert summary["9月"]["in_total"] == 4

    assert summary["6月"]["out_total"] == 0
    assert summary["7月"]["out_total"] == 11
    assert summary["8月"]["out_total"] == 3
    assert summary["9月"]["out_total"] == 10

    # 验证报备在场总人数
    assert summary["6月"]["onsite_total"] == 34
    assert summary["7月"]["onsite_total"] == 58
    assert summary["8月"]["onsite_total"] == 60
    assert summary["9月"]["onsite_total"] == 61

    # 验证实际在场总人数
    assert summary["6月"]["actual_onsite_total"] == 34
    assert summary["7月"]["actual_onsite_total"] == 47
    assert summary["8月"]["actual_onsite_total"] == 57
    assert summary["9月"]["actual_onsite_total"] == 51

    # 验证合规管控指标 100%
    compliance = data["compliance"]
    assert compliance["contract_rate"] == 100.0
    assert compliance["wage_settle_rate"] == 100.0
    assert compliance["commit_rate"] == 100.0
    assert compliance["inflow_total"] == 72
    assert compliance["outflow_total"] == 25

    # 验证人口特征
    demo = data["demographics"]
    assert demo["total_unique"] == 72
    assert demo["avg_age"] > 40
    assert "18-29岁 (青年)" in demo["age_dist"]
    assert len(demo["top_provinces"]) > 0


def test_figure_generations():
    service = PersonnelDataService()
    
    # 验证 Plotly 交互图表生成
    plotly_figs = service.generate_plotly_figures()
    assert "trend" in plotly_figs
    assert "teams" in plotly_figs
    assert "jobs" in plotly_figs
    assert "age" in plotly_figs
    assert "province" in plotly_figs

    # 验证 Matplotlib 报表生成
    fig = service.generate_matplotlib_figure()
    assert fig is not None
    import matplotlib.pyplot as plt
    plt.close(fig)
