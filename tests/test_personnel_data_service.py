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

    # 验证去重花名册：Excel 原始底册基准人数为 72 人，并验证四源动态联动生效
    unique_roster = data["unique_roster"]
    excel_base = unique_roster[unique_roster["Source"] == "Excel基准底册"]
    assert len(excel_base) == 72
    assert len(unique_roster) >= 72

    # 验证全系统四源实时联动状态
    assert "live_status" in data
    live_status = data["live_status"]
    assert "currently_onsite" in live_status
    assert "unpasted_count" in live_status
    assert live_status["master_count"] >= 0

    # 验证月度流动数据：官方原始【合计】基准表与动态增量
    official = data.get("official_summary", {})
    assert official["6月"]["in_total"] == 34
    assert official["7月"]["in_total"] == 24
    assert official["8月"]["in_total"] == 13
    assert official["9月"]["in_total"] == 4
    assert official["6月"]["out_total"] == 0
    assert official["7月"]["out_total"] == 11
    assert official["8月"]["out_total"] == 3
    assert official["9月"]["out_total"] == 10
    assert official["6月"]["onsite_total"] == 34
    assert official["7月"]["onsite_total"] == 58
    assert official["8月"]["onsite_total"] == 60
    assert official["9月"]["onsite_total"] == 61
    assert official["6月"]["actual_onsite_total"] == 34
    assert official["7月"]["actual_onsite_total"] == 47
    assert official["8月"]["actual_onsite_total"] == 57
    assert official["9月"]["actual_onsite_total"] == 51

    # 验证大屏实时动态联动汇总（当月包含进场在办与主表新同步人员）
    summary = data["monthly_summary"]
    assert summary["9月"]["in_total"] >= 4
    assert summary["9月"]["onsite_total"] >= 61
    assert summary["9月"]["actual_onsite_total"] >= 51

    # 验证合规管控指标（在办新工人合同处于办理中，整体合规率仍处于高位受控）
    compliance = data["compliance"]
    assert compliance["contract_rate"] >= 85.0
    assert compliance["wage_settle_rate"] == 100.0
    assert compliance["inflow_total"] >= 72
    assert compliance["outflow_total"] >= 24

    # 验证人口特征
    demo = data["demographics"]
    assert demo["total_unique"] >= 72
    assert demo["avg_age"] > 40
    assert "18-29岁 (青年)" in demo["age_dist"]
    assert len(demo["top_provinces"]) > 0


def test_deduplication_and_offboarded_exclusion():
    service = PersonnelDataService()
    data = service.load_all_data(force_reload=True)
    df_inflow = data["df_inflow"]
    df_outflow = data["df_outflow"]
    unique_roster = data["unique_roster"]

    # 1. 验证张克美去重：9 月进场明细与总花名册中均仅能且必须出现 1 次
    m9_inflow = df_inflow[df_inflow["Month"] == "9月"]
    zk_inflow = m9_inflow[m9_inflow["Name"].str.strip() == "张克美"]
    assert len(zk_inflow) == 1, f"9月进场名单中张克美应仅出现1次，实际出现 {len(zk_inflow)} 次"

    zk_roster = unique_roster[unique_roster["Name"].str.strip() == "张克美"]
    assert len(zk_roster) == 1, f"去重花名册中张克美应仅出现1次，实际出现 {len(zk_roster)} 次"

    # 2. 验证张学成离场判定：绝不能作为新进场录入 df_inflow，且在 df_outflow 6月中规范留痕
    zxc_inflow = df_inflow[df_inflow["Name"].str.strip() == "张学成"]
    assert len(zxc_inflow) == 0, f"张学成已在6月离场，严禁出现在进场明细表中，实际出现 {len(zxc_inflow)} 条"

    zxc_outflow = df_outflow[df_outflow["Name"].str.strip() == "张学成"]
    assert len(zxc_outflow) == 1, "张学成应且仅能在离场月报表中留痕"
    assert zxc_outflow.iloc[0]["Month"] == "6月"


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


def test_get_monthly_rosters():
    service = PersonnelDataService()
    data = service.load_all_data(force_reload=True)
    summary = data["monthly_summary"]

    for m in ["6月", "7月", "8月", "9月"]:
        rep_df, act_df = service.get_monthly_rosters(m, data=data)
        assert len(rep_df) == summary[m]["onsite_total"]
        assert len(act_df) == summary[m]["actual_onsite_total"]
        assert "MonthStatus" in rep_df.columns
        assert "LiveStatus" in rep_df.columns
        assert (rep_df["MonthStatus"] == "🛫 当月办结离场").sum() == summary[m]["out_total"]
        assert (rep_df["MonthStatus"] == "✅ 月末在场留存").sum() == summary[m]["actual_onsite_total"]


def test_offboarding_pipeline_onsite_headcount_lifecycle():
    """
    测试离场人员全生命周期联动规则：
    1. 办理离场流程时 (pending) 即刻扣减在场人数，状态变为 '离场结算中'，不计入现场实时在场；
    2. 办完离场手续归档后 (archived)，在场人数保持扣减，状态变为 '已离场归档'，且严禁死灰复燃为进场在办；
    3. 撤销/删除后能够正确恢复。
    """
    from services.offboarding_service import OffboardingService
    from pathlib import Path
    import copy

    service = PersonnelDataService()
    off_svc = OffboardingService()

    master_state_path = Path("data/master_state.json")
    orig_ms_bytes = master_state_path.read_bytes() if master_state_path.exists() else None

    # 1. 基准状态加载
    base_data = service.load_all_data(force_reload=True)
    base_onsite = base_data["live_status"]["currently_onsite"]
    assert base_onsite >= 50

    # 选取一名当前在场的真实工人（例如王留刚）
    target_id = "372930198212205171"
    target_name = "王留刚"
    target_team = "江苏旭之升 (王宜强施工班组)"

    orig_off_records = copy.deepcopy(off_svc.get_records())
    orig_history = copy.deepcopy(off_svc.load_history())

    try:
        # A. 发起离场流程（办理离场手续中）
        off_svc.create_offboarding({
            "info": {
                "姓名": target_name,
                "班组": target_team,
                "身份证号": target_id,
                "离场日期": "2026-10-26"
            }
        })

        data_pending = service.load_all_data(force_reload=True)
        pending_onsite = data_pending["live_status"]["currently_onsite"]
        # 办理离场流程时就已经不算是在场了 -> 在场人数必须立减 1
        assert pending_onsite == base_onsite - 1

        u_pending = data_pending["unique_roster"]
        target_row = u_pending[u_pending["ID"] == target_id]
        assert len(target_row) == 1
        assert target_row.iloc[0]["LiveStatus"] == "离场结算中"

        # 现场实时在场名单中必须已剔除该工人
        onsite_workers = u_pending[u_pending["LiveStatus"].isin(["在场正常", "进场手续在办"])]
        assert target_name not in onsite_workers["Name"].values

        # B. 办完离场手续（全流程办结归档）
        off_svc.mark_completed(target_id, is_completed=True)

        data_archived = service.load_all_data(force_reload=True)
        archived_onsite = data_archived["live_status"]["currently_onsite"]
        # 办完离场手续后，现场实时在场人数保持扣减状态（不能反弹，也不能不变）
        assert archived_onsite == base_onsite - 1

        u_archived = data_archived["unique_roster"]
        target_row_archived = u_archived[u_archived["ID"] == target_id]
        assert len(target_row_archived) == 1
        assert target_row_archived.iloc[0]["LiveStatus"] == "已离场归档"

        # 现场实时在场名单中依然不包含该工人
        onsite_workers_arch = u_archived[u_archived["LiveStatus"].isin(["在场正常", "进场手续在办"])]
        assert target_name not in onsite_workers_arch["Name"].values

        # 验证退场归档工人曾超：系统主表同步已退场，绝不可被进场流水线死灰复燃为“进场手续在办”
        zc_rows = u_archived[u_archived["Name"] == "曾超"]
        assert len(zc_rows) == 1
        assert zc_rows.iloc[0]["LiveStatus"] == "已离场归档"
        assert zc_rows.iloc[0]["Name"] not in onsite_workers_arch["Name"].values

    finally:
        # 清理还原测试状态
        off_svc.save_records(orig_off_records)
        import json
        with open(off_svc.history_path, "w", encoding="utf-8") as f:
            json.dump(orig_history, f, ensure_ascii=False, indent=2)
        if orig_ms_bytes and master_state_path.exists():
            master_state_path.write_bytes(orig_ms_bytes)
            try:
                from modules.master_data import worker_service
                ms_obj = json.loads(orig_ms_bytes.decode("utf-8"))
                worker_service.save_workers(ms_obj.get("rows", []))
            except Exception:
                pass

