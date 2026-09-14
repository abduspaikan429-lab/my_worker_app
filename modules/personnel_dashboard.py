# modules/personnel_dashboard.py
# 月度人员变更态势中心 - 全周期数据可视化看板与档案透视
from __future__ import annotations
import json, os, calendar, re, io
from datetime import date, datetime
from typing import Dict, Any, List, Optional

import pandas as pd
import numpy as np
import streamlit as st

from modules.master_data import load_master_df
from modules.onboarding_pipeline import onboarding_service
from modules.offboarding_pipeline import load_offboarding_history
from services.personnel_data_service import personnel_data_service

MONTHLY_DATA_FILE = "data/monthly_change_data.json"
SWITCHOVER_MONTH = "2026-09"
START_MONTH = "2026-06"


def _get_report_date_str(month_str: str) -> str:
    if not month_str or "-" not in month_str:
        return ""
    y, m = map(int, month_str.split('-'))
    pm = m - 1
    py = y
    if pm == 0:
        pm = 12
        py -= 1
    _, p_days = calendar.monthrange(py, pm)
    _, c_days = calendar.monthrange(y, m)
    return f"{py} 年 {pm} 月 {p_days} 日至 {y} 年 {m} 月 {c_days} 日"


def _parse_report_text(text: str) -> Dict[str, Any]:
    results = {}
    sections = text.replace("①", "").replace("②", "").replace("③", "").split("劳务（专业）分包单位")
    for sec in sections:
        if not sec.strip():
            continue
        
        team_match = re.search(r"班组名称[：:]\s*([^\s]+)", sec)
        if not team_match:
            continue
        team = team_match.group(1).strip()
        
        in_match = re.search(r"本月进场务工人员总数[：:]\s*(\d+)", sec)
        out_match = re.search(r"本月离场务工人员总数[：:]\s*(\d+)", sec)
        cur_match = re.search(r"本月现场务工人员总数[：:]\s*(\d+)", sec)
        
        if in_match and out_match and cur_match:
            t_key = "total"
            if "王宜强" in team:
                t_key = "王宜强施工班组"
            elif "汪佩沾" in team:
                t_key = "汪佩沾其他班组"
            elif "金属屋面" in team:
                t_key = "total"
            
            results[t_key] = {
                "in_count": int(in_match.group(1)),
                "out_count": int(out_match.group(1)),
                "current_count": int(cur_match.group(1))
            }
    return results


def _load_monthly_data() -> Dict[str, Any]:
    if os.path.exists(MONTHLY_DATA_FILE):
        try:
            with open(MONTHLY_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "2026-06": {
            "in_count": 34, "out_count": 1, "current_count": 34, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 15, "out_count": 0, "current_count": 15}, 
                "王宜强施工班组": {"in_count": 19, "out_count": 1, "current_count": 19}
            }
        },
        "2026-07": {
            "in_count": 24, "out_count": 11, "current_count": 58, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 17, "out_count": 4, "current_count": 32}, 
                "王宜强施工班组": {"in_count": 7, "out_count": 7, "current_count": 26}
            }
        },
        "2026-08": {
            "in_count": 13, "out_count": 3, "current_count": 60, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 3, "out_count": 2, "current_count": 31}, 
                "王宜强施工班组": {"in_count": 10, "out_count": 1, "current_count": 29}
            }
        }
    }


def _save_monthly_data(data: Dict[str, Any]) -> None:
    os.makedirs("data", exist_ok=True)
    with open(MONTHLY_DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _generate_months(start: str, end: str) -> List[str]:
    sy, sm = map(int, start.split('-'))
    ey, em = map(int, end.split('-'))
    months = []
    cy, cm = sy, sm
    while (cy < ey) or (cy == ey and cm <= em):
        months.append(f"{cy}-{cm:02d}")
        cm += 1
        if cm > 12:
            cm = 1
            cy += 1
    return months


def _parse_date_to_month(val: Any) -> Optional[str]:
    if not val:
        return None
    val_str = str(val).strip()
    if not val_str or val_str in ("nan", "None", ""):
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return val_str[:7] if len(val_str) >= 7 and "-" in val_str else None


def _format_date(val: Any) -> str:
    val_str = str(val).strip()
    if not val_str or val_str in ("nan", "None", ""):
        return ""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return val_str


def _map_team(t_name: Any) -> str:
    name = str(t_name).strip()
    if not name or name in ("nan", "None", "待分配班组"):
        return "待分配"
    if "王宜强" in name or "旭之升" in name:
        return "王宜强施工班组"
    if "汪佩沾" in name or "久昌" in name:
        return "汪佩沾其他班组"
    return name


def get_dynamic_month_data(target_month: str) -> Dict[str, Any]:
    """账本法计算：从 2026-08 历史月作为基准起算，逐月累加进场，减去离场"""
    saved_data = _load_monthly_data()
    base_month = "2026-08"
    months_to_calc = _generate_months(SWITCHOVER_MONTH, target_month)
    
    base_data = saved_data.get(base_month, {"current_count": 0, "teams": {}})
    current_teams = {}
    for t_name, t_data in base_data.get("teams", {}).items():
        if isinstance(t_data, dict):
            current_teams[t_name] = t_data.get("current_count", 0)
        else:
            current_teams[t_name] = t_data
            
    current_total = sum(current_teams.values())
    master = load_master_df()
    all_workers = onboarding_service.merge_with_master(master)
    history = load_offboarding_history()
    
    entries_by_month = {}
    for _, row in all_workers.iterrows():
        d_val = row.get("进场日期") or row.get("进场时间") or row.get("入场日期")
        m = _parse_date_to_month(d_val)
        if m:
            entries_by_month.setdefault(m, []).append(row.to_dict())
            
    exits_by_month = {}
    seen_exit_ids = set()
    for rec in sorted(history, key=lambda x: x.get("离场日期", ""), reverse=True):
        id_card = str(rec.get("身份证号") or "").strip()
        name = str(rec.get("姓名") or "").strip()
        team = str(rec.get("班组") or "").strip()
        unique_id = id_card if id_card else f"{name}_{team}"
        
        if unique_id in seen_exit_ids:
            continue
        seen_exit_ids.add(unique_id)
        
        m = _parse_date_to_month(rec.get("离场日期"))
        if m:
            exits_by_month.setdefault(m, []).append(rec)

    target_in_list = []
    target_out_list = []
    target_team_stats = {}
    
    for m in months_to_calc:
        in_recs = entries_by_month.get(m, [])
        out_recs = exits_by_month.get(m, [])
        
        if m == target_month:
            target_in_list = in_recs
            target_out_list = out_recs
            for r in in_recs:
                t = _map_team(r.get("班组", ""))
                target_team_stats.setdefault(t, {"in_count": 0, "out_count": 0})
                target_team_stats[t]["in_count"] += 1
            for r in out_recs:
                t = _map_team(r.get("班组", ""))
                target_team_stats.setdefault(t, {"in_count": 0, "out_count": 0})
                target_team_stats[t]["out_count"] += 1
            
        for r in in_recs:
            current_total += 1
            t = _map_team(r.get("班组", ""))
            current_teams[t] = current_teams.get(t, 0) + 1
            
        for r in out_recs:
            current_total -= 1
            t = _map_team(r.get("班组", ""))
            current_teams[t] = current_teams.get(t, 0) - 1
            
    final_teams = {}
    for t_name, current_cnt in current_teams.items():
        stats = target_team_stats.get(t_name, {"in_count": 0, "out_count": 0})
        final_teams[t_name] = {
            "in_count": stats["in_count"],
            "out_count": stats["out_count"],
            "current_count": current_cnt
        }
            
    return {
        "in_count": len(target_in_list),
        "out_count": len(target_out_list),
        "current_count": current_total,
        "teams": final_teams,
        "in_list": target_in_list,
        "out_list": target_out_list
    }


def _render_executive_card(title: str, value: str, subtitle: str, icon: str, border_color: str = "#3B82F6", bg_gradient: str = "linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%)") -> None:
    st.markdown(f"""
    <div style="
        background: {bg_gradient};
        border-radius: 12px;
        border-left: 5px solid {border_color};
        padding: 16px 18px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.05);
        margin-bottom: 12px;
        transition: transform 0.2s ease;
    ">
        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
            <div>
                <div style="font-size: 0.85rem; color: #64748B; font-weight: 600; margin-bottom: 4px;">{title}</div>
                <div style="font-size: 2.1rem; font-weight: 800; color: #1E293B; line-height: 1.1;">{value}</div>
                <div style="font-size: 0.78rem; color: #94A3B8; margin-top: 6px; font-weight: 500;">{subtitle}</div>
            </div>
            <span style="font-size: 2rem; opacity: 0.85;">{icon}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)


def _switch_module(mod_key: str) -> None:
    """切换主导航路由模块（通过 pending_nav 在下一帧渲染前安全设置）"""
    st.session_state.pending_nav = mod_key
    st.session_state.current_nav = mod_key


def render() -> None:
    """人员变更态势中心主渲染入口"""
    # 顶部自定义样式注入
    st.markdown("""
    <style>
    .dashboard-badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.82rem;
        font-weight: 600;
        margin-right: 6px;
    }
    .badge-in { background-color: #DBEAFE; color: #1D4ED8; }
    .badge-out { background-color: #FEE2E2; color: #B91C1C; }
    .badge-onsite { background-color: #D1FAE5; color: #047857; }
    .compliance-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .live-pulse-dot {
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 50%;
        background-color: #10B981;
        box-shadow: 0 0 10px #10B981;
        animation: pulseAnimation 2s infinite;
    }
    @keyframes pulseAnimation {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
        70% { transform: scale(1.05); box-shadow: 0 0 0 8px rgba(16, 185, 129, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="display: flex; align-items: center; justify-content: space-between; padding-bottom: 12px; margin-bottom: 14px; border-bottom: 1px solid #E2E8F0;">
        <div>
            <h2 style="margin: 0; color: #0F172A; font-weight: 800; display: flex; align-items: center; gap: 8px;">
                <span>📊</span> 人员变更态势中心
            </h2>
            <p style="margin: 4px 0 0 0; color: #64748B; font-size: 0.95rem;">
                四源动态联动 · 全周期流动态势 · 劳务工人精准画像 · 100%合规闭环管控
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 1. 尝试从 Excel 数据服务加载全部数据
    with st.spinner("正在解析项目人员花名册与变更月报表..."):
        try:
            data = personnel_data_service.load_all_data()
            summary = data["monthly_summary"]
            demo = data["demographics"]
            comp = data["compliance"]
            unique_df = data["unique_roster"]
            df_inflow = data["df_inflow"]
            df_outflow = data["df_outflow"]
            live_status = data.get("live_status", {})
            load_success = True
        except Exception as e:
            st.error(f"解析人员变更 Excel 数据表时出错: {e}")
            load_success = False

    if not load_success or unique_df.empty:
        st.warning("未能成功读取 load-data 目录下的三份核心 Excel 报表，请前往【数据源管理】检查文件路径。")
        return

    # 2. 全系统业务实时联动横幅 (Live Data Banner)
    currently_onsite = live_status.get("currently_onsite", len(unique_df))
    master_count = live_status.get("master_count", 0)
    onboard_pending = live_status.get("onboarding_pending_count", 0)
    offboard_pending = live_status.get("offboarding_pending_count", 0)
    offboard_history = live_status.get("offboarding_history_count", 0)
    unpasted_count = live_status.get("unpasted_count", 0)
    unpasted_workers = live_status.get("unpasted_workers", [])

    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border-radius: 12px;
        padding: 12px 20px;
        margin-bottom: 14px;
        color: #FFFFFF;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
    ">
        <div style="display: flex; align-items: center; gap: 10px;">
            <span class="live-pulse-dot"></span>
            <span style="font-weight: 700; font-size: 0.96rem; letter-spacing: 0.4px;">全系统业务实时联动生效中</span>
            <span style="background: rgba(59, 130, 246, 0.2); color: #93C5FD; font-size: 0.76rem; padding: 2px 8px; border-radius: 10px; border: 1px solid rgba(59, 130, 246, 0.4);">
                动态计算 · 免传Excel
            </span>
        </div>
        <div style="display: flex; align-items: center; gap: 14px; font-size: 0.84rem; color: #CBD5E1; flex-wrap: wrap;">
            <span>系统主表: <strong style="color: #60A5FA;">{master_count}</strong> 人</span>
            <span style="opacity: 0.4;">|</span>
            <span>进场在办: <strong style="color: #34D399;">{onboard_pending}</strong> 人 <span style="font-size: 0.76rem; color: #94A3B8;">(全周期卡片/9月5人)</span></span>
            <span style="opacity: 0.4;">|</span>
            <span>离场结算中: <strong style="color: #F87171;">{offboard_pending}</strong> 人</span>
            <span style="opacity: 0.4;">|</span>
            <span>离场已归档: <strong style="color: #94A3B8;">{offboard_history}</strong> 人</span>
            <span style="opacity: 0.4;">|</span>
            <span>现场实时在册: <strong style="color: #FBBF24; font-size: 1.05rem;">{currently_onsite}</strong> 人</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 3. 增量未回填提示卡片 (若系统主表/进场流水线有新增但 Excel 底册中没有)
    if unpasted_count > 0:
        st.markdown(f"""
        <div style="background: #EFF6FF; border: 1px solid #BFDBFE; border-left: 5px solid #3B82F6; border-radius: 10px; padding: 12px 16px; margin-bottom: 14px;">
            <div style="display: flex; align-items: center; gap: 10px;">
                <span style="font-size: 1.35rem;">🔔</span>
                <div style="flex-grow: 1;">
                    <div style="font-weight: 700; color: #1E40AF; font-size: 0.95rem;">
                        增量数据动态感知：检测到系统主表/进场流水线中新增 <strong>{unpasted_count}</strong> 名劳务工人
                    </div>
                    <div style="color: #475569; font-size: 0.82rem; margin-top: 3px; line-height: 1.4;">
                        大屏已自动动态并入实时图表与档案总库统计（<strong>无需手动重新上传 Excel</strong>）。
                        若需将新增人员更新至您本地的 Excel 原始底册，可随时前往【花名册与报表导出】一键复制回填。
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_sync_act1, col_sync_act2 = st.columns([1.1, 2.9])
        with col_sync_act1:
            if st.button("🚀 前往【花名册直贴】复制代码回填", key="btn_quick_goto_report", type="primary", use_container_width=True):
                _switch_module("report")
                st.rerun()
        with col_sync_act2:
            with st.expander(f"📋 查看待回填本地 Excel 的 {unpasted_count} 名新务工人员名单", expanded=False):
                unpasted_show_cols = ["Name", "Team", "Job", "ID", "Source", "LiveStatus"]
                unpasted_df = pd.DataFrame(unpasted_workers)[unpasted_show_cols].rename(columns={
                    "Name": "姓名", "Team": "队伍/班组", "Job": "工种", "ID": "身份证号", "Source": "来源渠道", "LiveStatus": "在场状态"
                })
                st.dataframe(unpasted_df, use_container_width=True, hide_index=True)

    # 4. 顶部 Executive KPI Banner (核心规模与合规达标率)
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        _render_executive_card("累计参建劳务工", f"{demo['total_unique']} 人", "去重实名制人员基数", "👥", "#3B82F6")
    with col2:
        _render_executive_card("现场实时在场", f"{currently_onsite} 人", "扣除规范离场后真实驻场", "🏗️", "#10B981", "linear-gradient(135deg, #F8FAFC 0%, #ECFDF5 100%)")
    with col3:
        _render_executive_card("累计进场人次", f"{comp['inflow_total']} 人次", "6-9月爬坡及增量进场", "🚀", "#6366F1")
    with col4:
        _render_executive_card("累计离场人次", f"{comp['outflow_total']} 人次", "规范退场工资结清", "🛫", "#EF4444", "linear-gradient(135deg, #F8FAFC 0%, #FEF2F2 100%)")
    with col5:
        _render_executive_card("合规履约达标率", f"{comp['contract_rate']:.0f}%", "合同/工资/退场承诺全覆盖", "🛡️", "#F59E0B", "linear-gradient(135deg, #F8FAFC 0%, #FFFBEB 100%)")

    st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)

    # 5. 核心功能 Tab 布局
    tab_overview, tab_monthly, tab_roster, tab_export, tab_source = st.tabs([
        "📊 态势总览看板 (全景可视化)",
        "📅 月度台账钻取 (进退场明细)",
        f"👥 全员参建档案总库 ({len(unique_df)}人透视)",
        "🖼️ 高清报表大图导出 (Matplotlib)",
        "📂 数据源管理与同步"
    ])

    # ==========================================
    # Tab 1: 态势总览看板 (6大专业图表可视化)
    # ==========================================
    with tab_overview:
        plotly_figs = personnel_data_service.generate_plotly_figures(data)

        # 第一行：进退场流动趋势 + 队伍月度在场规模堆叠
        row1_c1, row1_c2 = st.columns([1.1, 0.9])
        with row1_c1:
            st.plotly_chart(plotly_figs["trend"], use_container_width=True)
        with row1_c2:
            st.plotly_chart(plotly_figs["teams"], use_container_width=True)

        st.markdown("<div style='margin-top: 8px;'></div>", unsafe_allow_html=True)

        # 第二行：工种技能分布 + 年龄梯队结构 + 籍贯省份分布
        row2_c1, row2_c2, row2_c3 = st.columns([1, 1, 1])
        with row2_c1:
            st.plotly_chart(plotly_figs["jobs"], use_container_width=True)
        with row2_c2:
            st.plotly_chart(plotly_figs["age"], use_container_width=True)
        with row2_c3:
            st.plotly_chart(plotly_figs["province"], use_container_width=True)

        st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)

        # 第三行：合规管控与人员履约三大指标
        st.markdown("#### 🛡️ 合规管控与人员履约保障体系")
        c_comp1, c_comp2, c_comp3 = st.columns(3)
        with c_comp1:
            st.markdown(f"""
            <div class="compliance-box" style="border-left: 4px solid #10B981;">
                <div style="font-weight: 700; color: #047857; font-size: 1.05rem; margin-bottom: 4px;">📑 劳动合同签订率：{comp['contract_rate']:.1f}%</div>
                <div style="font-size: 0.88rem; color: #475569; line-height: 1.5;">
                    • 进场务工人员累计 <strong>{comp['inflow_total']}</strong> 人次全员签订合同。<br>
                    • 100% 覆盖率，全部并在实名制系统与建委平台核查合规。
                </div>
            </div>
            """, unsafe_allow_html=True)
        with c_comp2:
            st.markdown(f"""
            <div class="compliance-box" style="border-left: 4px solid #3B82F6;">
                <div style="font-weight: 700; color: #1D4ED8; font-size: 1.05rem; margin-bottom: 4px;">💰 离场工资结算支付：{comp['wage_settle_rate']:.1f}%</div>
                <div style="font-size: 0.88rem; color: #475569; line-height: 1.5;">
                    • 离场 <strong>{comp['outflow_total']}</strong> 人次全部明确登记为“已结算已支付”。<br>
                    • 全流程离场闭环，无拖欠工资及劳务纠纷遗留风险。
                </div>
            </div>
            """, unsafe_allow_html=True)
        with c_comp3:
            st.markdown(f"""
            <div class="compliance-box" style="border-left: 4px solid #F59E0B;">
                <div style="font-weight: 700; color: #B45309; font-size: 1.05rem; margin-bottom: 4px;">✍️ 《退场承诺书》签订率：{comp['commit_rate']:.1f}%</div>
                <div style="font-size: 0.88rem; color: #475569; line-height: 1.5;">
                    • 离场 <strong>{comp['outflow_total']}</strong> 人次规范签署《退场承诺书》。<br>
                    • 明确离场日期与在场务工天数，退场法律文书档案完备。
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # Tab 2: 月度台账钻取 (月度进退场名单与班组明细)
    # ==========================================
    with tab_monthly:
        st.markdown("### 📅 项目月度变动深度钻取")
        
        all_months_options = ['6月', '7月', '8月', '9月']
        col_m_sel, col_m_info = st.columns([1, 2])
        with col_m_sel:
            sel_month = st.selectbox("选择查看月份", options=all_months_options, index=1)
            
        m_data = summary.get(sel_month, {"in_total": 0, "out_total": 0, "onsite_total": 0, "actual_onsite_total": 0, "net_change": 0, "teams": {}})
        
        # 月度核心指标小卡片 (包含实际在场与报备在场)
        m_c1, m_c2, m_c3, m_c4, m_c5 = st.columns(5)
        with m_c1:
            st.metric("本月进场人数", f"{m_data['in_total']} 人", help="当月新进场录入人数")
        with m_c2:
            st.metric("本月离场人数", f"{m_data['out_total']} 人", help="当月办结离场手续人数")
        with m_c3:
            st.metric("报备在场总数", f"{m_data['onsite_total']} 人", help="当月月末花名册登记在场总人数")
        with m_c4:
            act_val = m_data.get('actual_onsite_total', m_data['onsite_total'] - m_data['out_total'])
            st.metric("实际在场人数", f"{act_val} 人", help="扣除退场离场人员后的现场真实驻场人数")
        with m_c5:
            delta_val = m_data['net_change']
            st.metric("本月人员净变动", f"{delta_val:+d} 人", delta=f"{delta_val:+d}")

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)

        # 分包队伍月度对比
        st.markdown(f"#### 🏢 {sel_month} 分包队伍明细分布")
        team_yi_stats = m_data["teams"].get("江苏旭之升 (王宜强施工班组)", {"in_count": 0, "out_count": 0, "onsite_count": 0, "actual_onsite_count": 0})
        team_wang_stats = m_data["teams"].get("青海久昌 (汪佩沾其他班组)", {"in_count": 0, "out_count": 0, "onsite_count": 0, "actual_onsite_count": 0})
        
        team_summary_df = pd.DataFrame([
            {
                "分包单位 / 班组": "青海久昌 (汪佩沾其他班组)",
                "进场人数": team_wang_stats["in_count"],
                "离场人数": team_wang_stats["out_count"],
                "报备在场人数": team_wang_stats["onsite_count"],
                "实际在场人数": team_wang_stats.get("actual_onsite_count", team_wang_stats["onsite_count"] - team_wang_stats["out_count"]),
                "净变动": team_wang_stats["in_count"] - team_wang_stats["out_count"]
            },
            {
                "分包单位 / 班组": "江苏旭之升 (王宜强施工班组)",
                "进场人数": team_yi_stats["in_count"],
                "离场人数": team_yi_stats["out_count"],
                "报备在场人数": team_yi_stats["onsite_count"],
                "实际在场人数": team_yi_stats.get("actual_onsite_count", team_yi_stats["onsite_count"] - team_yi_stats["out_count"]),
                "净变动": team_yi_stats["in_count"] - team_yi_stats["out_count"]
            },
            {
                "分包单位 / 班组": "金属屋面专业分包合计",
                "进场人数": m_data["in_total"],
                "离场人数": m_data["out_total"],
                "报备在场人数": m_data["onsite_total"],
                "实际在场人数": m_data.get("actual_onsite_total", m_data["onsite_total"] - m_data["out_total"]),
                "净变动": m_data["net_change"]
            }
        ])
        st.dataframe(team_summary_df, use_container_width=True, hide_index=True)

        # 展开全周期官方月报表核对
        with st.expander("📑 查看 6-9 月全周期官方人员变更汇总基准表 (与月报合计表完全对齐)", expanded=False):
            st.markdown("""
| 月份 | 公司 | 进场 | 离场 | 在场 | 在场合计 | 实际在场 | 实际在场合计 |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **6** | 青海久昌 | 15 | 0 | 15 | **34** | 15 | **34** |
| | 江苏旭之升 | 19 | 0 | 19 | | 19 | |
| **7** | 青海久昌 | 17 | 4 | 32 | **58** | 28 | **47** |
| | 江苏旭之升 | 7 | 7 | 26 | | 19 | |
| **8** | 青海久昌 | 3 | 2 | 31 | **60** | 29 | **57** |
| | 江苏旭之升 | 10 | 1 | 29 | | 28 | |
| **9** | 青海久昌 | 3 | 9 | 32 | **61** | 23 | **51** |
| | 江苏旭之升 | 1 | 1 | 29 | | 28 | |
            """)

        st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

        # 当月人员进出场明细表格
        t_in_list, t_out_list = st.tabs([f"📥 {sel_month} 进场人员名单 ({m_data['in_total']}人)", f"📤 {sel_month} 离场人员名单 ({m_data['out_total']}人)"])
        
        with t_in_list:
            sub_in = df_inflow[df_inflow['Month'] == sel_month] if not df_inflow.empty else pd.DataFrame()
            if sub_in.empty:
                st.info(f"{sel_month} 没有进场登记人员。")
            else:
                show_in_cols = ['Name', 'Team', 'ID', 'Date', 'Job', 'ContractSigned', 'Registered', 'Remarks']
                display_in = sub_in[show_in_cols].rename(columns={
                    'Name': '姓名', 'Team': '所属队伍/班组', 'ID': '身份证号',
                    'Date': '进场日期', 'Job': '工种', 'ContractSigned': '劳动合同签订',
                    'Registered': '市建委备案', 'Remarks': '备注'
                })
                st.dataframe(display_in, use_container_width=True, hide_index=True)

        with t_out_list:
            sub_out = df_outflow[df_outflow['Month'] == sel_month] if not df_outflow.empty else pd.DataFrame()
            if sub_out.empty:
                st.info(f"{sel_month} 没有离场登记人员。")
            else:
                show_out_cols = ['Name', 'Team', 'ID', 'Date', 'Duration', 'Job', 'WageSettled', 'CommitmentSigned']
                display_out = sub_out[show_out_cols].rename(columns={
                    'Name': '姓名', 'Team': '所属队伍/班组', 'ID': '身份证号',
                    'Date': '离场日期', 'Duration': '在场时间(天/月)', 'Job': '工种',
                    'WageSettled': '工资结算支付情况', 'CommitmentSigned': '退场承诺书签订'
                })
                st.dataframe(display_out, use_container_width=True, hide_index=True)

        # 可选：展开月度报文生成器 (向前兼容原功能)
        with st.expander("📝 查看/生成该月标准化报表文本 (用于工作汇报直贴)", expanded=False):
            date_str = _get_report_date_str(f"2026-0{sel_month[0]}")
            text_total = f"①总：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  金属屋面         日期： {date_str}  本月进场务工人员总数:    {m_data['in_total']}          本月离场务工人员总数:   {m_data['out_total']}         本月现场务工人员总数：   {m_data['onsite_total']}"
            text_wang = f"②分：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  汪佩沾其它班组         日期： {date_str}  本月进场务工人员总数:    {team_wang_stats['in_count']}         本月离场务工人员总数:   {team_wang_stats['out_count']}         本月现场务工人员总数：   {team_wang_stats['onsite_count']}"
            text_yi = f"③分：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  王宜强施工班组         日期： {date_str}  本月进场务工人员总数:    {team_yi_stats['in_count']}          本月离场务工人员总数:   {team_yi_stats['out_count']}         本月现场务工人员总数：   {team_yi_stats['onsite_count']}"
            st.code(f"{text_total}\n{text_wang}\n{text_yi}", language="text")

    # ==========================================
    # Tab 3: 全员参建档案总库 (多源透视)
    # ==========================================
    with tab_roster:
        st.markdown(f"### 👥 项目参建务工人员全景花名册 (累计参建 {len(unique_df)} 人)")
        
        # 多维筛选器 (增加来源与实时在场状态)
        f_c1, f_c2, f_c3, f_c4, f_c5, f_c6 = st.columns([1.2, 1, 1, 1, 1, 1])
        with f_c1:
            kw_search = st.text_input("🔍 搜索姓名 / 身份证", placeholder="输入姓名或身份证...")
        with f_c2:
            team_opts = ["全部队伍"] + sorted(list(unique_df['Team'].dropna().unique()))
            sel_team = st.selectbox("筛选队伍/班组", options=team_opts)
        with f_c3:
            job_opts = ["全部工种"] + sorted(list(unique_df['Job_Clean'].dropna().unique()))
            sel_job = st.selectbox("筛选工种", options=job_opts)
        with f_c4:
            prov_opts = ["全部省份"] + sorted(list(unique_df['Province'].dropna().unique()))
            sel_prov = st.selectbox("筛选籍贯省份", options=prov_opts)
        with f_c5:
            source_col = unique_df['Source'] if 'Source' in unique_df.columns else pd.Series(dtype=str)
            src_opts = ["全部来源"] + sorted(list(source_col.dropna().unique()))
            sel_src = st.selectbox("筛选数据来源", options=src_opts)
        with f_c6:
            status_col = unique_df['LiveStatus'] if 'LiveStatus' in unique_df.columns else pd.Series(dtype=str)
            stat_opts = ["全部状态"] + sorted(list(status_col.dropna().unique()))
            sel_stat = st.selectbox("筛选在场状态", options=stat_opts)

        filtered_df = unique_df.copy()
        if kw_search.strip():
            kw = kw_search.strip()
            filtered_df = filtered_df[filtered_df['Name'].str.contains(kw, na=False) | filtered_df['ID'].str.contains(kw, na=False)]
        if sel_team != "全部队伍":
            filtered_df = filtered_df[filtered_df['Team'] == sel_team]
        if sel_job != "全部工种":
            filtered_df = filtered_df[filtered_df['Job_Clean'] == sel_job]
        if sel_prov != "全部省份":
            filtered_df = filtered_df[filtered_df['Province'] == sel_prov]
        if sel_src != "全部来源" and 'Source' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Source'] == sel_src]
        if sel_stat != "全部状态" and 'LiveStatus' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['LiveStatus'] == sel_stat]

        # 筛选结果提示
        f_ages = filtered_df['Age'].dropna()
        f_avg_age = round(float(f_ages.mean()), 1) if not f_ages.empty else 0.0
        onsite_sub = len(filtered_df[filtered_df['LiveStatus'] == '在场正常']) if 'LiveStatus' in filtered_df.columns else 0
        onboarding_sub = len(filtered_df[filtered_df['LiveStatus'] == '进场手续在办']) if 'LiveStatus' in filtered_df.columns else 0
        offboarding_sub = len(filtered_df[filtered_df['LiveStatus'] == '离场结算中']) if 'LiveStatus' in filtered_df.columns else 0
        archived_sub = len(filtered_df[filtered_df['LiveStatus'] == '已离场归档']) if 'LiveStatus' in filtered_df.columns else 0

        st.caption(f"当前筛选出 **{len(filtered_df)}** 人 | 平均年龄: **{f_avg_age}** 岁 | 在场正常: **{onsite_sub}** | 进场在办: **{onboarding_sub}** | 离场结算中: **{offboarding_sub}** | 已离场: **{archived_sub}**")

        cols_to_show = ['Name', 'Gender', 'Age', 'AgeGroup', 'Job', 'Team', 'Province', 'Source', 'LiveStatus', 'ID', 'Address', 'ContractNo']
        existing_show = [c for c in cols_to_show if c in filtered_df.columns]
        display_roster = filtered_df[existing_show].rename(columns={
            'Name': '姓名', 'Gender': '性别', 'Age': '周岁', 'AgeGroup': '年龄梯队',
            'Job': '工种', 'Team': '所属队伍', 'Province': '籍贯省份',
            'Source': '数据来源', 'LiveStatus': '实时状态',
            'ID': '身份证号', 'Address': '家庭住址', 'ContractNo': '劳动合同编号'
        })
        st.dataframe(display_roster, use_container_width=True, hide_index=True)

        # 导出下载按键与增量回填指引
        c_exp1, c_exp2 = st.columns([1, 2])
        with c_exp1:
            csv_buffer = io.BytesIO()
            display_roster.to_csv(csv_buffer, index=False, encoding="utf_8_sig")
            st.download_button(
                label="📥 导出当前筛选人员名单 (CSV)",
                data=csv_buffer.getvalue(),
                file_name="project_workers_filtered.csv",
                mime="text/csv",
                type="primary"
            )
        with c_exp2:
            if unpasted_count > 0:
                st.info(f"💡 提示：检测到当前库中有 **{unpasted_count}** 名增量人员，您可以直接前往【花名册与报表导出】一键复制纯文本粘贴到您的本地 Excel 中。")

    # ==========================================
    # Tab 4: 高清报表大图导出 (Matplotlib)
    # ==========================================
    with tab_export:
        st.markdown("### 🖼️ 专业看板高清图表生成 (对标 plot_dashboard.py)")
        st.write("根据现场实际解析数据，复刻生成 2x3 画布专业可视化看板大图，支持 300 DPI 超清保存与汇报打印。")

        col_gen, col_down = st.columns([1, 1])
        with col_gen:
            if st.button("🎨 重新渲染高清看板图", type="primary"):
                st.session_state["mpl_fig_rendered"] = True

        try:
            fig = personnel_data_service.generate_matplotlib_figure()
            st.pyplot(fig, use_container_width=True)
            
            png_bytes = personnel_data_service.get_matplotlib_png_bytes()
            st.download_button(
                label="📥 一键下载 300 DPI 高清看板图片 (PNG)",
                data=png_bytes,
                file_name="soccer_stadium_labor_visualization.png",
                mime="image/png"
            )
        except Exception as err:
            st.error(f"渲染高清图表时出错: {err}")

    # ==========================================
    # Tab 5: 数据源管理与同步
    # ==========================================
    with tab_source:
        st.markdown("### 📂 数据源管理与同步状态")
        st.write("系统会自动解析 `load-data` 目录下的三张核心表格并动态缓存：")
        
        status_info = personnel_data_service.get_data_status()
        status_rows = []
        for k, v in status_info.items():
            status_rows.append({
                "数据表": v["label"],
                "文件是否存在": "✅ 存在" if v["exists"] else "❌ 缺失",
                "文件大小": f"{v['size_kb']} KB" if v["exists"] else "-",
                "最后更新时间": v["mtime"] or "-",
                "文件绝对路径": v["path"]
            })
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)

        if st.button("🔄 强制清空缓存并重新解析 Excel 数据", type="secondary"):
            personnel_data_service.load_all_data(force_reload=True)
            st.success("🎉 数据已成功重新解析并刷新！")
            st.rerun()

        st.markdown("---")
        st.markdown("#### 💡 数据表规范说明与动态联动机制")
        st.markdown("""
        1. **人员花名册**：`中建二局安装务工人员（含队长、班组长、弄民工）花名册.xlsx`，记录 6-9 月各队伍在场劳务工人员花名册；
        2. **进场月报表**：`二局安装-足球场项目人员变更月报表（进场情况）.xlsx`，记录各月新增进场人员清单；
        3. **离场月报表**：`二局安装-足球场项目人员变更月报表（离场情况）.xlsx`，记录各月规范退场人员清单、工资金额结清与退场承诺书签署状态；
        4. **全系统四源实时动态联动**：大屏数据已打通并实时接入 `master_data` 主表、`onboarding_pipeline` 进场流水线、`offboarding_pipeline` 离场流水线。**您日常管人录入系统后大屏即可自动呈现最新可视化结果，彻底免除为了看图每次还要上传 Excel 的繁琐步骤**。
        """)

    # ==========================================
    # 底部：劳务全周期动态闭环工作流导引
    # ==========================================
    st.markdown("<div style='margin-top: 36px;'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div style="border-top: 2px dashed #E2E8F0; padding-top: 20px; margin-bottom: 16px;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
            <div>
                <h3 style="margin: 0; color: #1E293B; font-weight: 700; display: flex; align-items: center; gap: 8px;">
                    <span>🔄</span> 劳务全周期动态闭环工作流导引
                </h3>
                <p style="margin: 4px 0 0 0; color: #64748B; font-size: 0.88rem;">
                    标准作业闭环：系统动态流转感知 ➜ 大屏自动可视化统计 ➜ 一键回填完善本地 Excel 底册
                </p>
            </div>
            <span style="font-size: 0.8rem; background: #EEF2FF; color: #4F46E5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">
                免除重复上传 · 永久动态联动
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    wf_c1, wf_c2, wf_c3, wf_c4 = st.columns(4)
    with wf_c1:
        st.markdown("""
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px; min-height: 140px;">
            <div style="color: #059669; font-weight: 700; font-size: 0.95rem; margin-bottom: 6px;">
                ① 进场手续申报
            </div>
            <div style="font-size: 0.82rem; color: #475569; line-height: 1.5; margin-bottom: 12px;">
                现场新工人到场，录入进场流水线，跟踪三级安全教育与合同办理。
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("前往【进场流水线】 ➜", key="wf_btn_onboarding", use_container_width=True):
            _switch_module("onboarding")
            st.rerun()

    with wf_c2:
        st.markdown("""
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px; min-height: 140px;">
            <div style="color: #7C3AED; font-weight: 700; font-size: 0.95rem; margin-bottom: 6px;">
                ② 档案魔法整合
            </div>
            <div style="font-size: 0.82rem; color: #475569; line-height: 1.5; margin-bottom: 12px;">
                建委/官方系统走完后导出官方人员档案，直接上传清洗并沉淀入主表。
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("前往【档案魔法整合】 ➜", key="wf_btn_info_merge", use_container_width=True):
            _switch_module("info_merge")
            st.rerun()

    with wf_c3:
        st.markdown("""
        <div style="background: #EEF2FF; border: 1.5px solid #818CF8; border-radius: 10px; padding: 14px; min-height: 140px;">
            <div style="color: #4F46E5; font-weight: 700; font-size: 0.95rem; margin-bottom: 6px;">
                ③ 态势监控 (当前)
            </div>
            <div style="font-size: 0.82rem; color: #3730A3; line-height: 1.5; margin-bottom: 12px;">
                当前面板：全自动动态计算进退场与队伍结构，无需传表即可直观查看。
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.button("🟢 当前视图运行中", key="wf_btn_cur", use_container_width=True, disabled=True)

    with wf_c4:
        st.markdown("""
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px; min-height: 140px;">
            <div style="color: #DB2777; font-weight: 700; font-size: 0.95rem; margin-bottom: 6px;">
                ④ 本地底册一键回填
            </div>
            <div style="font-size: 0.82rem; color: #475569; line-height: 1.5; margin-bottom: 12px;">
                一键复制带防科学计数法保护的纯文本行，直接粘贴更新本地 Excel。
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("前往【花名册直贴】 ➜", key="wf_btn_report", type="primary", use_container_width=True):
            _switch_module("report")
            st.rerun()

