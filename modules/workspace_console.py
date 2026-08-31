# modules/workspace_console.py
import streamlit as st
import pandas as pd
from datetime import date
from services.task_engine import get_all_tasks
from modules.master_data import load_master_df
from modules.onboarding_pipeline import onboarding_service, worker_dialog as onboarding_dialog
from modules.offboarding_pipeline import offboarding_service, worker_dialog as offboarding_dialog

def get_tasks():
    onboarding_data = onboarding_service.get_pending_workers()
    offboarding_data = offboarding_service.get_pending_workers()
    master_df = load_master_df()
    
    return get_all_tasks(onboarding_data, offboarding_data, master_df)

def render_task_card(task, category, key_prefix):
    worker_id = task.get("worker_id", "")
    name = task.get("name", "未知")
    team = task.get("team", "未知")
    pipeline_type = task.get("type", "onboarding")
    status = task.get("status", "")
    action = task.get("action", "")
    anomaly = task.get("anomaly", None)
    
    is_onboard = pipeline_type == "onboarding"
    pipe_badge = '<span class="todo-tag tag-pipeline-in">🟢 进场</span>' if is_onboard else '<span class="todo-tag tag-pipeline-out">🔴 离场</span>'
    team_badge = f'<span class="todo-tag tag-team">{team}</span>'
    
    if category == "red":
        card_cls = "card-red"
        pill_cls = "status-pill-red"
        pill_text = "我需处理"
        callout_label = "⚡ 办理指令"
        callout_text = action
        btn_label = "🚀 立即办理"
        btn_type = "primary"
    elif category == "orange":
        card_cls = "card-orange"
        pill_cls = "status-pill-orange"
        pill_text = "等待协同"
        callout_label = "⏳ 正在等待"
        callout_text = action
        btn_label = "👀 跟进办理"
        btn_type = "secondary"
    elif category == "yellow":
        card_cls = "card-yellow"
        pill_cls = "status-pill-yellow"
        pill_text = "官方核对"
        callout_label = "🔄 待同步项"
        callout_text = action
        btn_label = "📊 查看办理"
        btn_type = "secondary"
    else: # anomaly
        card_cls = "card-anomaly"
        pill_cls = "status-pill-anomaly"
        pill_text = "逻辑冲突"
        callout_label = "⚠️ 预警原因"
        callout_text = anomaly if anomaly else action
        btn_label = "🛠️ 立即排查"
        btn_type = "primary"
        
    with st.container():
        st.markdown(f"""
        <div class="todo-card {card_cls}">
            <div class="todo-card-header">
                <div class="todo-worker-name">
                    <span>{name}</span>
                    {team_badge}
                </div>
                <div>{pipe_badge}</div>
            </div>
            <div class="todo-action-callout">
                <span class="action-callout-label">{callout_label}</span>
                <span class="action-callout-text">{callout_text}</span>
            </div>
            <div class="todo-meta-footer">
                <span>📍 状态: {status}</span>
                <span class="todo-status-pill {pill_cls}">{pill_text}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        action_key = f"{key_prefix}_{worker_id}_{category}_{abs(hash(action)) % 100000}"
        if st.button(btn_label, key=action_key, use_container_width=True, type=btn_type):
            if pipeline_type == "onboarding":
                onboarding_dialog(worker_id)
            else:
                offboarding_dialog(worker_id)

def render():
    # 顶部标题栏
    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined header-emoji" style="color:#6366F1;">dashboard</span>
        <div class="header-text">
            <h2>今日办公中控台</h2>
            <p>任务驱动式工作台 · 聚焦重点待办与全程状态监控</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)
    
    # 监控逻辑提示横幅
    st.markdown("""
    <div style="background: linear-gradient(135deg, #EEF2FF 0%, #F5F3FF 100%); border-radius: 12px; padding: 10px 16px; margin-bottom: 16px; border: 1px solid #E0E7FF; display: flex; align-items: center; justify-content: space-between; font-size: 13px; color: #4338CA;">
        <div style="display:flex; align-items:center; gap:8px;">
            <span class="material-symbols-outlined" style="font-size:18px;">sync_saved_locally</span>
            <span><b>全流程状态实时监控中</b>：点击【立即办理】直接弹窗修改手续，系统会根据资料勾选进度<b>自动流转监控状态</b>（🔴我需处理 ➔ 🟠等待协同 ➔ 🟡官方核对），直至全部手续闭环！</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    tasks_red, tasks_orange, tasks_yellow, anomalies = get_tasks()
    
    # 1. 顶部指标概览栏 (Metric Ribbon)
    st.markdown(f"""
    <div class="todo-stats-bar">
        <div class="todo-stat-card stat-red">
            <div class="todo-stat-info">
                <span class="todo-stat-title"><span class="material-symbols-outlined" style="font-size:16px;color:#E11D48;">priority_high</span> 我需处理</span>
                <span class="todo-stat-num">{len(tasks_red)}</span>
            </div>
            <span class="material-symbols-outlined" style="font-size:30px;color:rgba(225,29,72,0.18);">assignment_late</span>
        </div>
        <div class="todo-stat-card stat-orange">
            <div class="todo-stat-info">
                <span class="todo-stat-title"><span class="material-symbols-outlined" style="font-size:16px;color:#EA580C;">hourglass_top</span> 等待协同</span>
                <span class="todo-stat-num">{len(tasks_orange)}</span>
            </div>
            <span class="material-symbols-outlined" style="font-size:30px;color:rgba(234,88,12,0.18);">pending_actions</span>
        </div>
        <div class="todo-stat-card stat-yellow">
            <div class="todo-stat-info">
                <span class="todo-stat-title"><span class="material-symbols-outlined" style="font-size:16px;color:#2563EB;">sync_alt</span> 官方核对</span>
                <span class="todo-stat-num">{len(tasks_yellow)}</span>
            </div>
            <span class="material-symbols-outlined" style="font-size:30px;color:rgba(37,99,235,0.18);">cloud_sync</span>
        </div>
        <div class="todo-stat-card stat-anomaly">
            <div class="todo-stat-info">
                <span class="todo-stat-title"><span class="material-symbols-outlined" style="font-size:16px;color:#DB2777;">warning</span> 异常预警</span>
                <span class="todo-stat-num">{len(anomalies)}</span>
            </div>
            <span class="material-symbols-outlined" style="font-size:30px;color:rgba(219,39,119,0.18);">report_problem</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 2. 搜索与筛选工具条
    col_search, col_pipe = st.columns([3, 1])
    with col_search:
        search_query = st.text_input("🔍 搜索人员/班组/待办事项", placeholder="输入姓名、班组或操作关键词即时筛选...", label_visibility="collapsed")
    with col_pipe:
        pipeline_filter = st.selectbox("流水线筛选", ["全部流水线", "进场流水", "离场流水"], label_visibility="collapsed")
    
    # 过滤函数
    def filter_task_list(task_list):
        res = []
        for t in task_list:
            match_search = True
            if search_query:
                q = search_query.strip().lower()
                name_match = q in t.get("name", "").lower()
                team_match = q in t.get("team", "").lower()
                action_match = q in t.get("action", "").lower()
                anomaly_match = q in str(t.get("anomaly", "")).lower()
                status_match = q in t.get("status", "").lower()
                match_search = name_match or team_match or action_match or anomaly_match or status_match
            
            match_pipe = True
            if pipeline_filter == "进场流水" and t.get("type") != "onboarding":
                match_pipe = False
            elif pipeline_filter == "离场流水" and t.get("type") != "offboarding":
                match_pipe = False
                
            if match_search and match_pipe:
                res.append(t)
        return res

    f_red = filter_task_list(tasks_red)
    f_orange = filter_task_list(tasks_orange)
    f_yellow = filter_task_list(tasks_yellow)
    f_anomalies = filter_task_list(anomalies)

    # 3. 多维度 Tab 分类视图（默认聚焦优先项“🔴 我需处理”，绝不杂糅展示）
    tab_red, tab_orange, tab_yellow, tab_anomaly, tab_kanban = st.tabs([
        f"🔴 我需处理 ({len(f_red)})",
        f"🟠 等待协同 ({len(f_orange)})",
        f"🟡 官方核对 ({len(f_yellow)})",
        f"🚨 逻辑预警 ({len(f_anomalies)})",
        "📊 三栏全景看板"
    ])
    
    def render_grid(items, cat, prefix="grid"):
        if not items:
            st.markdown("""
            <div class="todo-empty-box">
                <span class="material-symbols-outlined empty-icon" style="color:#10B981;">check_circle</span>
                <h4>该分类下暂无待办事项</h4>
                <p>一切井然有序，没有需要处理的事项。</p>
            </div>
            """, unsafe_allow_html=True)
            return

        cols = st.columns(3)
        for idx, t in enumerate(items):
            with cols[idx % 3]:
                render_task_card(t, cat, f"{prefix}_{idx}")

    with tab_red:
        render_grid(f_red, "red", prefix="red")

    with tab_orange:
        render_grid(f_orange, "orange", prefix="orange")

    with tab_yellow:
        render_grid(f_yellow, "yellow", prefix="yellow")

    with tab_anomaly:
        render_grid(f_anomalies, "anomaly", prefix="anomaly")

    with tab_kanban:
        # 三栏看板视图：左（我需处理）、中（等待协同）、右（官方核对）
        k_col1, k_col2, k_col3 = st.columns(3)
        
        with k_col1:
            st.markdown(f"""
            <div class="kanban-col-header kanban-header-red">
                <span>🔴 我需处理</span>
                <span class="kanban-count-badge">{len(f_red)}</span>
            </div>
            """, unsafe_allow_html=True)
            if not f_red:
                st.markdown("<p style='color:#94A3B8; font-size:13px; text-align:center; padding:12px;'>暂无紧急待办</p>", unsafe_allow_html=True)
            else:
                for idx, t in enumerate(f_red):
                    render_task_card(t, "red", f"kanban_red_{idx}")

        with k_col2:
            st.markdown(f"""
            <div class="kanban-col-header kanban-header-orange">
                <span>🟠 等待协同</span>
                <span class="kanban-count-badge">{len(f_orange)}</span>
            </div>
            """, unsafe_allow_html=True)
            if not f_orange:
                st.markdown("<p style='color:#94A3B8; font-size:13px; text-align:center; padding:12px;'>暂无等待事项</p>", unsafe_allow_html=True)
            else:
                for idx, t in enumerate(f_orange):
                    render_task_card(t, "orange", f"kanban_orange_{idx}")

        with k_col3:
            st.markdown(f"""
            <div class="kanban-col-header kanban-header-yellow">
                <span>🟡 官方核对</span>
                <span class="kanban-count-badge">{len(f_yellow)}</span>
            </div>
            """, unsafe_allow_html=True)
            if not f_yellow:
                st.markdown("<p style='color:#94A3B8; font-size:13px; text-align:center; padding:12px;'>暂无核对事项</p>", unsafe_allow_html=True)
            else:
                for idx, t in enumerate(f_yellow):
                    render_task_card(t, "yellow", f"kanban_yellow_{idx}")

    st.markdown("<div style='margin-bottom: 24px;'></div>", unsafe_allow_html=True)
