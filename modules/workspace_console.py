import streamlit as st
import pandas as pd
from datetime import date
from services.task_engine import get_all_tasks
from modules.master_data import load_master_df
from modules.onboarding_pipeline import onboarding_service
from modules.offboarding_pipeline import offboarding_service

def get_tasks():
    onboarding_data = onboarding_service.get_pending_workers()
    offboarding_data = offboarding_service.get_pending_workers()
    master_df = load_master_df()
    
    return get_all_tasks(onboarding_data, offboarding_data, master_df)

def handle_task_click(module_name, worker_id):
    st.session_state.current_nav = module_name
    st.session_state.target_worker_id = worker_id

def render():
    st.markdown("""
    <style>
    .task-row {
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 8px;
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        transition: all 0.2s;
    }
    .task-row:hover {
        background-color: #f1f5f9;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .task-status {
        color: #64748b;
        font-size: 0.9em;
        background: #e2e8f0;
        padding: 2px 8px;
        border-radius: 12px;
    }
    </style>
    <div class="page-header-deco">
        <span class="material-symbols-outlined" style="font-size:32px;color:#6366F1;">dashboard</span>
        <div class="header-text"><h2>今日办公中控台</h2>
        <p>任务驱动式工作台 · 告诉我下一步该做什么</p></div>
    </div>
    <div class="color-strip" style="background:linear-gradient(90deg,#C7D2FE,#EDE9FE);"></div>
    """, unsafe_allow_html=True)
    
    tasks_red, tasks_orange, tasks_yellow, anomalies = get_tasks()
    
    # 🚨 逻辑异常区
    if anomalies:
        st.markdown(f"### 🚨 逻辑异常与漏项检查 ({len(anomalies)})")
        for task in anomalies:
            st.error(f"**{task['name']}** ({task['status']}) - **异常**: {task['anomaly']}")
            if st.button("进入处理", key=f"anomaly_{task['worker_id']}_{task['action']}", help="跳转到对应人员处理", type="primary"):
                handle_task_click(task['type'], task['worker_id'])
                st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("### 📋 今日待办清单")
    
    # 🔴 我现在需要处理
    with st.container():
        st.markdown(f"#### 🔴 我现在需要处理 ({len(tasks_red)})")
        if not tasks_red:
            st.info("太棒了！当前没有需要您亲自处理的紧急待办事项。")
        else:
            for task in tasks_red:
                with st.container():
                    col1, col2, col3, col4 = st.columns([2, 1.5, 4, 1.5])
                    with col1:
                        st.markdown(f"**{task['name']}** | {task['team']}")
                    with col2:
                        st.markdown(f"<span class='task-status'>{task['status']}</span>", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"**下一步：**<span style='color: #ef4444; font-weight: 500;'>{task['action']}</span>", unsafe_allow_html=True)
                    with col4:
                        if st.button("🚀 去处理", key=f"red_{task['worker_id']}_{task['action']}", use_container_width=True):
                            handle_task_click(task['type'], task['worker_id'])
                            st.rerun()
                    st.markdown("<hr style='margin: 0.3em 0; border: none; border-bottom: 1px solid #f1f5f9;'>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # 🟠 等别人处理
    with st.container():
        st.markdown(f"#### 🟠 等别人处理 ({len(tasks_orange)})")
        if not tasks_orange:
            st.info("当前没有等待他人处理的事项。")
        else:
            for task in tasks_orange:
                with st.container():
                    col1, col2, col3, col4 = st.columns([2, 1.5, 4, 1.5])
                    with col1:
                        st.markdown(f"**{task['name']}** | {task['team']}")
                    with col2:
                        st.markdown(f"<span class='task-status'>{task['status']}</span>", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"**等待：**<span style='color: #f59e0b; font-weight: 500;'>{task['action']}</span>", unsafe_allow_html=True)
                    with col4:
                        if st.button("去跟进", key=f"orange_{task['worker_id']}_{task['action']}", use_container_width=True):
                            handle_task_click(task['type'], task['worker_id'])
                            st.rerun()
                    st.markdown("<hr style='margin: 0.3em 0; border: none; border-bottom: 1px solid #f1f5f9;'>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # 🟡 等官方系统/数据
    with st.container():
        st.markdown(f"#### 🟡 等官方系统/数据 ({len(tasks_yellow)})")
        if not tasks_yellow:
            st.info("当前没有需要同步或等待官方系统的特殊事项。")
        else:
            for task in tasks_yellow:
                with st.container():
                    col1, col2, col3, col4 = st.columns([2, 1.5, 4, 1.5])
                    with col1:
                        st.markdown(f"**{task['name']}** | {task['team']}")
                    with col2:
                        st.markdown(f"<span class='task-status'>{task['status']}</span>", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"**下一步：**<span style='color: #3b82f6; font-weight: 500;'>{task['action']}</span>", unsafe_allow_html=True)
                    with col4:
                        if st.button("去核对", key=f"yellow_{task['worker_id']}_{task['action']}", use_container_width=True):
                            handle_task_click(task['type'], task['worker_id'])
                            st.rerun()
                    st.markdown("<hr style='margin: 0.3em 0; border: none; border-bottom: 1px solid #f1f5f9;'>", unsafe_allow_html=True)
