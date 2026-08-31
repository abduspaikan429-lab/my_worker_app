# main.py
import streamlit as st
import pandas as pd
from utils.style_loader import load_css
from modules import info_merge, pro_data_processing, onboarding_pipeline, offboarding_pipeline, report_generator, attendance_payroll, personnel_dashboard, workspace_console
from modules.master_data import load_master_df
from modules.offboarding_pipeline import load_offboarding_history

# 1. 全局配置
st.set_page_config(
    page_title="建筑劳务综合管理平台",
    layout="wide"
)

# 2. 注入高质感极简 CSS
load_css()

# 3. 顶部欢迎 Banner (动态雷达)
from datetime import date
cached_count = 0
onsite_count = 0
onboarding_count = 0
offboarding_count = 0
today_onboard = 0
today_offboard = 0
pending_tasks_count = 0

try:
    persisted_master = load_master_df()
    
    # 1. 在场人数
    if not persisted_master.empty:
        all_workers = onboarding_pipeline.onboarding_service.merge_with_master(persisted_master)
        onsite_df = offboarding_pipeline.offboarding_service.filter_onsite_df(all_workers)
        onsite_count = len(onsite_df)
    
    today_str = str(date.today())
    
    # 2. 进场人数
    onboarding_data = onboarding_pipeline.onboarding_service.get_pending_workers()
    onboarding_count = len(onboarding_data)
    today_onboard = sum(1 for d in onboarding_data.values() if d.get("created_at") == today_str)
    
    # 3. 离场人数
    offboarding_data = offboarding_pipeline.offboarding_service.get_pending_workers()
    offboarding_count = len(offboarding_data)
    
    history = load_offboarding_history()
    today_offboard = sum(1 for d in history if d.get("离场日期", "") == today_str or d.get("info", {}).get("离场日期") == today_str)
    
    # 4. 待办总数
    from modules.workspace_console import get_tasks
    tasks_red, tasks_orange, tasks_yellow, anomalies = get_tasks()
    distinct_workers = {t['worker_id'] for t in (tasks_red + tasks_orange + tasks_yellow + anomalies)}
    pending_tasks_count = len(distinct_workers)

except Exception as e:
    print(f"Stats Error: {e}")

# 模块配置：名称、图标、颜色映射
MODULES_CONFIG = {
    "info_merge": {
        "name": "档案魔法整合",
        "icon": "",
        "material_icon": ":material/folder_managed:",
        "color": "#B3A4F3",
        "gradient": "linear-gradient(135deg, #EEE7FF 0%, #E0E7FF 100%)",
        "description": "多系统档案数据智能清洗与合并",
    },
    "pro_data": {
        "name": "专业数据处理",
        "icon": "",
        "material_icon": ":material/analytics:",
        "color": "#60A5FA",
        "gradient": "linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%)",
        "description": "檩条分轴合并、重量计算等专业工具集",
    },
    "onboarding": {
        "name": "进场流水线追踪",
        "icon": "",
        "material_icon": ":material/how_to_reg:",
        "color": "#34D399",
        "gradient": "linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%)",
        "description": "工人进场流程全生命周期追踪管理",
    },
    "offboarding": {
        "name": "离场流水线追踪",
        "icon": "",
        "material_icon": ":material/flight_takeoff:",
        "color": "#F87171",
        "gradient": "linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%)",
        "description": "人员离场与工资结算全流程闭环跟进",
    },
    "report": {
        "name": "花名册与报表",
        "icon": "",
        "material_icon": ":material/contact_page:",
        "color": "#F472B6",
        "gradient": "linear-gradient(135deg, #FDF2F8 0%, #FCE7F3 100%)",
        "description": "花名册快速提取与变更月报一键生成",
    },
    "attendance_payroll": {
        "name": "考勤对账与工资结算",
        "icon": "",
        "material_icon": ":material/payments:",
        "color": "#14B8A6",
        "gradient": "linear-gradient(135deg, #F0FDFA 0%, #CCFBF1 100%)",
        "description": "三方考勤对账、在线定稿与工资一键导出",
    },
    "personnel_dashboard": {
        "name": "人员变更面板",
        "icon": "",
        "material_icon": ":material/bar_chart:",
        "color": "#6366F1",
        "gradient": "linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%)",
        "description": "四源实时联动，在场/进场/离场数据一站式总览",
    },
}

st.markdown(f"""
<div class="welcome-banner" style="display: flex; flex-direction: column; gap: 8px; padding-bottom: 12px;">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div class="welcome-banner-left">
            <span class="welcome-title">Hello, 劳务指挥官</span>
            <span class="welcome-subtitle">劳务管理一站式指挥中心</span>
        </div>
        <div class="welcome-banner-right" style="gap: 12px;">
            <div class="status-badge" style="background: rgba(99, 102, 241, 0.1); color: #4f46e5;">
                <span class="material-symbols-outlined" style="font-size: 16px;">group</span>
                <span>当前在场: <b>{onsite_count}</b></span>
            </div>
            <div class="status-badge" style="background: rgba(52, 211, 153, 0.1); color: #059669;">
                <span class="material-symbols-outlined" style="font-size: 16px;">how_to_reg</span>
                <span>办理进场: <b>{onboarding_count}</b></span>
            </div>
            <div class="status-badge" style="background: rgba(248, 113, 113, 0.1); color: #dc2626;">
                <span class="material-symbols-outlined" style="font-size: 16px;">flight_takeoff</span>
                <span>办理离场: <b>{offboarding_count}</b></span>
            </div>
        </div>
    </div>
    <div style="display: flex; justify-content: flex-end; gap: 15px; margin-top: 2px;">
        <div class="status-badge" style="background: rgba(251, 146, 60, 0.1); color: #ea580c; border: 1px solid rgba(251, 146, 60, 0.2);">
            <span class="material-symbols-outlined" style="font-size: 16px;">assignment_late</span>
            <span>待办总数: <b>{pending_tasks_count}</b></span>
        </div>
        <div class="status-badge" style="background: transparent; padding: 0 4px; color: #64748b;">
            <span>今日新增进场: <b>{today_onboard}</b></span>
        </div>
        <div class="status-badge" style="background: transparent; padding: 0 4px; color: #64748b;">
            <span>今日离场: <b>{today_offboard}</b></span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# 4. 侧边栏
if "current_nav" not in st.session_state:
    st.session_state.current_nav = "workspace_console"
if "pro_data_tool_index" not in st.session_state:
    st.session_state.pro_data_tool_index = 0

is_pro_active = (st.session_state.current_nav == "pro_data")

def select_pro_tool(idx):
    st.session_state.current_nav = "pro_data"
    st.session_state.pro_data_tool_index = idx
    st.session_state._nav_radio = None
    tool_keys = [
        "① 檩条分轴合并与重量计算",
        "② 檩条下料单自动合计与汇总",
        "③ 扩展专业工具 (预留)",
    ]
    if idx < len(tool_keys):
        st.session_state._pro_data_radio = tool_keys[idx]

# 4.1 专业功能独立板块（置于“劳务管理平台”上方）
st.sidebar.markdown("""
<div class="sidebar-pro-badge-header">
    <span class="material-symbols-outlined" style="font-size: 14px; color: #818CF8;">tune</span>
    <span style="font-size: 12px; font-weight: 700; color: #64748B;">专业功能</span>
</div>
""", unsafe_allow_html=True)

pro_col1, pro_col2, pro_col3 = st.sidebar.columns([1, 1, 1])

with pro_col1:
    b1_active = is_pro_active and st.session_state.pro_data_tool_index == 0
    st.button(
        "①",
        key="btn_pro_tool_1",
        use_container_width=True,
        type="primary" if b1_active else "secondary",
        on_click=select_pro_tool,
        args=(0,),
        help="① 檩条分轴合并与重量计算"
    )

with pro_col2:
    b2_active = is_pro_active and st.session_state.pro_data_tool_index == 1
    st.button(
        "②",
        key="btn_pro_tool_2",
        use_container_width=True,
        type="primary" if b2_active else "secondary",
        on_click=select_pro_tool,
        args=(1,),
        help="② 檩条下料单自动合计与汇总"
    )

with pro_col3:
    b3_active = is_pro_active and st.session_state.pro_data_tool_index == 2
    st.button(
        "③",
        key="btn_pro_tool_3",
        use_container_width=True,
        type="primary" if b3_active else "secondary",
        on_click=select_pro_tool,
        args=(2,),
        help="③ 扩展专业工具 (预留)"
    )

# 4.2 侧边栏 - 装饰性标题（劳务管理平台）
st.sidebar.markdown("""
<div class="sidebar-title-deco">
    <span>劳务管理平台</span>
</div>
""", unsafe_allow_html=True)

# 5. 模块导航路由
modules_map = {
    "workspace_console": (":material/dashboard: 今日办公中控台", workspace_console.render),
    "personnel_dashboard": (":material/bar_chart: 人员变更面板", personnel_dashboard.render),
    "info_merge": (":material/folder_managed: 档案魔法整合", info_merge.render),
    "onboarding": (":material/how_to_reg: 进场流水线追踪", onboarding_pipeline.render),
    "offboarding": (":material/flight_takeoff: 离场流水线追踪", offboarding_pipeline.render),
    "report": (":material/contact_page: 花名册与报表导出", report_generator.render),
    "attendance_payroll": (":material/payments: 考勤对账与工资结算", attendance_payroll.render),
    "pro_data": (":material/analytics: 专业数据处理", pro_data_processing.render),
}

# 侧边栏常规主导航项（7项高频劳务管理功能，不占用空间放置低频工具）
main_nav_keys = [
    "workspace_console",
    "personnel_dashboard",
    "info_merge",
    "onboarding",
    "offboarding",
    "report",
    "attendance_payroll",
]

# 侧边栏导航分组
st.sidebar.markdown("""
<div class="deco-dots">
    <span class="dot"></span>
    <span class="dot"></span>
    <span class="dot"></span>
    <span class="dot"></span>
</div>
<b style="color:#64748B; font-size:13px;">功能导航</b>
""", unsafe_allow_html=True)

def on_nav_change():
    if st.session_state._nav_radio is not None:
        st.session_state.current_nav = st.session_state._nav_radio

radio_index = main_nav_keys.index(st.session_state.current_nav) if st.session_state.current_nav in main_nav_keys else None

st.sidebar.radio(
    "功能导航",
    main_nav_keys,
    format_func=lambda k: modules_map[k][0],
    label_visibility="collapsed",
    index=radio_index,
    key="_nav_radio",
    on_change=on_nav_change
)

# 6. 执行选中的模块
if st.session_state.current_nav in modules_map:
    modules_map[st.session_state.current_nav][1]()
else:
    workspace_console.render()

