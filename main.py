# main.py
import streamlit as st
import pandas as pd
from utils.style_loader import load_css
from modules import info_merge, pro_data_processing, onboarding_pipeline, offboarding_pipeline, report_generator, daily_assistant, attendance_payroll, personnel_dashboard
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
cached_count = 0
today_onboarding = 0
onsite_count = 0
try:
    persisted_master = load_master_df()
    today_onboarding = len(onboarding_pipeline.onboarding_service.get_records())
    if not persisted_master.empty:
        cached_count = len(persisted_master)
        # 在场 = (主表 + 进场流程) - 已离场归档 - 正在办理离场结算
        all_workers = onboarding_pipeline.onboarding_service.merge_with_master(persisted_master)
        onsite_df = offboarding_pipeline.offboarding_service.filter_onsite_df(all_workers)
        onsite_count = len(onsite_df)
    elif 'merged_df' in st.session_state and isinstance(st.session_state.merged_df, pd.DataFrame):
        cached_count = len(st.session_state.merged_df)
        all_workers = onboarding_pipeline.onboarding_service.merge_with_master(st.session_state.merged_df)
        onsite_df = offboarding_pipeline.offboarding_service.filter_onsite_df(all_workers)
        onsite_count = len(onsite_df)
    else:
        onboarding_df = onboarding_pipeline.onboarding_service.get_onboarding_df()
        onsite_df = offboarding_pipeline.offboarding_service.filter_onsite_df(onboarding_df)
        onsite_count = len(onsite_df)
except Exception:
    pass

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
    "daily": {
        "name": "微信办公辅助",
        "icon": "",
        "material_icon": ":material/chat_bubble:",
        "color": "#FB923C",
        "gradient": "linear-gradient(135deg, #FFF7ED 0%, #FEF3C7 100%)",
        "description": "考勤日报生成、闪念备忘与疑问管理",
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
<div class="welcome-banner">
    <div class="welcome-banner-left">
        <span class="welcome-title">Hello, 劳务指挥官</span>
        <span class="welcome-subtitle">劳务管理一站式指挥中心</span>
    </div>
    <div class="welcome-banner-right">
        <div class="status-badge">
            <span class="material-symbols-outlined" style="font-size: 16px;">folder</span>
            <span>档案缓存: <b>{cached_count}</b> 人</span>
        </div>
        <div class="status-badge">
            <span class="material-symbols-outlined" style="font-size: 16px;">groups</span>
            <span>当前在场: <b>{onsite_count}</b> 人</span>
        </div>
        <div class="status-badge">
            <span class="material-symbols-outlined" style="font-size: 16px;">rocket_launch</span>
            <span>追踪进场: <b>{today_onboarding}</b> 人</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# 4. 侧边栏 - 装饰性标题
st.sidebar.markdown("""
<div class="sidebar-title-deco">
    <span>劳务管理平台</span>
</div>
""", unsafe_allow_html=True)

# 5. 模块导航路由
modules_map = {
    "personnel_dashboard": (":material/bar_chart: 人员变更面板", personnel_dashboard.render),
    "info_merge": (":material/folder_managed: 档案魔法整合", info_merge.render),
    "pro_data": (":material/analytics: 专业数据处理", pro_data_processing.render),
    "onboarding": (":material/how_to_reg: 进场流水线追踪", onboarding_pipeline.render),
    "offboarding": (":material/flight_takeoff: 离场流水线追踪", offboarding_pipeline.render),
    "report": (":material/contact_page: 花名册与报表导出", report_generator.render),
    "daily": (":material/chat_bubble: 微信日常办公辅助", daily_assistant.render),
    "attendance_payroll": (":material/payments: 考勤对账与工资结算", attendance_payroll.render),
}

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

selected_module_key = st.sidebar.radio(
    "功能导航",
    list(modules_map.keys()),
    format_func=lambda k: modules_map[k][0],
    label_visibility="collapsed"
)

# 6. 执行选中的模块
modules_map[selected_module_key][1]()
