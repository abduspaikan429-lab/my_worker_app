# modules/pro_data_processing.py
import streamlit as st
from modules import merge_axis, purlins_summary

def render_custom_tool():
    st.markdown("""
    <div class="custom-card" style="text-align: center; padding: 40px 20px;">
        <span class="material-symbols-outlined" style="font-size: 48px; color: #818CF8; margin-bottom: 12px;">extension</span>
        <h3 style="color: #334155; margin-bottom: 8px;">③ 预留扩展专业工具</h3>
        <p style="color: #64748B; font-size: 14px; max-width: 500px; margin: 0 auto 16px auto;">
            此区域已预留，可根据工程与技术部门后续需求，随时接入钢结构图纸解析、工程量核算或下料规格优化等新工具。
        </p>
        <span class="tag-badge badge-purple">✨ 持续扩展中</span>
    </div>
    """, unsafe_allow_html=True)

def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined header-emoji" style="color:#6366F1;">analytics</span>
        <div class="header-text">
            <h2>专业数据处理</h2>
            <p>低频专业计算工具集 · 檩条分轴合并、下料单自动汇总与工程计算</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)
    
    tools_map = {
        "① 檩条分轴合并与重量计算": merge_axis.render,
        "② 檩条下料单自动合计与汇总": purlins_summary.render,
        "③ 扩展专业工具 (预留)": render_custom_tool,
    }
    
    tool_keys = list(tools_map.keys())
    
    if "pro_data_tool_index" not in st.session_state or st.session_state.pro_data_tool_index >= len(tool_keys):
        st.session_state.pro_data_tool_index = 0
        
    def on_tool_change():
        if "_pro_data_radio" in st.session_state and st.session_state._pro_data_radio in tool_keys:
            st.session_state.pro_data_tool_index = tool_keys.index(st.session_state._pro_data_radio)
    
    # 顶部子导航
    selected_tool = st.radio(
        "选择专业工具",
        tool_keys,
        index=st.session_state.pro_data_tool_index,
        horizontal=True,
        label_visibility="collapsed",
        key="_pro_data_radio",
        on_change=on_tool_change
    )
    
    st.divider()
    
    # 执行对应的工具渲染
    tools_map[selected_tool]()
