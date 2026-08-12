# modules/pro_data_processing.py
import streamlit as st
from modules import merge_axis, purlins_summary

def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="header-emoji">📈</span>
        <div class="header-text">
            <h2>专业数据处理</h2>
            <p>本模块包含各类针对性强、专业性强的数据处理与计算工具</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)
    
    tools_map = {
        "檩条分轴合并与重量计算": merge_axis.render,
        "檩条下料单自动合计与汇总": purlins_summary.render,
    }
    
    # 顶部子导航，支持灵活扩展
    selected_tool = st.radio(
        "选择专业工具",
        list(tools_map.keys()),
        horizontal=True,
        label_visibility="collapsed"
    )
    
    st.divider()
    
    # 执行对应的工具渲染
    tools_map[selected_tool]()
