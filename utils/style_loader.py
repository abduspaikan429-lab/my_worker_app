# utils/style_loader.py
import os
import streamlit as st

def load_css(css_path: str = None):
    """
    读取 CSS 文件并注入 Streamlit 页面。
    如果未指定 css_path，默认加载 assets/style.css
    """
    if css_path is None:
        # 相对于当前文件或项目根目录自动推导路径
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        css_path = os.path.join(base_dir, "assets", "style.css")

    if os.path.exists(css_path):
        with open(css_path, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>\n{css_content}\n</style>", unsafe_allow_html=True)
    else:
        st.warning(f"找不到样式文件: {css_path}")
