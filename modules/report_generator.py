import streamlit as st
import pandas as pd
from modules.master_data import get_last_changes, load_master_df

def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null", "<na>"]:
        return ""
    if s.endswith('.0'):
        s = s[:-2]
    return s

def format_contract_status(val):
    val = clean_val(val)
    if not val:
        return "否"
    # 如果包含签署完成、已签等字眼，转为是
    if any(keyword in val for keyword in ["完成", "已签", "是", "签订"]):
        return "是"
    return "否"

def ensure_excel_text(val):
    """
    通过在最前面添加单引号，确保粘贴到 Excel 时被强制识别为纯文本，
    从而防止身份证号等长数字变为科学计数法。
    """
    val = clean_val(val)
    if val:
        # 如果已经有单引号开头就不重复添加
        if not val.startswith("'"):
            return f"'{val}"
        return val
    return ""

def extract_roster(df):
    """花名册数据提取"""
    out_df = pd.DataFrame()
    out_df['姓名'] = df['姓名'] if '姓名' in df.columns else ""
    out_df['性别'] = df['性别'] if '性别' in df.columns else ""
    
    if '工种' in df.columns:
        out_df['工种(或岗位)'] = df['工种']
    elif '人员类型' in df.columns:
        out_df['工种(或岗位)'] = df['人员类型']
    else:
        out_df['工种(或岗位)'] = ""
        
    if '详细地址' in df.columns:
        out_df['家庭住址'] = df['详细地址']
    elif '家庭住址' in df.columns:
        out_df['家庭住址'] = df['家庭住址']
    else:
        out_df['家庭住址'] = ""
        
    if '身份证号' in df.columns:
        out_df['身份证号'] = df['身份证号'].apply(ensure_excel_text)
    else:
        out_df['身份证号'] = ""
        
    out_df['劳动合同编号'] = df['劳动合同编号'].apply(clean_val) if '劳动合同编号' in df.columns else ""
    
    return out_df

def extract_monthly_report(df, in_jianwei="是"):
    """变更月报数据提取"""
    out_df = pd.DataFrame()
    out_df['姓名'] = df['姓名'] if '姓名' in df.columns else ""
    
    if '身份证号' in df.columns:
        out_df['身份证号'] = df['身份证号'].apply(ensure_excel_text)
    else:
        out_df['身份证号'] = ""
    
    if '进场日期' in df.columns:
        out_df['进场时间'] = df['进场日期']
    elif '进场时间' in df.columns:
        out_df['进场时间'] = df['进场时间']
    else:
        out_df['进场时间'] = ""
        
    out_df['工种'] = df['工种'] if '工种' in df.columns else ""
    
    if '是否在市建委' in df.columns:
        out_df['是否在市建委'] = df['是否在市建委']
    else:
        out_df['是否在市建委'] = in_jianwei
        
    if '合同签订状态' in df.columns:
        out_df['是否签订《劳动合同》'] = df['合同签订状态'].apply(format_contract_status)
    elif '劳动合同' in df.columns:
        out_df['是否签订《劳动合同》'] = df['劳动合同'].apply(format_contract_status)
    else:
        out_df['是否签订《劳动合同》'] = "否"
        
    return out_df

def _paste_text(df: pd.DataFrame) -> str:
    """仅输出数据行，不输出表头；直接复制后粘贴到原有 Excel 空行。"""
    if df is None or df.empty:
        return ""
    return df.fillna("").astype(str).to_csv(sep="\t", index=False, header=False, lineterminator="\n")


def _render_copy_grid(title: str, df: pd.DataFrame, key: str) -> None:
    st.markdown(f"#### {title}")
    if df is None or df.empty:
        st.info("当前没有可追加的数据。")
        return
    st.dataframe(df, use_container_width=True, hide_index=True, height=min(520, 120 + len(df) * 36))
    st.markdown(
        '<div class="hint-box">这里只显示数据行，不含表头。点击下方代码框右上角的复制按钮，再粘贴到你原有 Excel 的最后一行；不会下载文件，也不会新增 Sheet。</div>',
        unsafe_allow_html=True,
    )
    st.code(_paste_text(df), language="text")


def _parse_date_series(df: pd.DataFrame, date_col: str) -> pd.Series:
    def parse_date(value):
        text = clean_val(value)
        if not text:
            return pd.NaT
        if text.isdigit() and len(text) == 5:
            try:
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(text), unit="D")
            except (ValueError, OverflowError):
                return pd.NaT
        return pd.to_datetime(text, errors="coerce")
    return df[date_col].apply(parse_date)


def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="header-emoji">📑</span>
        <div class="header-text">
            <h2>花名册与报表直贴</h2>
            <p>官网数据同步后，只复制新增人员行到你现有的 Excel 台账</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)

    df = load_master_df()
    if df.empty:
        st.warning("当前还没有项目人员主表。请先在【档案魔法整合】中上传智慧护薪和三局系统导出表并确认同步。")
        return

    changes = get_last_changes()
    new_df = pd.DataFrame(changes.get("new_rows") or [])
    updated_df = pd.DataFrame(changes.get("updated_rows") or [])
    st.markdown(
        f'<div class="celebrate-banner"><span>当前主表 {len(df)} 人 · 最近一次同步新增 {len(new_df)} 人 · 信息变化 {len(updated_df)} 人</span></div>',
        unsafe_allow_html=True,
    )

    tab_changes, tab_full, tab_quick = st.tabs(["🆕 本次同步直贴", "📦 主表全量提取", "⚡ 按姓名极速提取"])

    with tab_changes:
        if changes.get("updated_at"):
            st.caption(f"最近同步：{changes['updated_at']}；来源：{'、'.join(changes.get('source_files') or []) or '未记录'}")
        change_tab1, change_tab2 = st.tabs(["新增人员", "资料变化人员"])
        with change_tab1:
            roster_new = extract_roster(new_df) if not new_df.empty else pd.DataFrame()
            monthly_new = extract_monthly_report(new_df) if not new_df.empty else pd.DataFrame()
            sub1, sub2 = st.tabs(["花名册新增行", "变更月报新增行"])
            with sub1:
                _render_copy_grid("花名册新增行", roster_new, "copy_new_roster")
            with sub2:
                _render_copy_grid("变更月报新增行", monthly_new, "copy_new_monthly")
        with change_tab2:
            _render_copy_grid("资料变化明细（供核对）", updated_df, "copy_changed_people")

    with tab_full:
        working = df.copy()
        search = st.text_input("搜索姓名或身份证号", key="report_master_search")
        if search:
            mask = pd.Series(False, index=working.index)
            for col in ["姓名", "身份证号"]:
                if col in working.columns:
                    mask = mask | working[col].astype(str).str.contains(search, na=False)
            working = working[mask]

        date_col = "进场日期" if "进场日期" in working.columns else ("进场时间" if "进场时间" in working.columns else None)
        if date_col:
            dates = _parse_date_series(working, date_col)
            months = sorted(dates.dt.strftime("%Y-%m").dropna().unique(), reverse=True)
            selected_month = st.selectbox("按进场月份筛选", ["全部"] + months, key="report_master_month")
            if selected_month != "全部":
                working = working[dates.dt.strftime("%Y-%m") == selected_month]

        st.caption(f"当前筛选 {len(working)} 人；以下仍然只显示数据行，适合粘贴到既有 Excel。")
        full1, full2 = st.tabs(["花名册", "变更月报"])
        with full1:
            _render_copy_grid("花名册数据行", extract_roster(working), "copy_full_roster")
        with full2:
            in_jianwei = st.text_input("是否在市建委默认值", value="是", key="report_jianwei_default")
            _render_copy_grid("变更月报数据行", extract_monthly_report(working, in_jianwei), "copy_full_monthly")

    with tab_quick:
        names = st.text_input("输入需要提取的姓名，多个姓名用空格分隔", key="report_quick_names")
        if names.strip():
            name_list = [name.strip() for name in names.split() if name.strip()]
            selected = df[df.get("姓名", pd.Series(index=df.index, dtype=str)).isin(name_list)]
            if selected.empty:
                st.warning("未找到匹配人员，请检查姓名。")
            else:
                left, right = st.columns(2)
                with left:
                    _render_copy_grid("花名册数据行", extract_roster(selected), "copy_quick_roster")
                with right:
                    _render_copy_grid("变更月报数据行", extract_monthly_report(selected), "copy_quick_monthly")
