import streamlit as st
import pandas as pd
from modules.master_data import get_last_changes, load_master_df
from modules.archive_export import extract_cscec2_archive

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

FULL_INFO_COLUMNS = [
    "序号",
    "姓名",
    "性别",
    "民族",
    "年龄",
    "身份证号",
    "手机号",
    "详细地址",
    "班组",
    "工种",
    "人员类型",
    "进场日期",
    "在场/进退场状态",
    "工资卡号",
    "开户银行",
    "劳动合同编号",
    "合同签订状态",
    "项目简称",
    "分包/所属企业",
    "结算单价/标准",
    "进场告知书",
    "进场告知书编号",
    "是否参与统计",
    "申请原因",
    "项目全称",
    "退场日期",
    "出生日期",
    "籍贯",
    "政治面貌",
    "文化程度",
    "紧急联系人",
    "紧急联系电话",
    "是否班组长",
    "维权教育签订",
    "结算方式",
    "是否使用智慧护薪",
    "银行卡号",
    "协议签订",
]

COLUMN_SYNONYMS = {
    "姓名": ["姓名", "工人姓名"],
    "身份证号": ["身份证号", "身份证号码", "证件号码", "证件号"],
    "手机号": ["手机号", "手机号码", "联系电话", "电话", "联系方式", "手机"],
    "详细地址": ["详细地址", "家庭住址", "住址", "户籍地址"],
    "班组": ["班组", "所属班组", "班组名称", "施工班组"],
    "工种": ["工种", "工种名称", "岗位", "工种/岗位", "工种(或岗位)"],
    "人员类型": ["人员类型", "工人类型"],
    "进场日期": ["进场日期", "进场时间", "入场日期", "入场时间"],
    "在场/进退场状态": ["在场/进退场状态", "在场状态", "进退场状态", "在场情况"],
    "工资卡号": ["工资卡号", "银行卡号", "银行账号", "卡号"],
    "开户银行": ["开户银行", "银行", "开户行", "开户支行", "发卡行", "银行名称"],
    "劳动合同编号": ["劳动合同编号", "合同编号"],
    "合同签订状态": ["合同签订状态", "劳动合同", "合同签订", "合同状态"],
    "分包/所属企业": ["分包/所属企业", "所属企业", "分包单位", "劳务公司", "劳务分包单位", "所属公司", "单位名称"],
    "结算单价/标准": ["结算单价/标准", "工资标准", "结算单价(元)", "结算单价", "日薪", "日薪(元)", "日薪标准", "日工资", "基本工资", "单价", "工价"],
    "进场告知书": ["进场告知书", "进场告知书(纸质)"],
    "项目全称": ["项目全称", "工程名称"],
    "退场日期": ["退场日期", "退场时间", "离场日期", "离场时间"],
    "银行卡号": ["银行卡号", "工资卡号", "银行账号", "卡号"],
}


def extract_full_info(df: pd.DataFrame, ensure_text: bool = True) -> pd.DataFrame:
    """
    提取全量信息表数据行（严格对齐用户本地Excel的38列标准表头）
    - 序号列默认留空（便于复制粘贴到已有Excel末尾由公式或用户自增填充）
    - 严格保持38列标准顺序
    - 长数字文本字段进行防科学计数法保护
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=FULL_INFO_COLUMNS)

    out_df = pd.DataFrame(index=df.index)

    for col in FULL_INFO_COLUMNS:
        if col == "序号":
            out_df[col] = ""  # 用户明确要求序号留空
            continue

        # 1. 优先取原 df 中的同名列
        val_series = None
        if col in df.columns:
            val_series = df[col]
        else:
            # 2. 检查同义词回退
            synonyms = COLUMN_SYNONYMS.get(col, [])
            for syn in synonyms:
                if syn in df.columns:
                    val_series = df[syn]
                    break

        if val_series is not None:
            out_df[col] = val_series
        else:
            out_df[col] = ""

    # 从18位身份证号兜底推算年龄、出生日期、性别（若原值为空）
    if "身份证号" in out_df.columns:
        import datetime
        curr_year = datetime.datetime.now().year
        for idx in out_df.index:
            cid = clean_val(out_df.at[idx, "身份证号"]).replace("'", "")
            if len(cid) == 18 and cid[:17].isdigit():
                if not clean_val(out_df.at[idx, "出生日期"]):
                    out_df.at[idx, "出生日期"] = f"{cid[6:10]}-{cid[10:12]}-{cid[12:14]}"
                if not clean_val(out_df.at[idx, "性别"]):
                    out_df.at[idx, "性别"] = "男" if int(cid[16]) % 2 != 0 else "女"
                if not clean_val(out_df.at[idx, "年龄"]):
                    out_df.at[idx, "年龄"] = str(curr_year - int(cid[6:10]))

    # 防科学计数法与格式清理
    text_keywords = ['身份证', '手机', '电话', '卡号', '银行卡', '工资卡', '编号', '代码']
    for col in out_df.columns:
        if col == "序号":
            out_df[col] = ""
        elif ensure_text and any(kw in col for kw in text_keywords):
            out_df[col] = out_df[col].apply(ensure_excel_text)
        else:
            out_df[col] = out_df[col].apply(clean_val)

    return out_df[FULL_INFO_COLUMNS]

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

def extract_offboarded_report(df: pd.DataFrame, ensure_text: bool = True) -> pd.DataFrame:
    """提取退场人员27列标准报表数据行（长数字强制文本保护）"""
    if df is None or df.empty:
        return pd.DataFrame()
    from scripts.export_offboarded_workers import map_row, COLUMNS
    records = []
    for _, r in df.iterrows():
        row_dict = map_row(r.to_dict())
        records.append(row_dict)
    out_df = pd.DataFrame(records)[COLUMNS]
    if ensure_text:
        text_cols = ["身份证号码", "工资卡号", "劳动合同编号", "联系电话"]
        for c in text_cols:
            if c in out_df.columns:
                out_df[c] = out_df[c].apply(ensure_excel_text)
    return out_df

def paste_text(df: pd.DataFrame) -> str:
    """仅输出数据行，不输出表头；直接复制后粘贴到原有 Excel 空行。"""
    if df is None or df.empty:
        return ""
    return df.fillna("").astype(str).to_csv(sep="\t", index=False, header=False, lineterminator="\n")

_paste_text = paste_text


def render_copy_grid(title: str, df: pd.DataFrame, key: str = "") -> None:
    st.markdown(f"#### {title}")
    if df is None or df.empty:
        st.info("当前没有可追加的数据。")
        return
    st.dataframe(df, use_container_width=True, hide_index=True, height=min(520, 120 + len(df) * 36))
    st.markdown(
        '<div class="hint-box">这里只显示数据行，不含表头。点击下方代码框右上角的复制按钮，再粘贴到你原有 Excel 的最后一行；不会下载文件，也不会新增 Sheet。</div>',
        unsafe_allow_html=True,
    )
    st.code(paste_text(df), language="text")

_render_copy_grid = render_copy_grid


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
            <p>官网数据同步后，只复制新增人员行到你现有的 Excel 台账（支持全量信息表、中建二局标准表、花名册与变更月报）</p>
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

    tab_changes, tab_full, tab_quick, tab_offboarded = st.tabs([
        "🆕 本次同步直贴", "📦 主表全量提取", "⚡ 按姓名极速提取", "🚪 退场人员标准报表"
    ])

    with tab_changes:
        if changes.get("updated_at"):
            st.caption(f"最近同步：{changes['updated_at']}；来源：{'、'.join(changes.get('source_files') or []) or '未记录'}")
        change_tab1, change_tab2 = st.tabs(["新增人员 (新进场)", "资料变化人员"])
        with change_tab1:
            full_new = extract_full_info(new_df) if not new_df.empty else pd.DataFrame()
            cscec_new = extract_cscec2_archive(new_df) if not new_df.empty else pd.DataFrame()
            roster_new = extract_roster(new_df) if not new_df.empty else pd.DataFrame()
            monthly_new = extract_monthly_report(new_df) if not new_df.empty else pd.DataFrame()
            sub1, sub2, sub3, sub4 = st.tabs([
                "📊 全量信息表新增行",
                "📋 中建二局标准档案新增行",
                "📑 花名册新增行",
                "📅 变更月报新增行",
            ])
            with sub1:
                render_copy_grid("全量信息表新增行", full_new, "copy_new_full")
            with sub2:
                render_copy_grid("中建二局标准档案新增行", cscec_new, "copy_new_cscec")
            with sub3:
                render_copy_grid("花名册新增行", roster_new, "copy_new_roster")
            with sub4:
                render_copy_grid("变更月报新增行", monthly_new, "copy_new_monthly")
        with change_tab2:
            render_copy_grid("资料变化明细（供核对）", updated_df, "copy_changed_people")

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

        st.caption(f"当前筛选 {len(working)} 人；以下只显示数据行，适合直接复制粘贴到既有 Excel。")
        full1, full2, full3, full4 = st.tabs([
            "📊 全量信息表",
            "📋 中建二局标准档案",
            "📑 花名册",
            "📅 变更月报",
        ])
        with full1:
            render_copy_grid("全量信息表数据行", extract_full_info(working), "copy_full_info")
        with full2:
            render_copy_grid("中建二局标准档案数据行", extract_cscec2_archive(working), "copy_full_cscec")
        with full3:
            render_copy_grid("花名册数据行", extract_roster(working), "copy_full_roster")
        with full4:
            in_jianwei = st.text_input("是否在市建委默认值", value="是", key="report_jianwei_default")
            render_copy_grid("变更月报数据行", extract_monthly_report(working, in_jianwei), "copy_full_monthly")

    with tab_quick:
        names = st.text_input("输入需要提取的姓名，多个姓名用空格分隔", key="report_quick_names")
        if names.strip():
            name_list = [name.strip() for name in names.split() if name.strip()]
            selected = df[df.get("姓名", pd.Series(index=df.index, dtype=str)).isin(name_list)]
            if selected.empty:
                st.warning("未找到匹配人员，请检查姓名。")
            else:
                q1, q2, q3, q4 = st.tabs([
                    "📊 全量信息表",
                    "📋 中建二局标准档案",
                    "📑 花名册",
                    "📅 变更月报",
                ])
                with q1:
                    render_copy_grid("全量信息表数据行", extract_full_info(selected), "copy_quick_full")
                with q2:
                    render_copy_grid("中建二局标准档案数据行", extract_cscec2_archive(selected), "copy_quick_cscec")
                with q3:
                    render_copy_grid("花名册数据行", extract_roster(selected), "copy_quick_roster")
                with q4:
                    render_copy_grid("变更月报数据行", extract_monthly_report(selected), "copy_quick_monthly")

    with tab_offboarded:
        from pathlib import Path
        from scripts.export_offboarded_workers import build_excel_workbook

        st.markdown("### 🚪 退场人员标准报表（27列对齐）")
        st.markdown(
            "按照官方/项目退场要求（工人姓名、性别、出生日期、年龄、籍贯、民族、政治面貌、工种、家庭住址、身份证号码、发证机关、有效起始日期、有效截止日期、开户银行、开户支行、工资卡号、单位工资、劳动合同登记、劳动合同编号、安全培训情况、进场教育、首次进场时间、进场时间、退场时间、退场工资结清情况、持上岗证情况、联系电话）自动映射主表数据。"
        )

        off_27_df = df[df["在场/进退场状态"].astype(str).str.strip().isin(["已退场", "离场", "退场", "已离场"])].copy()
        off_37_df = df[df["退场日期"].astype(str).str.strip() != ""].copy()

        c1, c2, c3 = st.columns([1, 1, 1.2])
        with c1:
            st.metric("当前已退场人员", f"{len(off_27_df)} 人", help="主表中当前在场/进退场状态为『已退场』或『离场』的人员")
        with c2:
            st.metric("含历史退场记录", f"{len(off_37_df)} 人", help="主表中包含退场时间记录的所有人员（含退场后又二次进场的10人）")
        with c3:
            excel_path = Path("退场人员信息表.xlsx")
            if not excel_path.exists():
                build_excel_workbook(excel_path)
            with open(excel_path, "rb") as f:
                st.download_button(
                    label="📥 一键下载标准 Excel 表格",
                    data=f.read(),
                    file_name="退场人员信息表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="btn_download_offboarded_excel"
                )

        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        sub_tab1, sub_tab2 = st.tabs(["📋 当前退场人员 (27人)", "📑 含历史退场记录人员 (37人)"])
        with sub_tab1:
            render_copy_grid("当前退场人员27列数据行（可一键直贴既有Excel）", extract_offboarded_report(off_27_df), "copy_off_27")
        with sub_tab2:
            render_copy_grid("含历史退场记录人员27列数据行", extract_offboarded_report(off_37_df), "copy_off_37")

