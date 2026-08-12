# modules/merge_axis.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.utils import get_column_letter

# ==================== 配置区 ====================
SPEC_WEIGHT = {                   # 规格 → 米重（kg/m）
    "220*100*3": 14.79,
    "220*100*4": 19.59,
    "220*100*5": 24.34,
}
SPEC_ORDER = ["220*100*3", "220*100*4", "220*100*5"]   # 合计中规格的显示顺序
# ================================================

def process_merge_axis(file_input):
    """
    核心业务逻辑：完全保持原 merge_axis.py 处理逻辑，不对表格计算逻辑做任何修改
    """
    xls = pd.ExcelFile(file_input)
    all_sheets = xls.sheet_names

    pattern = re.compile(r'^(\d+)轴\s*[-]\s*([a-dA-D])$')
    axis_groups = {}
    for name in all_sheets:
        m = pattern.match(name)
        if m:
            axis_num = m.group(1)
            suffix = m.group(2).lower()
            axis_groups.setdefault(axis_num, {})[suffix] = name

    if not axis_groups:
        return None, None, "未找到任何符合格式（如 '1轴-a'）的 sheet，请检查源文件。"

    sorted_axes = sorted(axis_groups.keys(), key=int)
    sheets_to_write = {}

    for axis in sorted_axes:
        group = axis_groups[axis]
        dfs = []
        used = []
        for letter in ['a', 'b', 'c', 'd']:
            if letter in group:
                dfs.append(xls.parse(group[letter]))
                used.append(group[letter])
        if not dfs:
            continue

        combined = pd.concat(dfs, ignore_index=True)

        # 列名智能映射
        rename = {}
        for col in combined.columns:
            col_str = str(col)
            if '实量长度' in col_str:
                rename[col] = '实量长度/根（mm）'
            elif '规格' in col_str:
                rename[col] = '规格（mm）'
            elif '数量' in col_str or '根数' in col_str:
                rename[col] = '数量（根）'
            elif '序号' in col_str:
                rename[col] = '序号'
            elif '编号' in col_str:
                rename[col] = '编号'
        combined.rename(columns=rename, inplace=True)

        required = ['编号', '规格（mm）', '实量长度/根（mm）', '数量（根）']
        if missing := [r for r in required if r not in combined.columns]:
            continue

        # ---------- 计算轴线（修复版）----------
        extracted = combined["编号"].astype(str).str.extract(r'R(\d+)', expand=False)
        combined["轴线"] = "R" + extracted + "轴"
        combined.loc[extracted.isna(), "轴线"] = ""

        # 米重和重量
        combined["米重"] = combined["规格（mm）"].map(SPEC_WEIGHT)
        combined["重量（kg）"] = round(
            pd.to_numeric(combined["实量长度/根（mm）"], errors='coerce') * combined["米重"] / 1000, 5
        )

        # ---------- 合计行 ----------
        total_qty = combined["数量（根）"].sum()
        total_weight = combined["重量（kg）"].sum()
        spec_qty = {sp: combined.loc[combined["规格（mm）"] == sp, "数量（根）"].sum() for sp in SPEC_ORDER}
        spec_wt = {sp: round(combined.loc[combined["规格（mm）"] == sp, "重量（kg）"].sum(), 5) for sp in SPEC_ORDER}

        new_cols = []
        for sp in SPEC_ORDER:
            new_cols.append(f"{sp}数量(根)")
            new_cols.append(f"{sp}重量(kg)")
        for col in new_cols:
            combined[col] = ""

        final_cols = ["序号", "编号", "规格（mm）", "实量长度/根（mm）",
                      "数量（根）", "重量（kg）", "轴线", "米重"] + new_cols
        combined = combined[final_cols]

        summary = {col: "" for col in final_cols}
        summary["序号"] = "合计"
        summary["数量（根）"] = total_qty
        summary["重量（kg）"] = round(total_weight, 5)
        for sp in SPEC_ORDER:
            summary[f"{sp}数量(根)"] = spec_qty[sp]
            summary[f"{sp}重量(kg)"] = spec_wt[sp]

        summary_df = pd.DataFrame([summary])
        combined = pd.concat([combined, summary_df], ignore_index=True)

        sheets_to_write[f"{axis}轴"] = combined

    if not sheets_to_write:
        return None, None, "没有生成任何轴 sheet，请检查数据格式。"

    output_stream = BytesIO()
    with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
        for sheet_name, df in sheets_to_write.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            num_rows, num_cols = df.shape

            # 1. 冻结首行
            ws.freeze_panes = "A2"

            # 2. 筛选区域（不含合计行）
            if num_rows > 2:
                last_data_row = num_rows - 1
                last_col_letter = get_column_letter(num_cols)
                ws.auto_filter.ref = f"A1:{last_col_letter}{last_data_row}"

            # 3. 设置为分页预览视图
            ws.sheet_view.view = 'pageBreakPreview'

    return output_stream.getvalue(), sheets_to_write, None

def render():
    """UI 渲染入口函数（复用 assets/style.css 全局样式）"""
    st.header(":material/view_timeline: 檩条分轴数据合并与重量计算")
    st.markdown("<p style='color: #64748B; margin-bottom: 20px;'>自动识别源 Excel 中的各轴 Sheet（如 1轴-a, 1轴-b），合并计算轴线、米重与总重量，并生成标准 Excel。</p>", unsafe_allow_html=True)

    st.markdown("""
        <div class="step-indicator">
            <span>Step 1. 上传下料单</span>
            <span>Step 2. 自动分组计算</span>
            <span>Step 3. 预览与导出结果</span>
        </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.subheader("数据源上传")
        st.info("提示：Excel 中需包含以『数字轴-字母』（如 1轴-a）命名的 Sheet。")
        uploaded_file = st.file_uploader("上传檩条下料单 Excel 文件", type=["xlsx", "xls"], key="merge_axis_uploader")

    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
    with col_btn2:
        process_btn = st.button("开始自动计算并合并", type="primary", use_container_width=True)

    if 'merge_axis_run' not in st.session_state:
        st.session_state.merge_axis_run = False
        
    if process_btn:
        st.session_state.merge_axis_run = True
        
    if not uploaded_file:
        st.session_state.merge_axis_run = False
        
    if not st.session_state.merge_axis_run:
        return

    try:
        with st.spinner("正在自动分组合并并计算数据..."):
            excel_bytes, sheets_dict, err = process_merge_axis(uploaded_file)

        if err:
            st.warning(f"{err}")
            return

        # 统计汇总数据
        total_axes = len(sheets_dict)
        total_qty = sum(float(df.loc[df["序号"] == "合计", "数量（根）"].values[0]) for df in sheets_dict.values() if "合计" in df["序号"].values)
        total_weight = sum(float(df.loc[df["序号"] == "合计", "重量（kg）"].values[0]) for df in sheets_dict.values() if "合计" in df["序号"].values)

        with st.container(border=True):
            st.subheader("计算结果概览")
            st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-title">识别轴数</div>
                        <div class="metric-value">{total_axes} <span style="font-size:0.9rem; font-weight:normal; color:#64748B;">个轴</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">总根数</div>
                        <div class="metric-value">{int(total_qty):,} <span style="font-size:0.9rem; font-weight:normal; color:#64748B;">根</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">总重量 (kg)</div>
                        <div class="metric-value">{total_weight:,.2f} <span style="font-size:0.9rem; font-weight:normal; color:#64748B;">kg</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">折合吨数 (t)</div>
                        <div class="metric-value" style="color: #1677FF;">{total_weight/1000:,.3f} <span style="font-size:0.9rem; font-weight:normal; color:#64748B;">t</span></div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.divider()

            tab_preview, tab_export = st.tabs(["各轴数据预览", "数据导出"])

        with tab_preview:
            selected_axis = st.selectbox("选择轴号查看明细", list(sheets_dict.keys()))
            if selected_axis:
                st.dataframe(sheets_dict[selected_axis], use_container_width=True, height=600)

        with tab_export:
            st.markdown("### 导出合并计算后的 Excel 文件")
            st.write("导出的 Excel 文件每个轴 sheet 均已冻结首行、添加自动筛选（合计行除外），并设置为分页预览视图。")

            st.download_button(
                label="下载合并结果 Excel",
                data=excel_bytes,
                file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_合并计算.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=False
            )

    except Exception as e:
        st.error(f"处理数据时出错: {e}")
