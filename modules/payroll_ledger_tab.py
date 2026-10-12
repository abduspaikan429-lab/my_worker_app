# -*- coding: utf-8 -*-
"""
modules/payroll_ledger_tab.py
工人工资发放动态总台账（多劳务分包商独立核算与全盘汇总）渲染模块
特点：
1. 视觉遵循高质感专业工程风
2. 支持按劳务公司（如江苏旭之升、青海久昌等）切换建账，账目清晰不串户
3. 支持在线文件上传与整段文本粘贴追加新批次，自动绑定分包公司
4. 支持选择任意本地文件夹导出并保存 Excel（支持 Windows 原生弹窗选择）
5. 支持按姓名、身份证、工种、在场退场状态智能检索
"""

import os
import streamlit as st
import pandas as pd
from datetime import datetime

from services.payroll_ledger_service import (
    load_all_payroll_data,
    save_uploaded_batch,
    list_all_batch_files,
    delete_batch_file,
    get_all_companies,
    export_ledger_to_folder,
    get_excel_bytes,
    open_folder_in_explorer,
    ask_directory_gui,
    PROJECT_ROOT,
    BATCHES_DIR
)




def render_payroll_ledger_tab():
    st.markdown("""
    <div style="margin-bottom: 16px;">
        <h3 style="margin-bottom: 4px; color: #1F4E79; font-weight: 700; display: flex; align-items: center; gap: 8px;">
            <span class="material-symbols-outlined" style="font-size: 26px; color: #2F5597;">receipt_long</span>
            工人工资发放动态总台账 · 多分包公司管理中心
        </h3>
        <p style="color: #64748B; font-size: 13px; margin: 0;">
            多劳务公司独立建账、全盘汇聚与离场销账追踪。各分包商批次互不干扰，支持终身追加与指定文件夹导出。
        </p>
    </div>
    """, unsafe_allow_html=True)

    # 1. 劳务分包商切换控制栏
    all_companies = get_all_companies()
    comp_options = ["全部劳务公司 (全盘汇总)"] + all_companies

    if "current_ledger_company" not in st.session_state:
        st.session_state.current_ledger_company = comp_options[1] if len(comp_options) > 1 else comp_options[0]

    col_comp_sel, col_comp_info = st.columns([2, 3], gap="medium")
    with col_comp_sel:
        chosen_comp = st.selectbox(
            "🏢 当前核算与查看的劳务分包商",
            comp_options,
            index=comp_options.index(st.session_state.current_ledger_company) if st.session_state.current_ledger_company in comp_options else 0,
            help="切换查看各劳务公司的独立发薪账目；导出时也将基于此范围生成专属台账",
            key="_select_ledger_comp"
        )
        st.session_state.current_ledger_company = chosen_comp

    is_all_scope = (chosen_comp == "全部劳务公司 (全盘汇总)")
    filter_arg = None if is_all_scope else chosen_comp

    # 2. 根据选中的劳务公司加载数据
    data = load_all_payroll_data(company_filter=filter_arg)
    stats = data["stats"]
    batch_shorts = data["batch_shorts"]
    summary_df = data["summary_df"]
    flow_df = data["flow_df"]
    sorted_workers = data["sorted_workers"]

    with col_comp_info:
        st.markdown(f"""
        <div style="background: #F1F5F9; border-radius: 8px; padding: 10px 14px; margin-top: 10px; font-size: 13px; color: #334155; display: flex; justify-content: space-between; align-items: center;">
            <div>
                <b>当前视角：</b><span style="color:#1D4ED8; font-weight:700;">{chosen_comp}</span>
                <span style="color:#64748B; margin-left: 8px;">(已归档 {stats['batch_count']} 个发薪批次)</span>
            </div>
            <div style="font-size: 12px; color: #64748B;">
                多公司互不串户 · 独立台账隔离
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 3. 顶部核心指标看板
    st.markdown(f"""
    <div class="metric-container" style="display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; margin-bottom: 20px; margin-top: 10px;">
        <div class="metric-card" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #64748B;">归档发薪批次</div>
            <div class="metric-value" style="font-size: 20px; font-weight: 700; color: #1E293B;">{stats['batch_count']} <span style="font-size: 12px; font-weight: normal;">批</span></div>
        </div>
        <div class="metric-card" style="background: #F0FDF4; border: 1px solid #DCFCE7; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #166534;">累计发薪总额</div>
            <div class="metric-value" style="font-size: 20px; font-weight: 700; color: #15803D;">¥{stats['total_pay']:,.2f}</div>
        </div>
        <div class="metric-card" style="background: #EFF6FF; border: 1px solid #DBEAFE; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #1E40AF;">涉及工人总数</div>
            <div class="metric-value" style="font-size: 20px; font-weight: 700; color: #2563EB;">{stats['worker_count']} <span style="font-size: 12px; font-weight: normal;">人 ({stats['flow_count']}人次)</span></div>
        </div>
        <div class="metric-card" style="background: #FAF5FF; border: 1px solid #F3E8FF; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #6B21A8;">累计出勤总工日</div>
            <div class="metric-value" style="font-size: 20px; font-weight: 700; color: #7E22CE;">{stats['total_days']:,.1f} <span style="font-size: 12px; font-weight: normal;">天</span></div>
        </div>
        <div class="metric-card" style="background: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #065F46;">当前在场人员发薪</div>
            <div class="metric-value" style="font-size: 18px; font-weight: 700; color: #047857;">{stats['active_count']} 人 <span style="font-size: 12px;">(¥{stats['active_pay']:,.0f})</span></div>
        </div>
        <div class="metric-card" style="background: #FFF1F2; border: 1px solid #FECDD3; border-radius: 8px; padding: 12px;">
            <div class="metric-title" style="font-size: 12px; color: #9F1239;">已退场结清人员</div>
            <div class="metric-value" style="font-size: 18px; font-weight: 700; color: #BE123C;">{stats['left_count']} 人 <span style="font-size: 12px;">(¥{stats['left_pay']:,.0f})</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 4. 核心操作区两栏布局：左侧【追加新批次】，右侧【选择目录导出】
    col_upload, col_export = st.columns([1, 1], gap="medium")

    # ─────────────── 左侧：追加新批次 ───────────────
    with col_upload:
        with st.container(border=True):
            st.markdown("""
            <div style="font-weight: 600; font-size: 14px; margin-bottom: 8px; color: #1E293B; display: flex; align-items: center; gap: 6px;">
                <span class="material-symbols-outlined" style="font-size: 18px; color: #0284C7;">upload_file</span>
                📥 追加新批次工资表
            </div>
            """, unsafe_allow_html=True)

            # 选择归属分包商与录入方式
            default_assign = chosen_comp if not is_all_scope else all_companies[0]
            col_target_c1, col_target_c2 = st.columns([1.6, 2])
            with col_target_c1:
                target_assign_comp = st.selectbox(
                    "归属劳务分包公司",
                    all_companies,
                    index=all_companies.index(default_assign) if default_assign in all_companies else 0,
                    help="系统将把此批次归入该劳务公司的台账账目中（若文本自带分包商名称则以文本为准）"
                )
            with col_target_c2:
                append_mode = st.radio("录入方式", ["上传官方 Excel (.xlsx/.xls)", "粘贴工资表文本", "上传 CSV/TXT"], horizontal=True)

            uploaded_items = []
            if append_mode == "上传官方 Excel (.xlsx/.xls)":
                uploaded_files = st.file_uploader(
                    "选择官方导出的 Excel 工资表 (支持批量多选)",
                    type=["xlsx", "xls"],
                    accept_multiple_files=True,
                    help="支持一次性选择或拖入多个 .xlsx / .xls 文件（例如同时上传6月、7月、8月各批次）"
                )
                if uploaded_files:
                    for uf in uploaded_files:
                        uploaded_items.append((uf.getvalue(), uf.name))
                    st.caption(f"📁 已选取 **{len(uploaded_files)}** 个 Excel 文件待批量解析追加")
            elif append_mode == "粘贴工资表文本":
                uploaded_text = st.text_area(
                    "粘贴工资表文本",
                    placeholder="在此直接粘贴官方工资表内容（含工资单编号、表头及人员数据行）...\n系统会自动识别分包商名称并归档！",
                    height=100,
                    label_visibility="collapsed"
                )
                if uploaded_text and uploaded_text.strip():
                    uploaded_items.append((uploaded_text.strip(), "pasted_text.csv"))
            else:
                uploaded_files = st.file_uploader(
                    "选择 CSV/TXT 文件 (支持批量多选)",
                    type=["csv", "txt"],
                    accept_multiple_files=True,
                    label_visibility="collapsed"
                )
                if uploaded_files:
                    for uf in uploaded_files:
                        uploaded_items.append((uf.getvalue(), uf.name))
                    st.caption(f"📁 已选取 **{len(uploaded_files)}** 个文件待批量解析")

            c_btn_app, c_btn_view = st.columns([1.5, 1])
            with c_btn_app:
                btn_label = f":material/add_circle: 批量解析并追加 ({len(uploaded_items)}个)" if len(uploaded_items) > 1 else ":material/add_circle: 解析并追加到台账"
                if st.button(btn_label, type="primary", use_container_width=True):
                    if not uploaded_items:
                        st.warning("⚠️ 请先选取一个或多个 Excel 文件或粘贴文本内容！")
                    else:
                        total_batches_saved = 0
                        all_messages = []
                        for content, fname in uploaded_items:
                            cnt, msgs = save_uploaded_batch(
                                content,
                                assigned_company=target_assign_comp,
                                filename=fname
                            )
                            total_batches_saved += cnt
                            all_messages.extend(msgs)
                        
                        if total_batches_saved > 0:
                            st.success(f"🎉 批量导入完成！共处理 {len(uploaded_items)} 个文件，成功归档 {total_batches_saved} 个发薪批次。")
                            for m in all_messages:
                                st.caption(f"· {m}")
                            st.session_state.current_ledger_company = target_assign_comp
                            st.rerun()
                        else:
                            for m in all_messages:
                                st.error(f"❌ {m}")
            
            with c_btn_view:
                with st.popover(":material/history: 已归档批次管理", use_container_width=True):
                    st.caption("查看或管理已存入底库的批次文件：")
                    b_files = list_all_batch_files()
                    if not b_files:
                        st.info("暂无归档批次")
                    else:
                        for bf in b_files:
                            c_f1, c_f2 = st.columns([3, 1])
                            c_f1.text(f"{bf['filename']} ({bf['size']})")
                            if c_f2.button("删除", key=f"del_{bf['filename']}", type="secondary"):
                                delete_batch_file(bf["filename"])
                                st.success(f"已删除 {bf['filename']}")
                                st.rerun()

    # ─────────────── 右侧：选择目录导出 ───────────────
    with col_export:
        with st.container(border=True):
            st.markdown("""
            <div style="font-weight: 600; font-size: 14px; margin-bottom: 8px; color: #1E293B; display: flex; align-items: center; gap: 6px;">
                <span class="material-symbols-outlined" style="font-size: 18px; color: #059669;">folder_open</span>
                💾 自定义文件夹导出 Excel 台账
            </div>
            """, unsafe_allow_html=True)

            if "export_folder_path" not in st.session_state:
                st.session_state.export_folder_path = PROJECT_ROOT

            col_path, col_browse = st.columns([3, 1.2])
            with col_path:
                target_folder = st.text_input(
                    "保存目标文件夹路径",
                    value=st.session_state.export_folder_path,
                    placeholder="输入或选择要保存的目标文件夹路径...",
                    label_visibility="collapsed"
                )
                st.session_state.export_folder_path = target_folder
            
            with col_browse:
                if st.button(":material/folder_open: 浏览选择...", use_container_width=True, help="点击唤起 Windows 文件夹选择框"):
                    chosen = ask_directory_gui(st.session_state.export_folder_path)
                    if chosen:
                        st.session_state.export_folder_path = chosen
                        st.rerun()

            # 快捷路径小药丸
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                if st.button("🖥️ 桌面", key="set_desktop", use_container_width=True):
                    st.session_state.export_folder_path = desktop_path
                    st.rerun()
            with col_p2:
                if st.button("📁 项目根目录", key="set_root", use_container_width=True):
                    st.session_state.export_folder_path = PROJECT_ROOT
                    st.rerun()
            with col_p3:
                data_dir = os.path.join(PROJECT_ROOT, "data")
                if st.button("📂 data 目录", key="set_data", use_container_width=True):
                    st.session_state.export_folder_path = data_dir
                    st.rerun()

            # 智能建议文件名（随选中的劳务公司动态调整）
            comp_file_tag = "全部分包大盘" if is_all_scope else chosen_comp
            suggested_filename = f"科技文化中心项目_{comp_file_tag}_工人工资发放动态台账.xlsx"

            export_filename = st.text_input(
                "导出文件名",
                value=suggested_filename,
                help="默认导出为带 SUMIFS 活表动态公式联动的高质感 Excel"
            )

            col_btn_save, col_btn_down = st.columns([1.5, 1])
            with col_btn_save:
                if st.button(f":material/save: 保存【{chosen_comp[:7]}】台账", type="primary", use_container_width=True):
                    ok, full_saved_path, err = export_ledger_to_folder(
                        st.session_state.export_folder_path,
                        filename=export_filename,
                        company_scope=filter_arg
                    )
                    if ok:
                        st.session_state.last_saved_path = full_saved_path
                        st.session_state.last_saved_folder = st.session_state.export_folder_path
                        st.success(f"🎉 成功导出台账至：\n`{full_saved_path}`")
                    else:
                        st.error(f"❌ 导出失败：{err}")

            with col_btn_down:
                excel_bytes = get_excel_bytes(company_scope=filter_arg)
                st.download_button(
                    label=":material/download: 浏览器下载",
                    data=excel_bytes,
                    file_name=export_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # 打开所在文件夹
            if st.session_state.get("last_saved_folder"):
                if st.button(f":material/open_in_new: 打开所在文件夹 ({os.path.basename(st.session_state.last_saved_folder)})", key="btn_open_exp", use_container_width=True):
                    open_folder_in_explorer(st.session_state.last_saved_folder)



    # 5. 动态总台账数据表格与多视角检索展示
    st.markdown("---")
    st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
        <div style="font-weight: 700; font-size: 15px; color: #1F4E79; display: flex; align-items: center; gap: 6px;">
            <span class="material-symbols-outlined" style="font-size: 20px;">table_view</span>
            {chosen_comp} · 动态工资台账与明细检索
        </div>
        <div style="color: #64748B; font-size: 12px;">
            💡 提示：本表包含动态公式映射，可直接点击列头排序或使用筛选器
        </div>
    </div>
    """, unsafe_allow_html=True)

    if summary_df.empty:
        st.info(f"✨ 当前【{chosen_comp}】暂无归档发薪记录，请在左侧上传或粘贴该公司的官方工资表！")
        return

    # 检索与过滤条
    col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
    with col_s1:
        search_kw = st.text_input("🔍 搜索人员", placeholder="输入工人姓名、身份证号或手机号模糊搜索...", label_visibility="collapsed")
    with col_s2:
        all_jobs = ["全部工种"] + sorted(list(set(w["job_type"] for w in sorted_workers)))
        sel_job = st.selectbox("工种筛选", all_jobs, label_visibility="collapsed")
    with col_s3:
        sel_status = st.selectbox("状态筛选", ["全部状态", "已进场", "已退场"], label_visibility="collapsed")

    filtered_df = summary_df.copy()
    if search_kw:
        kw = search_kw.strip()
        filtered_df = filtered_df[
            filtered_df["姓名"].str.contains(kw, na=False) |
            filtered_df["身份证号"].str.contains(kw, na=False) |
            filtered_df["联系电话"].str.contains(kw, na=False)
        ]
    if sel_job != "全部工种":
        filtered_df = filtered_df[filtered_df["工种"] == sel_job]
    if sel_status != "全部状态":
        filtered_df = filtered_df[filtered_df["状态"] == sel_status]

    # 三视角 Tab 切换
    tab_view_all, tab_view_off, tab_view_flow = st.tabs([
        f":material/badge: 工资总台账 ({len(filtered_df)}人)",
        f":material/person_off: 已退场工人结算专栏 ({stats['left_count']}人)",
        f":material/list_alt: 发放流水明细底库 ({stats['flow_count']}条)"
    ])

    with tab_view_all:
        # 如果是单公司视图，隐藏重复的“所属劳务公司”列
        display_df = filtered_df.copy()
        if not is_all_scope and "所属劳务公司" in display_df.columns:
            display_df = display_df.drop(columns=["所属劳务公司"])

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "累计实发总额": st.column_config.NumberColumn(
                    "累计实发总额(元)",
                    format="¥%.2f"
                ),
                "累计出勤(天)": st.column_config.NumberColumn(
                    "累计出勤(天)",
                    format="%.1f 天"
                ),
            }
        )

    with tab_view_off:
        off_df = summary_df[summary_df["状态"] == "已退场"].copy()
        if off_df.empty:
            st.info(f"当前【{chosen_comp}】暂无已退场人员记录。")
        else:
            st.caption(f"当前【{chosen_comp}】共 {len(off_df)} 名工人已办理退场，累计结清工资总计：**¥{stats['left_pay']:,.2f}**")
            cols_to_show = ["序号", "姓名", "所属劳务公司", "身份证号", "联系电话", "工种", "状态", "日薪标准", "累计出勤(天)", "累计实发总额", "实名制银行卡号", "开户银行及支行"]
            if not is_all_scope and "所属劳务公司" in cols_to_show:
                cols_to_show.remove("所属劳务公司")
            cols_to_show = [c for c in cols_to_show if c in off_df.columns]
            st.dataframe(
                off_df[cols_to_show],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "累计实发总额": st.column_config.NumberColumn("已结清总额(元)", format="¥%.2f"),
                    "累计出勤(天)": st.column_config.NumberColumn("累计出勤(天)", format="%.1f 天"),
                }
            )

    with tab_view_flow:
        st.caption(f"集中展示发薪原始流水记录（共 {len(flow_df)} 条流水记录）：")
        display_flow = flow_df.copy()
        if not is_all_scope and "所属劳务公司" in display_flow.columns:
            display_flow = display_flow.drop(columns=["所属劳务公司"])
        st.dataframe(
            display_flow,
            use_container_width=True,
            hide_index=True,
            column_config={
                "实发总额(元)": st.column_config.NumberColumn("实发总额(元)", format="¥%.2f"),
                "实际出勤(天)": st.column_config.NumberColumn("实际出勤(天)", format="%.1f 天"),
                "实发基本工资": st.column_config.NumberColumn("实发基本工资", format="¥%.2f"),
                "津贴": st.column_config.NumberColumn("津贴", format="¥%.2f"),
                "扣减支出": st.column_config.NumberColumn("扣减支出", format="¥%.2f"),
            }
        )
