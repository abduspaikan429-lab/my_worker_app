import re

with open('modules/attendance_payroll.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Insert parse_watermark_attendance right after build_diff_table
insert_target = """    if company:
        result.insert(0, '分包/所属企业', company)
    return result"""

parse_func = """

def parse_watermark_attendance(df, company=''):
    \"\"\"
    单源水印签到表解析。
    基于 _safe_read 的结果，优先通过识别每日考勤计算出勤天数，
    如果未找到每日明细，则尝试读取总天数列。
    \"\"\"
    if df.empty:
        return __import__('pandas').DataFrame()
    
    days_col = _resolve_col(df, _DAYS_CANDS)
    records = _extract_daily_source(df, '水印签到表')
    
    rows = []
    if records:
        for key, item in records.items():
            name = item.get('姓名', '')
            identity = item.get('身份证号', '')
            days_dict = item.get('days', {})
            
            actual_days = 0
            has_daily_records = len(days_dict) > 0
            
            if has_daily_records:
                for day, val in days_dict.items():
                    state = _attendance_state(val)
                    if state == '有考勤':
                        actual_days += 1
            
            if not has_daily_records and days_col:
                name_col = _resolve_col(df, _NAME_CANDS)
                mask = (df[name_col] == name)
                if identity:
                    id_c = _resolve_col(df, _ID_CANDS)
                    if id_c:
                        mask = mask & (df[id_c].apply(_clean_identity) == identity)
                match = df[mask]
                if not match.empty:
                    val = _to_float(match.iloc[0][days_col])
                    actual_days = val if not __import__('pandas').isna(val) else 0
                    
            rows.append({
                '分包/所属企业': company,
                '姓名': name,
                '身份证号': identity,
                '解析出勤天数': float(actual_days),
                '最终核定天数': float(actual_days)
            })
    else:
        name_col = _resolve_col(df, _NAME_CANDS)
        id_col = _resolve_col(df, _ID_CANDS)
        for _, row in df.iterrows():
            name = _raw_attendance_value(row.get(name_col, ''))
            identity = _clean_identity(row.get(id_col, '')) if id_col else ''
            if not _looks_like_person_name(name) and not _is_valid_identity(identity):
                continue
            actual_days = _to_float(row.get(days_col)) if days_col else 0
            if __import__('pandas').isna(actual_days):
                actual_days = 0
            rows.append({
                '分包/所属企业': company,
                '姓名': name,
                '身份证号': identity,
                '解析出勤天数': float(actual_days),
                '最终核定天数': float(actual_days)
            })

    res_df = __import__('pandas').DataFrame(rows)
    if not res_df.empty:
        res_df['__key__'] = res_df.apply(lambda r: r['身份证号'] if _is_valid_identity(r['身份证号']) else f"姓名:{r['姓名']}", axis=1)
        res_df.drop_duplicates(subset='__key__', keep='first', inplace=True)
        res_df.drop(columns=['__key__'], inplace=True)
    return res_df"""

if parse_func not in content:
    content = content.replace(insert_target, insert_target + parse_func)

# 2. Replace Tab names
content = content.replace("':material/fact_check: 考勤对账与在线定稿',", "':material/fact_check: 考勤解析与定稿',")

# 3. Replace Tab 1 completely
# Use string slice from "    with tab_check:" to "    # ============================================================\n    # Tab 2"
start_idx = content.find("    with tab_check:")
end_idx = content.find("    # ============================================================\n    # Tab 2")

new_tab_check = """    with tab_check:
        att_status = st.session_state.get('_att_status', None) # None, 'draft', 'finalized'
        parsed_df = st.session_state.get('_att_parsed_df')
        final_df  = st.session_state.get('final_attendance')

        # ── Step 1：上传与解析 ────────────────────────────
        with st.container(border=True):
            st.markdown(\"\"\"
            <div class="step-indicator">
                <span class="step-num">1</span>
                <span>上传水印签到表并解析</span>
            </div>
            \"\"\", unsafe_allow_html=True)

            st.caption('请上传考勤水印签到表（系统将自动识别 √ 及时间格式计算出勤天数）')
            
            company_names = ['江苏旭之升建筑工程有限公司', '青海久昌建筑装饰工程有限公司']
            company = st.selectbox('选择分包/所属企业', company_names)
            file_watermark = st.file_uploader('水印签到表 (Excel)', type=['xlsx', 'xls'], key='att_watermark')

            if st.button(':material/document_scanner: 解析签到表', type='primary', key='btn_parse_watermark'):
                if not file_watermark:
                    st.warning(':material/info: 请上传水印签到表')
                else:
                    with st.spinner('正在解析签到表……'):
                        df_raw = _safe_read(file_watermark, '水印签到表')
                        res_df = parse_watermark_attendance(df_raw, company=company)
                        
                        if res_df.empty:
                            st.error('未能识别到有效人员或考勤数据，请检查表格格式。')
                        else:
                            st.session_state['_att_parsed_df'] = res_df
                            st.session_state['_att_status'] = 'draft'
                            st.session_state.pop('final_attendance', None)
                            st.session_state.pop('att_data_editor', None)
                            st.success(f'解析成功！共识别到 {len(res_df)} 名人员。')
                            st.rerun()

        # ── Step 2：在线确认与定稿 ─────────────────────────────────
        if att_status in ['draft', 'finalized'] and parsed_df is not None and not parsed_df.empty:
            st.markdown('---')
            with st.container(border=True):
                st.markdown(\"\"\"
                <div class="step-indicator">
                    <span class="step-num">2</span>
                    <span>核对考勤明细与定稿</span>
                </div>
                \"\"\", unsafe_allow_html=True)

                if att_status == 'draft':
                    st.markdown(
                        '<div class="hint-box">'
                        ':material/info: 当前为 <b>草稿 (draft)</b> 状态。<br>'
                        '请在下方双击 <b>✏ 最终核定天数</b> 逐行人工确认和微调，确认无误后点击“考勤定稿”。'
                        '</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        '<div class="celebrate-banner" style="margin:0;">'
                        '<span class="celebrate-icon">✅</span>'
                        '考勤已定稿 (finalized)。可进入下一步计算工资或重新修改后再次定稿。'
                        '</div>',
                        unsafe_allow_html=True,
                    )

                # ── 构建基础 DataFrame ─────────────────────────
                edit_base = parsed_df.copy()
                for extra in ['工种', '班组', '联系电话', '开户银行', '银行卡号', '性别', '手机号']:
                    if extra not in edit_base.columns:
                        edit_base[extra] = ''

                # 优先加载已保存的定稿数据
                def _row_key(row):
                    return f"{row.get('分包/所属企业', '')}|{row.get('身份证号', '')}|{row.get('姓名', '')}"

                if final_df is not None and '最终核定天数' in final_df.columns:
                    saved_map = {_row_key(r): r.get('最终核定天数') for _, r in final_df.iterrows()}
                    edit_base['最终核定天数'] = edit_base.apply(
                        lambda r: saved_map.get(_row_key(r), r.get('最终核定天数')), axis=1
                    )

                # ── 在线编辑器 ─────────────────────────────────
                display_edit_cols = [
                    '分包/所属企业', '姓名', '身份证号', '解析出勤天数', '最终核定天数',
                    '工种', '班组', '联系电话', '开户银行', '银行卡号',
                ]
                display_edit_cols = [c for c in display_edit_cols if c in edit_base.columns]

                col_cfg = {
                    '姓名':         st.column_config.TextColumn('姓名', disabled=True),
                    '分包/所属企业': st.column_config.TextColumn('公司', disabled=True),
                    '解析出勤天数': st.column_config.NumberColumn('系统解析天数', disabled=True, format='%.1f'),
                    '最终核定天数': st.column_config.NumberColumn(
                        '✏ 最终核定天数',
                        min_value=0, max_value=31, step=0.5, format='%.1f',
                        help='双击单行进行人工确认/修改',
                    ),
                    '工种':     st.column_config.TextColumn('工种'),
                    '班组':     st.column_config.TextColumn('班组'),
                    '身份证号': st.column_config.TextColumn('身份证号'),
                    '联系电话': st.column_config.TextColumn('联系电话'),
                    '开户银行': st.column_config.TextColumn('开户银行'),
                    '银行卡号': st.column_config.TextColumn('银行卡号'),
                }

                edited = st.data_editor(
                    edit_base[display_edit_cols],
                    column_config=col_cfg,
                    use_container_width=True,
                    num_rows='fixed',
                    hide_index=True,
                    key='att_data_editor',
                )

                c_save, c_info = st.columns([2, 5])
                with c_save:
                    if st.button(':material/check_circle: 考勤定稿', type='primary', key='btn_save_final', use_container_width=True):
                        st.session_state['final_attendance'] = enrich_with_master(edited)
                        st.session_state['_att_status'] = 'finalized'
                        st.success(':material/check_circle: 定稿成功 (finalized)！请切换至【考勤表与工资表导出】标签页进行工资计算。')
                        st.rerun()
                with c_info:
                    if att_status == 'finalized':
                        st.markdown('<div class="hint-box" style="margin:0;">:material/info: 已定稿。若重新修改并保存，将更新定稿数据。</div>', unsafe_allow_html=True)

        if att_status is not None:
            with st.expander(':material/refresh: 重新上传 (清空当前数据)'):
                st.warning('将清除当前的草稿和定稿数据，返回初始状态。')
                if st.button(':material/delete_forever: 确认清空', key='btn_reset_tab1', type='secondary'):
                    for k in ['_att_status', '_att_parsed_df', 'final_attendance']:
                        st.session_state.pop(k, None)
                    st.rerun()

"""
if start_idx != -1 and end_idx != -1:
    content = content[:start_idx] + new_tab_check + content[end_idx:]
else:
    print("Could not find Tab 1 boundaries")

# 4. Replace Tab 2 gating check
old_tab_export_gating = """    with tab_export:
        final_df = st.session_state.get('final_attendance')

        if final_df is None or final_df.empty:
            st.markdown(
                '<div class="alert-box alert-danger" style="margin-top:20px;">'
                ':material/warning: 尚未完成考勤定稿。请先在【考勤对账与在线定稿】标签页完成 Step B 并保存。'
                '</div>',
                unsafe_allow_html=True,
            )
            return"""

new_tab_export_gating = """    with tab_export:
        att_status = st.session_state.get('_att_status')
        final_df = st.session_state.get('final_attendance')

        if att_status != 'finalized' or final_df is None or final_df.empty:
            st.markdown(
                '<div class="alert-box alert-danger" style="margin-top:20px;">'
                ':material/warning: 必须先完成【考勤定稿】(状态: finalized) 才能进行工资计算和导出。'
                '</div>',
                unsafe_allow_html=True,
            )
            return"""

content = content.replace(old_tab_export_gating, new_tab_export_gating)

with open('modules/attendance_payroll.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement complete.")
