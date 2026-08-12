import streamlit as st
import datetime
import os
import json
import uuid

MEMO_FILE = "data/memos.json"

def load_memos():
    if not os.path.exists("data"):
        os.makedirs("data")
    if os.path.exists(MEMO_FILE):
        try:
            with open(MEMO_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_memos(memos):
    if not os.path.exists("data"):
        os.makedirs("data")
    with open(MEMO_FILE, "w", encoding="utf-8") as f:
        json.dump(memos, f, ensure_ascii=False, indent=2)

def render():
    st.markdown("""
    <div class="page-header-deco">
        <span class="header-emoji">💬</span>
        <div class="header-text">
            <h2>微信日常办公辅助</h2>
            <p>快速生成微信汇报文本，记录闪念备忘，让日常办公更轻松</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["考勤日报生成器", "闪念备忘与疑问库"])

    with tab1:
        # 1. 班组数据初始化（默认汪佩沾、王宜强两个班组）
        if 'attendance_teams_v2' not in st.session_state:
            st.session_state.attendance_teams_v2 = [
                {"id": str(uuid.uuid4()), "name": "汪佩沾", "present": 22, "absent": 7, "leave": 1},
                {"id": str(uuid.uuid4()), "name": "王宜强", "present": 12, "absent": 13, "leave": 0}
            ]

        # 采用左右分栏布局：左侧录入，右侧实时预览
        left_col, right_col = st.columns([7, 3], gap="medium")

        with left_col:
            st.markdown("#### 班组考勤录入")
            date_val = st.date_input("选择日期", datetime.date.today())
            formatted_date = f"{date_val.year}年{date_val.month}月{date_val.day}号"
            
            # 使用表头代替每个输入框的 label，极致压缩垂直空间
            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
            hc1, hc2, hc3, hc4, hc5 = st.columns([2, 2.5, 2.5, 2.5, 1])
            with hc1: st.caption("班组名称")
            with hc2: st.caption("已出勤")
            with hc3: st.caption("未出勤")
            with hc4: st.caption("请假")
            with hc5: st.caption("操作")
            
            for idx, team in enumerate(st.session_state.attendance_teams_v2):
                c1, c2, c3, c4, c5 = st.columns([2, 2.5, 2.5, 2.5, 1], vertical_alignment="center")
                with c1:
                    team['name'] = st.text_input("班组名称", value=team['name'], key=f"name_{team['id']}", label_visibility="collapsed")
                with c2:
                    team['present'] = st.number_input("已出勤", min_value=0, step=1, value=team['present'], key=f"p_{team['id']}", label_visibility="collapsed")
                with c3:
                    team['absent'] = st.number_input("未出勤", min_value=0, step=1, value=team['absent'], key=f"a_{team['id']}", label_visibility="collapsed")
                with c4:
                    team['leave'] = st.number_input("请假", min_value=0, step=1, value=team['leave'], key=f"l_{team['id']}", label_visibility="collapsed")
                with c5:
                    if len(st.session_state.attendance_teams_v2) > 1:
                        if st.button("删除", key=f"del_team_{team['id']}", help="删除"):
                            st.session_state.attendance_teams_v2.pop(idx)
                            st.rerun()

            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
            if st.button("添加更多班组", use_container_width=True):
                st.session_state.attendance_teams_v2.append({"id": str(uuid.uuid4()), "name": "新施工班组", "present": 0, "absent": 0, "leave": 0})
                st.rerun()

        with right_col:
            st.markdown("<div class='card-container' style='height: 100%;'>", unsafe_allow_html=True)
            st.markdown("### 实时预览与复制")
            
            # 实时生成报表文本
            report_lines = []
            for team in st.session_state.attendance_teams_v2:
                total = team['present'] + team['absent'] + team['leave']
                
                t_str = (
                    f"考勤日报：{formatted_date}\n"
                    f"施工班组：{team['name']}；共{total}人\n"
                    f"已出勤：{team['present']}人\n"
                    f"未出勤：{team['absent']}人\n"
                    f"请假：{team['leave']}人"
                )
                report_lines.append(t_str)
            
            report_text = "\n\n".join(report_lines)

            st.markdown("<span style='color: var(--text-secondary); font-size: 0.85rem;'>点击代码框右上角的复制图标即可复制。</span>", unsafe_allow_html=True)
            st.code(report_text, language='text')
            st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        if 'memos' not in st.session_state:
            st.session_state.memos = load_memos()

        def add_memo(text, cat):
            st.session_state.memos.append({
                "id": str(datetime.datetime.now().timestamp()), 
                "text": text, 
                "category": cat
            })
            save_memos(st.session_state.memos)

        def remove_memo(memo_id):
            st.session_state.memos = [m for m in st.session_state.memos if m["id"] != memo_id]
            save_memos(st.session_state.memos)

        st.markdown("<div class='module-card module-daily'>", unsafe_allow_html=True)
        st.markdown("#### 添加新备忘")
        
        c1, c2, c3 = st.columns([5, 3, 2])
        with c1:
            new_memo = st.text_input("备忘内容", key="new_memo_input", label_visibility="collapsed", placeholder="输入备忘事项...")
        with c2:
            category = st.selectbox("分类", ["领导临时交办", "业务红线", "待询疑问"], key="new_memo_cat", label_visibility="collapsed")
        with c3:
            if st.button("快速添加", use_container_width=True, type="primary"):
                if new_memo.strip():
                    add_memo(new_memo.strip(), category)
                    st.rerun()
                else:
                    st.warning("请输入备忘内容！")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("#### 任务看板")
        if not st.session_state.memos:
            st.markdown('<div class="celebrate-banner"><span>太棒了，当前没有任何待办备忘！</span></div>', unsafe_allow_html=True)
        else:
            # 采用 3 列看板布局，直观且节省空间
            col_leader, col_redline, col_question = st.columns(3)
            
            leader_memos = [m for m in st.session_state.memos if "领导" in m['category']]
            redline_memos = [m for m in st.session_state.memos if "红线" in m['category']]
            question_memos = [m for m in st.session_state.memos if "待询" in m['category']]

            def render_memo_column(memos, title, color_class):
                # 标题和计数
                st.markdown(f"**{title}** <span class='tag-badge {color_class}' style='margin-left: 8px;'>{len(memos)}</span>", unsafe_allow_html=True)
                if not memos:
                    st.markdown("<span style='color:var(--text-secondary);font-size:0.85rem;'>暂无事项</span>", unsafe_allow_html=True)
                
                # 渲染卡片
                for memo in reversed(memos): # 最新在前
                    with st.container(border=True):
                        mc1, mc2 = st.columns([8, 2], vertical_alignment="center")
                        with mc1:
                            st.markdown(f"<span style='font-size:0.95rem;'>{memo['text']}</span>", unsafe_allow_html=True)
                        with mc2:
                            if st.button("完成", key=f"del_{memo['id']}", help="标记完成"):
                                remove_memo(memo['id'])
                                st.rerun()

            with col_leader:
                render_memo_column(leader_memos, "领导临时交办", "badge-orange")
            with col_redline:
                render_memo_column(redline_memos, "业务红线", "badge-green")
            with col_question:
                render_memo_column(question_memos, "待询疑问", "badge-blue")
