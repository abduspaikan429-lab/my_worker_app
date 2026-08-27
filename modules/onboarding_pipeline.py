import streamlit as st
import pandas as pd
import io
import re
import json
import os
from datetime import date
import hashlib
import time

from services.onboarding_service import OnboardingService

DATA_FILE = "data/onboarding_data.json"
onboarding_service = OnboardingService()

def load_data():
    return onboarding_service.get_records()

def save_data():
    if "onboarding_data" in st.session_state:
        onboarding_service.save_records(st.session_state.onboarding_data)

def save_data_if_changed():
    """
    仅在数据实际发生变化时才写磁盘，防止每次页面重渲染都触发 I/O。
    利用 MD5 对比当前数据与上次保存时的哈希值，只有不匹配时才执行 save_data。
    """
    if "onboarding_data" not in st.session_state:
        return
    current_hash = hashlib.md5(
        json.dumps(st.session_state.onboarding_data, ensure_ascii=False, sort_keys=True).encode()
    ).hexdigest()
    if st.session_state.get("_onboarding_last_hash") != current_hash:
        save_data()
        st.session_state._onboarding_last_hash = current_hash

PAPER_ITEMS = [
    "体检单", "三级教育", "承诺书", "岗前培训", 
    "签到表按手印(2张)", "花名册", "劳动合同(纸质)", "进场告知书(纸质)"
]
SYSTEM_ITEMS = [
    "更新花名册", "更新月更报表", "更新签到表"
]
ACCESS_ITEMS = [
    "门禁录入完成", "百工聚合同/告知书签署及加卡", "智慧护薪合同发起及工人/班组长确认"
]
TOTAL_ITEMS = len(PAPER_ITEMS) + len(SYSTEM_ITEMS) + len(ACCESS_ITEMS)

def init_empty_worker(name, team=""):
    name = name.strip()
    team = team.strip() if team.strip() else "待分配班组"
    worker_id = f"{name}_{team}"
    
    if worker_id not in st.session_state.onboarding_data:
        st.session_state.onboarding_data[worker_id] = {
            "info": {
                "姓名": name,
                "班组": team,
                "身份证号": "",
                "手机号": "",
                "工种": "",
                "银行卡号": "",
                "进场日期": str(date.today()),
            },
            "paper": {k: False for k in PAPER_ITEMS},
            "system": {k: False for k in SYSTEM_ITEMS},
            "access": {k: False for k in ACCESS_ITEMS},
            "created_at": str(date.today()),
        }
        return True
    return False

def get_progress(worker_data):
    completed = 0
    completed += sum(1 for v in worker_data.get("paper", {}).values() if v)
    completed += sum(1 for v in worker_data.get("system", {}).values() if v)
    completed += sum(1 for v in worker_data.get("access", {}).values() if v)
    return completed, TOTAL_ITEMS

def generate_wechat_notice():
    if "onboarding_data" not in st.session_state:
        return ""
    
    missing_baigongju_map = {}
    missing_huxin_map = {}

    for worker_id, data in st.session_state.onboarding_data.items():
        name = data["info"]["姓名"]
        team = data["info"]["班组"]
        
        missing_huxin = not data["access"]["智慧护薪合同发起及工人/班组长确认"]
        missing_baigongju = not data["access"]["百工聚合同/告知书签署及加卡"]
        
        if missing_baigongju:
            if team not in missing_baigongju_map:
                missing_baigongju_map[team] = []
            missing_baigongju_map[team].append(name)
            
        if missing_huxin:
            if team not in missing_huxin_map:
                missing_huxin_map[team] = []
            missing_huxin_map[team].append(name)
            
    if not missing_baigongju_map and not missing_huxin_map:
        return "所有人员已完成智慧护薪与百工聚确认！"
        
    leader_map = {
        "王宜强": "郭工"
    }
    
    lines = []
    
    for team, workers in missing_baigongju_map.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，提醒一下{workers_str}，在百工聚上签合同和进场告知书，并且添加银行卡一类卡信息哟")
        
    for team, workers in missing_huxin_map.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，提醒{workers_str}在人社小灵光里面签合同哦")
        
    return "\n".join(lines)

def export_to_excel():
    if "onboarding_data" not in st.session_state or not st.session_state.onboarding_data:
        return None
        
    rows = []
    for worker_id, data in st.session_state.onboarding_data.items():
        row = data["info"].copy()
        row.update(data["paper"])
        row.update(data["system"])
        row.update(data["access"])
        
        completed, total = get_progress(data)
        row["进场进度"] = f"{completed}/{total}"
        row["进度百分比"] = f"{int(completed/total*100)}%"
        rows.append(row)
        
    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='进场状态追踪')
    return output.getvalue()

def get_dialog_decorator():
    if hasattr(st, "dialog"): return st.dialog
    if hasattr(st, "experimental_dialog"): return st.experimental_dialog
    return lambda x: lambda f: f

@get_dialog_decorator()("进场手续办理面板")
def worker_dialog(worker_id):
    if worker_id not in st.session_state.onboarding_data:
        st.rerun()
        return
        
    data = st.session_state.onboarding_data[worker_id]
    
    # Ensure keys exist for backward compatibility
    if "paper" not in data: data["paper"] = {}
    if "system" not in data: data["system"] = {}
    if "access" not in data: data["access"] = {}
    
    info = data["info"]
    
    completed, total = get_progress(data)
    progress_pct = int((completed / total) * 100)
    
    st.markdown(f"### :material/person: {info['姓名']} <span style='font-size:16px;color:gray;'>({info['班组']})</span>", unsafe_allow_html=True)
    st.markdown(f"**当前进度**: {completed}/{total} 项完成 ({progress_pct}%)")
    st.markdown(f'<div class="progress-bar-container"><div class="progress-bar-fill" style="width: {progress_pct}%;"></div></div>', unsafe_allow_html=True)
    
    bank_card = str(info.get("银行卡号", "")).strip()
    if bank_card:
        if len(bank_card) < 15:
            st.markdown('<div class="alert-box alert-danger">:material/warning: 警告：银行卡号长度不合规，请核实是否为一类卡！</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="alert-box alert-success">:material/check_circle: 银行卡号：{bank_card}</div>', unsafe_allow_html=True)

    col_p, col_s, col_a = st.columns(3)
    with col_p:
        st.markdown('<span class="tag-badge badge-blue">纸质/电子资料</span>', unsafe_allow_html=True)
        for item in PAPER_ITEMS:
            data["paper"][item] = st.checkbox(item, value=data["paper"].get(item, False), key=f"d_p_{worker_id}_{item}")
    with col_s:
        st.markdown('<span class="tag-badge badge-green">人员信息添加</span>', unsafe_allow_html=True)
        for item in SYSTEM_ITEMS:
            data["system"][item] = st.checkbox(item, value=data["system"].get(item, False), key=f"d_s_{worker_id}_{item}")
    with col_a:
        st.markdown('<span class="tag-badge badge-orange">门禁与平台合同</span>', unsafe_allow_html=True)
        for item in ACCESS_ITEMS:
            label = f"**{item}**" if "百工聚" in item else item
            data["access"][item] = st.checkbox(label, value=data["access"].get(item, False), key=f"d_a_{worker_id}_{item}")

    # 检查是否全部完成
    new_c, new_t = get_progress(data)
    if new_c == new_t and new_t > 0 and data.get("status") != "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.success("🎉 该人员当前展示的手续均已完成！")
        if st.button("归档此记录", type="primary", use_container_width=True):
            onboarding_service.mark_completed(worker_id, True)
            st.rerun()

    if data.get("status") == "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.info(f"✅ 该记录已于 {data.get('completed_at', '')} 归档。")
        if st.button("撤销归档 (恢复办理)", use_container_width=True):
            onboarding_service.mark_completed(worker_id, False)
            st.rerun()

    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
    if st.button("🗑️ 彻底删除该人员记录", key=f"delete_worker_{worker_id}", use_container_width=True):
        if worker_id in st.session_state.onboarding_data:
            del st.session_state.onboarding_data[worker_id]
            st.rerun()

def render():
    if "onboarding_data" not in st.session_state:
        st.session_state.onboarding_data = load_data()

    if "target_worker_id" in st.session_state and st.session_state.target_worker_id:
        worker_id = st.session_state.target_worker_id
        if worker_id in st.session_state.onboarding_data:
            del st.session_state.target_worker_id
            worker_dialog(worker_id)

    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined" style="font-size: 32px; color: #B3A4F3;">rocket_launch</span>
        <div class="header-text">
            <h2>进场流水线管理</h2>
            <p>工人进场流程全生命周期追踪，一站式搞定</p>
        </div>
    </div>
    <div class="color-strip"></div>
    """, unsafe_allow_html=True)
    
    # 最高层级使用 tabs
    tab_add, tab_track = st.tabs(["新增进场人员", "进场流程追踪看板"])

    with tab_add:
        # 左侧单人快速录入，右侧批量导入，包裹在 container 里面
        col_left, col_right = st.columns(2)
        
        with col_left:
            with st.container(border=True):
                st.markdown("#### 单个快速添加")
                st.markdown("<p style='color: #64748B; font-size: 13px;'>仅需姓名即可生成追踪卡片，详细档案可后续补充。</p>", unsafe_allow_html=True)
                with st.form("single_add_form", clear_on_submit=True):
                    new_name = st.text_input("姓名 (必填)*", placeholder="输入工人姓名")
                    new_team_sel = st.selectbox("班组", ["待分配班组", "汪佩沾", "王宜强"])
                    submit_single = st.form_submit_button("快速添加", type="primary", use_container_width=True)
                    
                    if submit_single:
                        if not new_name.strip():
                            st.error("姓名不能为空！")
                        else:
                            actual_team = "" if new_team_sel == "待分配班组" else new_team_sel
                            added = init_empty_worker(new_name, actual_team)
                            if added:
                                st.success(f"成功添加：{new_name}！")
                                st.rerun()
                            else:
                                st.warning(f"该工人 ({new_name} - {actual_team if actual_team else '待分配班组'}) 已存在！")

        with col_right:
            with st.container(border=True):
                st.markdown("#### 批量文本导入")
                st.markdown("<p style='color: #64748B; font-size: 13px;'>直接粘贴多个姓名（用换行、空格或逗号分隔），并指定统一的班组即可快速导入。</p>", unsafe_allow_html=True)
                
                with st.form("batch_text_add_form", clear_on_submit=True):
                    batch_names = st.text_area("批量输入姓名 (必填)*", placeholder="例如:\n张三\n李四\n王五", height=130)
                    batch_team_sel = st.selectbox("统一设置班组", ["待分配班组", "汪佩沾", "王宜强"])
                    submit_batch = st.form_submit_button("批量导入", type="primary", use_container_width=True)
                    
                    if submit_batch:
                        if not batch_names.strip():
                            st.error("请输入至少一个姓名！")
                        else:
                            names_list = re.split(r'[,\s、，]+', batch_names.strip())
                            names_list = [n for n in names_list if n]
                            actual_team = "" if batch_team_sel == "待分配班组" else batch_team_sel
                            
                            success_count = 0
                            duplicate_count = 0
                            for name in names_list:
                                added = init_empty_worker(name, actual_team)
                                if added:
                                    success_count += 1
                                else:
                                    duplicate_count += 1
                                    
                            if success_count > 0:
                                st.success(f"成功添加 {success_count} 名工人！" + (f" （跳过 {duplicate_count} 个重复记录）" if duplicate_count > 0 else ""))
                                st.rerun()
                            elif duplicate_count > 0:
                                st.warning(f"所有人员均已存在（跳过 {duplicate_count} 个记录）。")
                            else:
                                st.warning("未找到有效的姓名。")

    with tab_track:
        if st.session_state.onboarding_data:
            workers_list = list(st.session_state.onboarding_data.items())
            total_workers = len(workers_list)
            completed_count = 0
            incomplete_count = 0
            
            for wid, d in workers_list:
                c, t = get_progress(d)
                if c == t:
                    completed_count += 1
                else:
                    incomplete_count += 1
            
            # 1. 顶部统计指标卡 (看板数据概览)
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("总进场人员", f"{total_workers} 人")
            with m2:
                st.metric("手续齐备", f"{completed_count} 人")
            with m3:
                st.metric("缺材料/办理中", f"{incomplete_count} 人")

            st.markdown("<div style='margin-top: 16px;'></div>", unsafe_allow_html=True)

            # 2. 搜索与操作工具栏
            f1, f2 = st.columns([3, 1])
            with f1:
                search_query = st.text_input(
                    "搜索姓名或班组",
                    placeholder="输入姓名或班组关键词快速过滤...",
                    label_visibility="collapsed"
                )
            with f2:
                st.download_button(
                    label=":material/download: 导出进度 Excel",
                    data=export_to_excel(),
                    file_name="工人进场状态追踪表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            # 3. 流程指南收纳在 Expander 中，避免占用大块视野
            with st.expander(":material/menu_book: 查看标准进场全流程指南", expanded=False):
                st.markdown("""
                1. **线下资料收集**：进场需提交合同、体检单、三级教育、承诺书、岗前培训、两张签到表按手印、花名册等资料。
                2. **一站式登记与扫码**：资料收齐后去一站式大厅录入人员信息，检查浙里办、人社小灵光、云筑网是否注册，并扫进场二维码。
                3. **门禁授权**：等待总包发录门禁通知。
                4. **线上签约与加卡 (关键)**：录完门禁后，提醒工人在【百工聚】签合同、进场通知书及添加银行卡（必须是一类卡）。在【智慧护薪】发起合同后提醒确认。
                5. **台账更新**：更新花名册、变更月报、签到表。
                """)

            st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)

            # 4. 过滤数据
            show_completed = st.checkbox("显示已归档记录", value=False)
            filtered_workers = []
            for worker_id, data in workers_list:
                info = data["info"]
                if not show_completed and data.get("status") == "completed":
                    continue
                if search_query and search_query not in info["姓名"] and search_query not in info["班组"]:
                    continue
                filtered_workers.append((worker_id, data))

            # 5. 分视图展示：卡片办理网格 vs 详细待办表格
            view_tab1, view_tab2 = st.tabs(["人员明细办理区", "全局待办清单"])

            with view_tab1:
                if not filtered_workers:
                    st.info("未找到符合搜索条件的人员")
                else:
                    cols = st.columns(4)
                    for idx, (worker_id, data) in enumerate(filtered_workers):
                        info = data["info"]
                        c, t = get_progress(data)
                        pct = int((c / t) * 100) if t > 0 else 0
                        
                        missing_paper = [k for k, v in data["paper"].items() if not v]
                        missing_system = [k for k, v in data["system"].items() if not v]
                        missing_access = [k for k, v in data["access"].items() if not v]
                        all_missing = missing_paper + missing_system + missing_access
                        
                        with cols[idx % 4]:
                            with st.container(border=True):
                                st.markdown(f"**{info['姓名']}** <span style='font-size:12px;color:#94A3B8;'>({info['班组']})</span>", unsafe_allow_html=True)
                                
                                # 进度条与标签
                                if all_missing:
                                    st.markdown(f'<span class="tag-badge badge-pink">缺 {len(all_missing)} 项手续</span>', unsafe_allow_html=True)
                                else:
                                    st.markdown('<span class="tag-badge badge-green">手续全部齐备</span>', unsafe_allow_html=True)
                                
                                st.markdown(f'<div class="progress-bar-container" style="margin: 10px 0 6px 0;"><div class="progress-bar-fill" style="width: {pct}%;"></div></div>', unsafe_allow_html=True)
                                st.caption(f"已完成: {c}/{t} ({pct}%)")

                                if st.button(":material/edit: 办理", key=f"btn_{worker_id}", use_container_width=True):
                                    worker_dialog(worker_id)

            with view_tab2:
                summary_data = []
                for wid, d in filtered_workers:
                    c, t = get_progress(d)
                    m_paper = [k for k, v in d["paper"].items() if not v]
                    m_system = [k for k, v in d["system"].items() if not v]
                    m_access = [k for k, v in d["access"].items() if not v]
                    m_all = m_paper + m_system + m_access
                    
                    summary_data.append({
                        "姓名": d["info"]["姓名"],
                        "班组": d["info"]["班组"],
                        "状态": "手续齐备" if c == t else f"缺 {len(m_all)} 项材料",
                        "待办事项": "无" if c == t else "、".join(m_all),
                        "进度": f"{c}/{t}"
                    })
                if summary_data:
                    st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

            # 6. 智能催办引擎收纳在 Expander 中
            with st.expander(":material/chat: 智能预警与微信群催办引擎", expanded=False):
                if st.button("生成微信群催办文案"):
                    notice_text = generate_wechat_notice()
                    st.text_area("复制以下文案发送至微信群：", value=notice_text, height=180)

        else:
            st.info("当前暂无人员，请点击【新增进场人员】标签页进行添加。")
            
    # 每次渲染结束后，仅当数据实际发生变化时才写磁盘（哈希对比机制）
    save_data_if_changed()
