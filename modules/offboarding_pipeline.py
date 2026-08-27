import streamlit as st
import pandas as pd
import io
import json
import os
import hashlib
import time
from datetime import date
from modules.master_data import load_master_df
from modules.onboarding_pipeline import onboarding_service

from services.offboarding_service import OffboardingService

DATA_FILE = "data/offboarding_data.json"
HISTORY_FILE = "data/offboarding_history.json"
offboarding_service = OffboardingService()


def load_offboarding_history() -> list[dict]:
    """加载历史离场归档记录列表。"""
    return offboarding_service.load_history()


def archive_offboarding(worker_id: str, data: dict) -> None:
    """将已完成离场手续的人员追加写入离场归档文件，记录离场日期。"""
    offboarding_service.archive_offboarding(worker_id, data)


def load_data():
    return offboarding_service.get_records()


def save_data():
    if "offboarding_data" in st.session_state:
        offboarding_service.save_records(st.session_state.offboarding_data)

def save_data_if_changed():
    """
    仅在数据实际发生变化时才写磁盘，防止每次页面重渲染都触发 I/O。
    利用 MD5 对比当前数据与上次保存时的哈希值，只有不匹配时才执行 save_data。
    """
    if "offboarding_data" not in st.session_state:
        return
    current_hash = hashlib.md5(
        json.dumps(st.session_state.offboarding_data, ensure_ascii=False, sort_keys=True).encode()
    ).hexdigest()
    if st.session_state.get("_offboarding_last_hash") != current_hash:
        save_data()
        st.session_state._offboarding_last_hash = current_hash

OFFBOARDING_STEPS = [
    "1. 工人小灵光发起",
    "2. 班组长确认",
    "3. 劳资员（我）确认",
    "4. 提交财务发放",
    "5. 财务发放完成"
]
OFF_PAPER = ["收纸质离场结算单", "归档身份证+结算单照片", "结清证明上传"]
OFF_SYSTEM = ["处理离场月报", "更新花名册", "更新签到表"]
TOTAL_ITEMS = len(OFFBOARDING_STEPS) + len(OFF_PAPER) + len(OFF_SYSTEM)

def init_offboarding_worker(name, team, source_id=""):
    try:
        offboarding_service.create_offboarding({
            "info": {
                "姓名": name,
                "班组": team,
                "身份证号": source_id,
            }
        })
        return True
    except Exception:
        return False

def get_progress(worker_data):
    steps = worker_data.get("steps", {})
    paper = worker_data.get("paper", {})
    system = worker_data.get("system", {})
    completed = sum(1 for v in steps.values() if v) + sum(1 for v in paper.values() if v) + sum(1 for v in system.values() if v)
    return completed, TOTAL_ITEMS

def generate_wechat_notice():
    if "offboarding_data" not in st.session_state:
        return ""
    
    missing_worker_init = {}
    missing_leader_confirm = {}

    for worker_id, data in st.session_state.offboarding_data.items():
        name = data["info"]["姓名"]
        team = data["info"]["班组"]
        
        step1 = data["steps"].get("1. 工人小灵光发起", False)
        step2 = data["steps"].get("2. 班组长确认", False)
        
        if not step1:
            if team not in missing_worker_init:
                missing_worker_init[team] = []
            missing_worker_init[team].append(name)
        elif not step2: # 工人发起了，但是班组长没确认
            if team not in missing_leader_confirm:
                missing_leader_confirm[team] = []
            missing_leader_confirm[team].append(name)
            
    if not missing_worker_init and not missing_leader_confirm:
        return "所有离场人员均已完成初步发起与确认！"
        
    leader_map = {
        "王宜强": "郭工",
    }
    
    lines = []
    
    for team, workers in missing_worker_init.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，麻烦提醒一下这几位准备离场的兄弟：{workers_str}，在人社小灵光发起离场结算哦~")
        
    for team, workers in missing_leader_confirm.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，这几位兄弟已经发起了退场：{workers_str}，麻烦您在小灵光的待办事项里点一下班组长确认。")
        
    return "\n".join(lines)


def get_dialog_decorator():
    if hasattr(st, "dialog"): return st.dialog
    if hasattr(st, "experimental_dialog"): return st.experimental_dialog
    return lambda x: lambda f: f

@get_dialog_decorator()("离场手续办理面板")
def worker_dialog(worker_id):
    if worker_id not in st.session_state.offboarding_data:
        st.rerun()
        return
        
    data = st.session_state.offboarding_data[worker_id]
    
    if "steps" not in data: data["steps"] = {k: False for k in OFFBOARDING_STEPS}
    if "paper" not in data: data["paper"] = {k: False for k in OFF_PAPER}
    if "system" not in data: data["system"] = {k: False for k in OFF_SYSTEM}
    
    info = data["info"]
    completed, total = get_progress(data)
    progress_pct = int((completed / total) * 100)
    
    st.markdown(f"### :material/person_remove: {info['姓名']} <span style='font-size:16px;color:gray;'>({info['班组']})</span>", unsafe_allow_html=True)
    st.markdown(f"**离场进度**: {completed}/{total} 步完成 ({progress_pct}%)")
    st.markdown(f'<div class="progress-bar-container"><div class="progress-bar-fill" style="width: {progress_pct}%; background: linear-gradient(90deg, #F87171, #FCA5A5);"></div></div>', unsafe_allow_html=True)
    
    st.markdown('<div style="margin-top: 15px;"></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<span class="tag-badge badge-orange">小灵光及财务流程</span>', unsafe_allow_html=True)
        for item in OFFBOARDING_STEPS:
            data["steps"][item] = st.checkbox(item, value=data["steps"].get(item, False), key=f"d_off_{worker_id}_{item}")
    with col2:
        st.markdown('<span class="tag-badge badge-blue">纸质与证明材料</span>', unsafe_allow_html=True)
        for item in OFF_PAPER:
            data["paper"][item] = st.checkbox(item, value=data["paper"].get(item, False), key=f"d_op_{worker_id}_{item}")
    with col3:
        st.markdown('<span class="tag-badge badge-green">台账更新（尾项）</span>', unsafe_allow_html=True)
        for item in OFF_SYSTEM:
            data["system"][item] = st.checkbox(item, value=data["system"].get(item, False), key=f"d_os_{worker_id}_{item}")

    # 检查是否全部完成
    new_c, new_t = get_progress(data)
    if new_c == new_t and new_t > 0 and data.get("status") != "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.success("🎉 该人员所有离场手续均已完成！")
        if st.button("归档此记录", type="primary", use_container_width=True):
            offboarding_service.mark_completed(worker_id, True)
            st.rerun()
            
    if data.get("status") == "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.info(f"✅ 该记录已于 {data.get('completed_at', '')} 归档。")
        if st.button("撤销归档 (恢复办理)", use_container_width=True):
            offboarding_service.mark_completed(worker_id, False)
            st.rerun()

    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    if st.button("🗑️ 彻底删除该离场记录", key=f"delete_off_{worker_id}", use_container_width=True):
        if worker_id in st.session_state.offboarding_data:
            del st.session_state.offboarding_data[worker_id]
            st.rerun()

def render():
    if "offboarding_data" not in st.session_state:
        st.session_state.offboarding_data = load_data()

    if "target_worker_id" in st.session_state and st.session_state.target_worker_id:
        worker_id = st.session_state.target_worker_id
        if worker_id in st.session_state.offboarding_data:
            del st.session_state.target_worker_id
            worker_dialog(worker_id)

    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined" style="font-size: 32px; color: #F87171;">flight_takeoff</span>
        <div class="header-text">
            <h2>离场流水线追踪</h2>
            <p>人员退场全流程闭环跟进，工资结算一键催办</p>
        </div>
    </div>
    <div class="color-strip" style="background: linear-gradient(90deg, #FCA5A5, #FEF2F2);"></div>
    """, unsafe_allow_html=True)
    
    tab_add, tab_track = st.tabs(["发起人员离场", "离场进度看板"])

    with tab_add:
        # 加载主表与进场流程人员，并只保留当前在场人员（已进入离场结算或已归档人员不重复展示）
        df = load_master_df()
        all_workers = onboarding_service.merge_with_master(df)
        onsite_df = offboarding_service.filter_onsite_df(all_workers)
        
        st.markdown("#### 从花名册中选择人员发起离场")
        st.markdown("<p style='color: #64748B; font-size: 13px;'>在下方选择当前在场的工人，将其移入离场待办清单。</p>", unsafe_allow_html=True)
        
        if not onsite_df.empty:
            onsite_df = onsite_df.copy()
            onsite_df['display_name'] = onsite_df['姓名'] + " (" + onsite_df.get('班组', onsite_df.get('工种', '未知')) + ")"
            options = onsite_df.to_dict('records')
            
            # Use multiselect to pick workers
            selected_workers = st.multiselect(
                "选择待离场人员",
                options=options,
                format_func=lambda x: x['display_name'],
                placeholder="搜索并选择需要离场的工人..."
            )
            
            if st.button("🚀 批量发起离场流程", type="primary"):
                if not selected_workers:
                    st.error("请至少选择一名工人！")
                else:
                    success_count = 0
                    for worker in selected_workers:
                        name = str(worker.get("姓名", ""))
                        team = str(worker.get("班组", worker.get("工种", "")))
                        sid = str(worker.get("身份证号", ""))
                        if init_offboarding_worker(name, team, sid):
                            success_count += 1
                    st.success(f"成功将 {success_count} 名工人加入离场追踪清单！请前往【离场进度看板】办理手续。")
                    time.sleep(0.5)
                    st.rerun()
        else:
            if df.empty:
                st.warning("当前主表中没有人员数据。请先在【档案魔法整合】中同步人员台账。")
            else:
                st.info("当前主表中的所有人员均已在离场结算中或已归档离场。")
            
        st.markdown("---")
        st.markdown("#### 或手动输入姓名发起离场")
        with st.form("manual_offboarding_form", clear_on_submit=True):
            new_name = st.text_input("姓名 (必填)*", placeholder="输入离场工人姓名")
            new_team = st.selectbox("班组", ["待分配班组", "汪佩沾", "王宜强"])
            if st.form_submit_button("手动发起离场"):
                if not new_name.strip():
                    st.error("姓名不能为空！")
                else:
                    actual_team = "" if new_team == "待分配班组" else new_team
                    if init_offboarding_worker(new_name, actual_team):
                        st.success(f"成功添加离场任务：{new_name}！")
                    else:
                        st.warning("该工人已在离场追踪列表中。")

    with tab_track:
        if st.session_state.offboarding_data:
            workers_list = list(st.session_state.offboarding_data.items())
            total_workers = len(workers_list)
            completed_count = 0
            incomplete_count = 0
            
            for wid, d in workers_list:
                c, t = get_progress(d)
                if c == t:
                    completed_count += 1
                else:
                    incomplete_count += 1
            
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("正在办理离场", f"{total_workers} 人")
            with m2:
                st.metric("流程已完结", f"{completed_count} 人")
            with m3:
                st.metric("手续卡壳中", f"{incomplete_count} 人")

            st.markdown("<div style='margin-top: 16px;'></div>", unsafe_allow_html=True)

            f1, f2 = st.columns([3, 1])
            with f1:
                search_query = st.text_input(
                    "搜索离场人员",
                    placeholder="输入姓名或班组关键词快速过滤...",
                    label_visibility="collapsed"
                )

            with st.expander(":material/chat: 智能催办引擎 (催发起/催确认)", expanded=False):
                if st.button("生成微信群离场催办文案"):
                    notice_text = generate_wechat_notice()
                    st.text_area("复制以下文案发送至微信群：", value=notice_text, height=180)

            st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)

            filtered_workers = []
            show_completed = st.checkbox("显示已归档记录", value=False)
            for worker_id, data in workers_list:
                info = data["info"]
                if not show_completed and data.get("status") == "completed":
                    continue
                if search_query and search_query not in info["姓名"] and search_query not in info["班组"]:
                    continue
                filtered_workers.append((worker_id, data))

            if not filtered_workers:
                st.info("未找到符合搜索条件的人员")
            else:
                cols = st.columns(4)
                for idx, (worker_id, data) in enumerate(filtered_workers):
                    info = data["info"]
                    c, t = get_progress(data)
                    pct = int((c / t) * 100) if t > 0 else 0
                    
                    with cols[idx % 4]:
                        with st.container(border=True):
                            st.markdown(f"**{info['姓名']}** <span style='font-size:12px;color:#94A3B8;'>({info['班组']})</span>", unsafe_allow_html=True)
                            
                            if c == t:
                                st.markdown('<span class="tag-badge badge-green">已结清退场</span>', unsafe_allow_html=True)
                            else:
                                st.markdown('<span class="tag-badge badge-orange">办理中</span>', unsafe_allow_html=True)
                            
                            st.markdown(f'<div class="progress-bar-container" style="margin: 10px 0 6px 0;"><div class="progress-bar-fill" style="width: {pct}%; background: linear-gradient(90deg, #F87171, #FCA5A5);"></div></div>', unsafe_allow_html=True)
                            st.caption(f"已办理: {c}/{t}")

                            if st.button(":material/edit: 办理手续", key=f"btn_off_{worker_id}", use_container_width=True):
                                worker_dialog(worker_id)
        else:
            st.info("当前暂无办理离场的人员，请点击【发起人员离场】进行添加。")
            
    save_data_if_changed()
