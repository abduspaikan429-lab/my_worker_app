import streamlit as st
import pandas as pd
import io
import re
import json
import os
from datetime import date
import hashlib
import time

from services.onboarding_service import (
    PAPER_ITEMS,
    ACCESS_ITEMS,
    PHOTO_ITEMS,
    SYSTEM_ITEMS,
    TOTAL_ITEMS,
    ITEM_ALIASES,
    is_item_done,
    are_all_items_done,
    OnboardingService,
)

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
            "access": {k: False for k in ACCESS_ITEMS},
            "photo": {k: False for k in PHOTO_ITEMS},
            "system": {k: False for k in SYSTEM_ITEMS},
            "created_at": str(date.today()),
        }
        return True
    return False

def get_progress(worker_data):
    return onboarding_service.get_progress(worker_data)

def generate_wechat_notice():
    if "onboarding_data" not in st.session_state:
        return ""
    
    missing_baigongju_map = {}
    missing_huxin_map = {}

    for worker_id, data in st.session_state.onboarding_data.items():
        name = data.get("info", {}).get("姓名", "")
        team = data.get("info", {}).get("班组", "")
        if not name:
            continue
        
        missing_huxin = not is_item_done(data.get("access", {}), "人社小灵光签合同")
        missing_baigongju = not is_item_done(data.get("access", {}), "百工聚加卡签合同")
        
        if missing_baigongju:
            if team not in missing_baigongju_map:
                missing_baigongju_map[team] = []
            missing_baigongju_map[team].append(name)
            
        if missing_huxin:
            if team not in missing_huxin_map:
                missing_huxin_map[team] = []
            missing_huxin_map[team].append(name)
            
    if not missing_baigongju_map and not missing_huxin_map:
        return "所有人员已完成百工聚与人社小灵光确认！"
        
    leader_map = {
        "王宜强": "郭工"
    }
    
    lines = []
    
    for team, workers in missing_baigongju_map.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，提醒一下{workers_str}：在【百工聚】上签合同与告知书，并添加本人一类银行卡信息哟！")
        
    for team, workers in missing_huxin_map.items():
        leader = leader_map.get(team, "汪老板")
        workers_str = "、".join(workers)
        lines.append(f"{leader}，提醒一下{workers_str}：在【人社小灵光】小程序里签署电子劳动合同哦！")
        
    return "\n".join(lines)

def export_to_excel():
    if "onboarding_data" not in st.session_state or not st.session_state.onboarding_data:
        return None
        
    rows = []
    for worker_id, data in st.session_state.onboarding_data.items():
        row = data.get("info", {}).copy()
        
        # 10类纸质材料
        for k in PAPER_ITEMS:
            row[f"纸质_{k}"] = "已收" if is_item_done(data.get("paper", {}), k) else "未收"
        # 门禁与移动端
        for k in ACCESS_ITEMS:
            row[f"平台_{k}"] = "已办" if is_item_done(data.get("access", {}), k) else "未办"
        # 合规影像
        for k in PHOTO_ITEMS:
            row[f"影像_{k}"] = "已留存" if is_item_done(data.get("photo", {}), k) else "未留存"
        # 本地台账
        for k in SYSTEM_ITEMS:
            row[f"台账_{k}"] = "已更新" if is_item_done(data.get("system", {}), k) else "未更新"
            
        completed, total = get_progress(data)
        row["进场总进度"] = f"{completed}/{total}"
        row["进度百分比"] = f"{int(completed/total*100)}%" if total > 0 else "0%"
        rows.append(row)
        
    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='进场全要素状态追踪')
    return output.getvalue()

def get_dialog_decorator():
    if hasattr(st, "dialog"): return st.dialog
    if hasattr(st, "experimental_dialog"): return st.experimental_dialog
    return lambda x: lambda f: f

@get_dialog_decorator()("进场手续办理面板")
def worker_dialog(worker_id):
    if "onboarding_data" not in st.session_state:
        st.session_state.onboarding_data = load_data()
    if worker_id not in st.session_state.onboarding_data:
        st.rerun()
        return
        
    data = st.session_state.onboarding_data[worker_id]
    
    # 确保全部分类字典存在，向下无缝兼容历史记录
    if "paper" not in data: data["paper"] = {}
    if "access" not in data: data["access"] = {}
    if "photo" not in data: data["photo"] = {}
    if "system" not in data: data["system"] = {}
    
    info = data["info"]
    
    completed, total = get_progress(data)
    progress_pct = int((completed / total) * 100) if total > 0 else 0
    
    st.markdown(f"### :material/person: {info['姓名']} <span style='font-size:16px;color:gray;'>({info['班组']})</span>", unsafe_allow_html=True)
    st.markdown(f"**当前进度**: {completed}/{total} 项完成 ({progress_pct}%)")
    st.markdown(f'<div class="progress-bar-container"><div class="progress-bar-fill" style="width: {progress_pct}%;"></div></div>', unsafe_allow_html=True)
    
    bank_card = str(info.get("银行卡号", "") or info.get("工资卡号", "")).strip()
    if bank_card:
        if len(bank_card) < 15:
            st.markdown('<div class="alert-box alert-danger">:material/warning: 警告：银行卡号长度不合规，请核实是否为一类卡！</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="alert-box alert-success">:material/check_circle: 一类工资卡号：{bank_card}</div>', unsafe_allow_html=True)

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown('<span class="tag-badge badge-blue">📄 纸质版审查材料 (10类)</span>', unsafe_allow_html=True)
        for item in PAPER_ITEMS:
            curr_val = is_item_done(data["paper"], item)
            data["paper"][item] = st.checkbox(item, value=curr_val, key=f"d_p_{worker_id}_{item}")

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
        st.markdown('<span class="tag-badge badge-green">📑 本地 Excel 台账更新 (3项)</span>', unsafe_allow_html=True)
        for item in SYSTEM_ITEMS:
            curr_val = is_item_done(data["system"], item)
            data["system"][item] = st.checkbox(item, value=curr_val, key=f"d_s_{worker_id}_{item}")

    with col_right:
        st.markdown('<span class="tag-badge badge-orange">📱 门禁与移动端办理 (5项)</span>', unsafe_allow_html=True)
        for item in ACCESS_ITEMS:
            curr_val = is_item_done(data["access"], item)
            label = f"**{item}**" if ("百工聚" in item or "人社小灵光" in item) else item
            data["access"][item] = st.checkbox(label, value=curr_val, key=f"d_a_{worker_id}_{item}")

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
        st.markdown('<span class="tag-badge badge-pink">📸 必须留存合规影像 (4张)</span>', unsafe_allow_html=True)
        for item in PHOTO_ITEMS:
            curr_val = is_item_done(data["photo"], item)
            data["photo"][item] = st.checkbox(item, value=curr_val, key=f"d_ph_{worker_id}_{item}")

    # 检查是否全部完成
    new_c, new_t = get_progress(data)
    if new_c == new_t and new_t > 0 and data.get("status") != "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.success("🎉 该人员当前展示的手续均已完成！")
        if st.button("🎉 归档并同步至主表 (data/master)", type="primary", use_container_width=True):
            onboarding_service.mark_completed(worker_id, True)
            st.rerun()

    if data.get("status") == "completed":
        st.markdown("<hr style='margin: 15px 0;'/>", unsafe_allow_html=True)
        st.info(f"✅ 该记录已于 {data.get('completed_at', '')} 归档。")
        if data.get("synced_to_master"):
            st.caption(f"🔄 档案已于 {data.get('last_synced_at', '')} 同步至项目主表 (data/master)")
        if st.button("撤销归档 (恢复办理)", use_container_width=True):
            onboarding_service.mark_completed(worker_id, False)
            st.rerun()

    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
    if st.button("🗑️ 彻底删除该人员记录", key=f"delete_worker_{worker_id}", use_container_width=True):
        if worker_id in st.session_state.onboarding_data:
            del st.session_state.onboarding_data[worker_id]
            st.rerun()
            
    save_data_if_changed()

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
            f1, f2, f3 = st.columns([2.5, 1.2, 1.2])
            with f1:
                search_query = st.text_input(
                    "搜索姓名或班组",
                    placeholder="输入姓名或班组关键词快速过滤...",
                    label_visibility="collapsed"
                )
            with f2:
                if st.button(":material/sync: 同步至项目主表", use_container_width=True, help="将进场流水线人员档案同步至项目主表"):
                    sync_res = onboarding_service.sync_to_master()
                    if not sync_res.get("error"):
                        st.success(f"已同步主表：新增 {sync_res.get('added', 0)} 人，更新 {sync_res.get('updated', 0)} 人！")
                        st.rerun()
                    else:
                        st.info(sync_res.get("error"))
            with f3:
                st.download_button(
                    label=":material/download: 导出进度 Excel",
                    data=export_to_excel(),
                    file_name="工人进场状态追踪表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            # 3. 流程指南收纳在 Expander 中，避免占用大块视野
            with st.expander(":material/menu_book: 查看标准进场全流程指南 (合规必读)", expanded=False):
                st.markdown("""
                1. **10类纸质材料审查**：简易合同、体检单、三级教育、承诺书、进场承诺书、岗前培训试题、两张签到表按手印、花名册、退场承诺书、离场结算单。
                2. **门禁与移动端4项**：录门禁、扫现场进场码、注册【浙里办-工人保障在线】、在【百工聚】添加一类银行卡并签合同、在【人社小灵光】签约。
                3. **4张合规影像留存**：务工人员手持身份证+本人一类工资卡合影；手持合同封面、日工资页、签字页3张影像套件。
                4. **两系统导出与整合**：在【智慧护薪】与【三局系统】沉淀后导出表格，使用【档案魔法整合】自动清洗去重生成《全量信息表》与《中建二局标准档案表》。
                5. **台账回填与二局归档**：在【花名册与报表导出】复制新增人员回填本地 Excel（花名册、进场变更月报、水印签到表）；扫描纸质合同与进场承诺书上传【中建二局系统】完成闭环。
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
                        
                        m_paper = [k for k in PAPER_ITEMS if not is_item_done(data.get("paper", {}), k)]
                        m_access = [k for k in ACCESS_ITEMS if not is_item_done(data.get("access", {}), k)]
                        m_photo = [k for k in PHOTO_ITEMS if not is_item_done(data.get("photo", {}), k)]
                        m_system = [k for k in SYSTEM_ITEMS if not is_item_done(data.get("system", {}), k)]
                        all_missing = m_paper + m_access + m_photo + m_system
                        
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
                    m_paper = [k for k in PAPER_ITEMS if not is_item_done(d.get("paper", {}), k)]
                    m_access = [k for k in ACCESS_ITEMS if not is_item_done(d.get("access", {}), k)]
                    m_photo = [k for k in PHOTO_ITEMS if not is_item_done(d.get("photo", {}), k)]
                    m_system = [k for k in SYSTEM_ITEMS if not is_item_done(d.get("system", {}), k)]
                    m_all = m_paper + m_access + m_photo + m_system
                    
                    summary_data.append({
                        "姓名": d["info"]["姓名"],
                        "班组": d["info"]["班组"],
                        "状态": "手续齐备" if c == t else f"缺 {len(m_all)} 项手续",
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
