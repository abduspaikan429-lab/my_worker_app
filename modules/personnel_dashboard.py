# modules/personnel_dashboard.py
# 月度人员变更中心 - 历史与实时进离场自动关联计算 (细化到班组级别进退场)
from __future__ import annotations
import json, os, calendar, re
from datetime import date, datetime
import pandas as pd
import streamlit as st

from modules.master_data import load_master_df
from modules.onboarding_pipeline import onboarding_service
from modules.offboarding_pipeline import load_offboarding_history

MONTHLY_DATA_FILE = "data/monthly_change_data.json"
SWITCHOVER_MONTH = "2026-09" # 6, 7, 8 为纯历史输入
START_MONTH = "2026-06"

def _get_report_date_str(month_str):
    if not month_str or "-" not in month_str:
        return ""
    y, m = map(int, month_str.split('-'))
    pm = m - 1
    py = y
    if pm == 0:
        pm = 12
        py -= 1
    _, p_days = calendar.monthrange(py, pm)
    _, c_days = calendar.monthrange(y, m)
    return f"{py} 年 {pm} 月 {p_days} 日至 {y} 年 {m} 月 {c_days} 日"

def _parse_report_text(text):
    results = {}
    sections = text.replace("①", "").replace("②", "").replace("③", "").split("劳务（专业）分包单位")
    for sec in sections:
        if not sec.strip(): continue
        
        team_match = re.search(r"班组名称[：:]\s*([^\s]+)", sec)
        if not team_match: continue
        team = team_match.group(1).strip()
        
        in_match = re.search(r"本月进场务工人员总数[：:]\s*(\d+)", sec)
        out_match = re.search(r"本月离场务工人员总数[：:]\s*(\d+)", sec)
        cur_match = re.search(r"本月现场务工人员总数[：:]\s*(\d+)", sec)
        
        if in_match and out_match and cur_match:
            t_key = "total"
            if "王宜强" in team: t_key = "王宜强施工班组"
            elif "汪佩沾" in team: t_key = "汪佩沾其他班组"
            elif "金属屋面" in team: t_key = "total"
            
            results[t_key] = {
                "in_count": int(in_match.group(1)),
                "out_count": int(out_match.group(1)),
                "current_count": int(cur_match.group(1))
            }
    return results


def _load_monthly_data():
    if os.path.exists(MONTHLY_DATA_FILE):
        try:
            with open(MONTHLY_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # 默认空数据结构，支持细分到班组的进场、离场、在场
    return {
        "2026-06": {
            "in_count": 0, "out_count": 0, "current_count": 0, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 0, "out_count": 0, "current_count": 0}, 
                "王宜强施工班组": {"in_count": 0, "out_count": 0, "current_count": 0}
            }
        },
        "2026-07": {
            "in_count": 0, "out_count": 0, "current_count": 0, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 0, "out_count": 0, "current_count": 0}, 
                "王宜强施工班组": {"in_count": 0, "out_count": 0, "current_count": 0}
            }
        },
        "2026-08": {
            "in_count": 0, "out_count": 0, "current_count": 0, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 0, "out_count": 0, "current_count": 0}, 
                "王宜强施工班组": {"in_count": 0, "out_count": 0, "current_count": 0}
            }
        }
    }

def _save_monthly_data(data):
    os.makedirs("data", exist_ok=True)
    with open(MONTHLY_DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _generate_months(start, end):
    sy, sm = map(int, start.split('-'))
    ey, em = map(int, end.split('-'))
    months = []
    cy, cm = sy, sm
    while (cy < ey) or (cy == ey and cm <= em):
        months.append(f"{cy}-{cm:02d}")
        cm += 1
        if cm > 12:
            cm = 1
            cy += 1
    return months

def _parse_date_to_month(val):
    if not val:
        return None
    val_str = str(val).strip()
    if not val_str or val_str in ("nan", "None", ""):
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return val_str[:7] if len(val_str) >= 7 and "-" in val_str else None

def _format_date(val):
    val_str = str(val).strip()
    if not val_str or val_str in ("nan", "None", ""):
        return ""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return val_str

def _map_team(t_name):
    name = str(t_name).strip()
    if not name or name in ("nan", "None", "待分配班组"):
        return "待分配"
    if "王宜强" in name or "旭之升" in name:
        return "王宜强施工班组"
    if "汪佩沾" in name or "久昌" in name:
        return "汪佩沾其他班组"
    return name

def get_dynamic_month_data(target_month):
    """
    账本法计算：从 2026-08 历史月作为基准起算，逐月累加进场，减去离场，得出指定月的月末人数。
    """
    saved_data = _load_monthly_data()
    base_month = "2026-08"
    months_to_calc = _generate_months(SWITCHOVER_MONTH, target_month)
    
    base_data = saved_data.get(base_month, {"current_count": 0, "teams": {}})
    
    # 初始化团队的当前人数
    current_teams = {}
    for t_name, t_data in base_data.get("teams", {}).items():
        if isinstance(t_data, dict):
            current_teams[t_name] = t_data.get("current_count", 0)
        else:
            current_teams[t_name] = t_data # 兼容旧格式
            
    # 从底层班组累加出严格准确的期初总人数，防止历史数据中全局总数与班组总数不一致
    current_total = sum(current_teams.values())
    
    master = load_master_df()
    all_workers = onboarding_service.merge_with_master(master)
    history = load_offboarding_history()
    
    entries_by_month = {}
    for _, row in all_workers.iterrows():
        d_val = row.get("进场日期") or row.get("进场时间") or row.get("入场日期")
        m = _parse_date_to_month(d_val)
        if m:
            entries_by_month.setdefault(m, []).append(row.to_dict())
            
    exits_by_month = {}
    seen_exit_ids = set()
    for rec in sorted(history, key=lambda x: x.get("离场日期", ""), reverse=True):
        id_card = str(rec.get("身份证号") or "").strip()
        name = str(rec.get("姓名") or "").strip()
        team = str(rec.get("班组") or "").strip()
        unique_id = id_card if id_card else f"{name}_{team}"
        
        if unique_id in seen_exit_ids:
            continue
        seen_exit_ids.add(unique_id)
        
        m = _parse_date_to_month(rec.get("离场日期"))
        if m:
            exits_by_month.setdefault(m, []).append(rec)

    target_in_list = []
    target_out_list = []
    
    target_team_stats = {}
    
    for m in months_to_calc:
        in_recs = entries_by_month.get(m, [])
        out_recs = exits_by_month.get(m, [])
        
        if m == target_month:
            target_in_list = in_recs
            target_out_list = out_recs
            # 初始化该月的班组新增减少统计
            for r in in_recs:
                t = _map_team(r.get("班组", ""))
                target_team_stats.setdefault(t, {"in_count": 0, "out_count": 0})
                target_team_stats[t]["in_count"] += 1
            for r in out_recs:
                t = _map_team(r.get("班组", ""))
                target_team_stats.setdefault(t, {"in_count": 0, "out_count": 0})
                target_team_stats[t]["out_count"] += 1
            
        for r in in_recs:
            current_total += 1
            t = _map_team(r.get("班组", ""))
            current_teams[t] = current_teams.get(t, 0) + 1
            
        for r in out_recs:
            current_total -= 1
            t = _map_team(r.get("班组", ""))
            current_teams[t] = current_teams.get(t, 0) - 1
            
    # 构建最终的 teams 结构
    final_teams = {}
    for t_name, current_cnt in current_teams.items():
        stats = target_team_stats.get(t_name, {"in_count": 0, "out_count": 0})
        final_teams[t_name] = {
            "in_count": stats["in_count"],
            "out_count": stats["out_count"],
            "current_count": current_cnt
        }
            
    return {
        "in_count": len(target_in_list),
        "out_count": len(target_out_list),
        "current_count": current_total,
        "teams": final_teams,
        "in_list": target_in_list,
        "out_list": target_out_list
    }

def _kpi_card(label, value, color="#6366F1"):
    st.markdown(f"""<div class="kpi-card" style="border-left:4px solid {color};">
        <div class="kpi-value" style="color:{color};">{value}</div>
        <div class="kpi-label">{label}</div></div>""", unsafe_allow_html=True)

def render():
    st.markdown("""<style>
    .kpi-card{background:#fff;border-radius:14px;padding:20px 22px 16px;box-shadow:0 2px 12px rgba(0,0,0,0.07);margin-bottom:4px;}
    .kpi-value{font-size:2.2rem;font-weight:700;line-height:1.1;}
    .kpi-label{font-size:.82rem;color:#64748B;margin-top:4px;font-weight:500;}
    .team-box{background:#f8fafc; border-left:4px solid #6366F1; padding: 16px; border-radius: 8px;}
    .team-item{display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dashed #e2e8f0; font-size: 0.95em;}
    .team-total{display: flex; justify-content: space-between; padding: 12px 0 0 0; font-weight: bold; font-size: 1.1em;}
    </style>""", unsafe_allow_html=True)
    
    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined" style="font-size:32px;color:#6366F1;">calendar_month</span>
        <div class="header-text"><h2>月度人员变动中心</h2>
        <p>自动关联进退场，精准呈现每月月末在场与各班组进出场数字</p></div>
    </div>
    <div class="color-strip" style="background:linear-gradient(90deg,#C7D2FE,#EDE9FE);"></div>
    """, unsafe_allow_html=True)

    today_month = date.today().strftime("%Y-%m")
    all_months = _generate_months(START_MONTH, max(SWITCHOVER_MONTH, today_month))
    all_months.reverse() # 最新的月份在最上面
    
    col_sel, _ = st.columns([1, 2])
    with col_sel:
        selected_month = st.selectbox("📅 查看月份", options=all_months, index=0)
    
    is_historical = selected_month < SWITCHOVER_MONTH
    
    saved_data = _load_monthly_data()
    
    if is_historical:
        month_data = saved_data.get(selected_month, {
            "in_count": 0, "out_count": 0, "current_count": 0, 
            "teams": {
                "汪佩沾其他班组": {"in_count": 0, "out_count": 0, "current_count": 0}, 
                "王宜强施工班组": {"in_count": 0, "out_count": 0, "current_count": 0}
            }
        })
        
        t_yi = month_data["teams"].setdefault("王宜强施工班组", {"in_count": 0, "out_count": 0, "current_count": 0})
        t_wang = month_data["teams"].setdefault("汪佩沾其他班组", {"in_count": 0, "out_count": 0, "current_count": 0})
        
        t_yi["in_count"] = st.session_state.get(f"hist_yi_in_{selected_month}", t_yi.get("in_count", 0))
        t_yi["out_count"] = st.session_state.get(f"hist_yi_out_{selected_month}", t_yi.get("out_count", 0))
        t_yi["current_count"] = st.session_state.get(f"hist_yi_cur_{selected_month}", t_yi.get("current_count", 0))
        
        t_wang["in_count"] = st.session_state.get(f"hist_wang_in_{selected_month}", t_wang.get("in_count", 0))
        t_wang["out_count"] = st.session_state.get(f"hist_wang_out_{selected_month}", t_wang.get("out_count", 0))
        t_wang["current_count"] = st.session_state.get(f"hist_wang_cur_{selected_month}", t_wang.get("current_count", 0))
        
        # 强制总人数自动等于各班组之和，避免手动输入不一致导致上方数据没跟着变
        month_data["in_count"] = t_yi["in_count"] + t_wang["in_count"]
        month_data["out_count"] = t_yi["out_count"] + t_wang["out_count"]
        month_data["current_count"] = t_yi["current_count"] + t_wang["current_count"]
        
    else:
        with st.spinner("正在基于台账账本计算月末人数..."):
            month_data = get_dynamic_month_data(selected_month)
        
    st.markdown(f"### {selected_month.replace('-', '年')}月 人员变动")
    c1, c2, c3 = st.columns(3)
    with c1:
        _kpi_card("本月进场汇总", f'{month_data["in_count"]} 人', "#34D399")
    with c2:
        _kpi_card("本月离场汇总", f'{month_data["out_count"]} 人', "#F87171")
    with c3:
        _kpi_card("月末在场汇总", f'{month_data["current_count"]} 人', "#6366F1")
        
    st.markdown("<div style='margin-top:24px'></div>", unsafe_allow_html=True)
    
    st.markdown("#### 🏢 班组数据明细 (包含进场/离场/月末在场)")
    
    teams_dict = month_data.get("teams", {})
    metal_roof_keys = ["汪佩沾其他班组", "王宜强施工班组"]
    
    # 金属屋面合计计算
    metal_in = 0
    metal_out = 0
    metal_current = 0
    
    html = '<div class="team-box"><div style="font-weight:bold; font-size:1.2em; margin-bottom:12px;">金属屋面 (专业分包总称)</div>'
    for k in metal_roof_keys:
        t_data = teams_dict.get(k, {})
        if not isinstance(t_data, dict):
            # 兼容老数据结构
            t_data = {"in_count": 0, "out_count": 0, "current_count": t_data}
            
        c_in = t_data.get("in_count", 0)
        c_out = t_data.get("out_count", 0)
        c_cur = t_data.get("current_count", 0)
        
        metal_in += c_in
        metal_out += c_out
        metal_current += c_cur
        
        html += f'''
<div class="team-item">
    <span style="font-weight:600; min-width: 150px;">├── {k}</span>
    <span style="color:#10B981;">进场: {c_in}</span>
    <span style="color:#EF4444;">离场: {c_out}</span>
    <span style="color:#4F46E5;">月末在场: {c_cur}</span>
</div>
'''
        
    html += f'''
<div class="team-total" style="border-top: 1px solid #CBD5E1; padding-top: 12px; margin-top: 8px;">
    <span style="min-width: 150px;">└── 金属屋面合计</span>
    <span style="color:#10B981;">进场: {metal_in}</span>
    <span style="color:#EF4444;">离场: {metal_out}</span>
    <span style="color:#4F46E5;">月末在场: {metal_current}</span>
</div>
</div>
'''
    
    other_teams = {k: v for k, v in teams_dict.items() if k not in metal_roof_keys and (isinstance(v, dict) and v.get("current_count", 0) > 0)}
    if other_teams:
        html += '<div class="team-box" style="margin-top: 16px;"><div style="font-weight:bold; font-size:1.2em; margin-bottom:12px;">其他</div>'
        for k, v in other_teams.items():
            if not isinstance(v, dict):
                v = {"in_count": 0, "out_count": 0, "current_count": v}
            html += f'''
<div class="team-item">
    <span style="font-weight:600; min-width: 150px;">├── {k}</span>
    <span style="color:#10B981;">进场: {v.get('in_count',0)}</span>
    <span style="color:#EF4444;">离场: {v.get('out_count',0)}</span>
    <span style="color:#4F46E5;">月末在场: {v.get('current_count',0)}</span>
</div>
'''
        html += '</div>'
        
    st.markdown(html, unsafe_allow_html=True)
    
    # 历史月份专属：保存与修改数字功能
    if is_historical:
        st.markdown("<div style='margin-top:32px'></div>", unsafe_allow_html=True)
        st.info("📌 当前为历史月份，系统不从台账中反推数据。您可以随时补录或修改报表数字。")
        with st.expander(f"✏️ 录入/修改 {selected_month} 历史月报数据", expanded=False):
            tab_form, tab_parse = st.tabs(["📝 表单与生成的报文", "🤖 智能文本解析导入"])
            
            with tab_form:
                # 获取数据库里最初始的数据作为 fallback
                init_data = saved_data.get(selected_month, {})
                init_yi = init_data.get("teams", {}).get("王宜强施工班组", {}) if isinstance(init_data.get("teams", {}).get("王宜强施工班组"), dict) else {}
                init_wang = init_data.get("teams", {}).get("汪佩沾其他班组", {}) if isinstance(init_data.get("teams", {}).get("汪佩沾其他班组"), dict) else {}
                
                st.markdown("#### 王宜强施工班组")
                c11, c12, c13 = st.columns(3)
                yi_in = c11.number_input("王宜强 - 进场", min_value=0, value=init_yi.get("in_count",0), step=1, key=f"hist_yi_in_{selected_month}")
                yi_out = c12.number_input("王宜强 - 离场", min_value=0, value=init_yi.get("out_count",0), step=1, key=f"hist_yi_out_{selected_month}")
                yi_cur = c13.number_input("王宜强 - 在场", min_value=0, value=init_yi.get("current_count",0), step=1, key=f"hist_yi_cur_{selected_month}")

                st.markdown("#### 汪佩沾其他班组")
                c21, c22, c23 = st.columns(3)
                wang_in = c21.number_input("汪佩沾 - 进场", min_value=0, value=init_wang.get("in_count",0), step=1, key=f"hist_wang_in_{selected_month}")
                wang_out = c22.number_input("汪佩沾 - 离场", min_value=0, value=init_wang.get("out_count",0), step=1, key=f"hist_wang_out_{selected_month}")
                wang_cur = c23.number_input("汪佩沾 - 在场", min_value=0, value=init_wang.get("current_count",0), step=1, key=f"hist_wang_cur_{selected_month}")
                
                st.markdown("---")
                
                st.markdown("#### 总人数 (金属屋面) - *自动汇总*")
                col1, col2, col3 = st.columns(3)
                new_in = col1.number_input("本月进场总数", value=yi_in + wang_in, disabled=True)
                new_out = col2.number_input("本月离场总数", value=yi_out + wang_out, disabled=True)
                new_current = col3.number_input("月末现场总数", value=yi_cur + wang_cur, disabled=True)
                
                st.markdown("#### ✨ 自动生成的报表文字")
                date_str = _get_report_date_str(selected_month)
                text_total = f"①总：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  金属屋面         日期： {date_str}  本月进场务工人员总数:    {new_in}          本月离场务工人员总数:   {new_out}         本月现场务工人员总数：   {new_current}"
                text_wang = f"②分：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  汪佩沾其它班组         日期： {date_str}  本月进场务工人员总数:    {wang_in}         本月离场务工人员总数:   {wang_out}         本月现场务工人员总数：   {wang_cur}"
                text_yi = f"③分：劳务（专业）分包单位： 中建二局安装工程有限公司        班组名称：  王宜强施工班组         日期： {date_str}  本月进场务工人员总数:    {yi_in}          本月离场务工人员总数:   {yi_out}         本月现场务工人员总数：   {yi_cur}"
                st.code(f"{text_total}\n{text_wang}\n{text_yi}", language="text")
                
                if st.button("💾 保存历史报表数据", type="primary", use_container_width=True):
                    if selected_month not in saved_data:
                        saved_data[selected_month] = {}
                    saved_data[selected_month].update({
                        "in_count": new_in,
                        "out_count": new_out,
                        "current_count": new_current,
                        "teams": {
                            "王宜强施工班组": {"in_count": yi_in, "out_count": yi_out, "current_count": yi_cur},
                            "汪佩沾其他班组": {"in_count": wang_in, "out_count": wang_out, "current_count": wang_cur}
                        }
                    })
                    _save_monthly_data(saved_data)
                    st.success(f"{selected_month} 数据保存成功！")
                    st.rerun()
            
            with tab_parse:
                st.markdown("将图文识别或复制的文字粘贴在此处，系统将自动解析并录入：")
                paste_text = st.text_area("请粘贴报表文字：", height=150, help="可以一次性粘贴总计和各个分包班组的内容，系统会自动识别 '班组名称' 和对应人数。")
                if st.button("🤖 解析并保存", type="primary"):
                    parsed_res = _parse_report_text(paste_text)
                    if not parsed_res:
                        st.error("未能解析出有效的报表数据，请检查文字格式是否正确。")
                    else:
                        if selected_month not in saved_data:
                            saved_data[selected_month] = {}
                        
                        updates = {}
                        if "total" in parsed_res:
                            updates["in_count"] = parsed_res["total"]["in_count"]
                            updates["out_count"] = parsed_res["total"]["out_count"]
                            updates["current_count"] = parsed_res["total"]["current_count"]
                        
                        teams_update = saved_data[selected_month].get("teams", {}).copy()
                        if "汪佩沾其他班组" in parsed_res:
                            teams_update["汪佩沾其他班组"] = parsed_res["汪佩沾其他班组"]
                        if "王宜强施工班组" in parsed_res:
                            teams_update["王宜强施工班组"] = parsed_res["王宜强施工班组"]
                            
                        updates["teams"] = teams_update
                        saved_data[selected_month].update(updates)
                        _save_monthly_data(saved_data)
                        
                        # 清理 session state 强制刷新表单
                        for k in list(st.session_state.keys()):
                            if k.startswith("hist_") and selected_month in k:
                                del st.session_state[k]
                        
                        st.success("🎉 解析并保存成功！已自动录入人数，请在左侧【表单】确认结果。")
                        st.rerun()
                    
    # 动态月份专属：人员进离场名单明细
    if not is_historical:
        st.markdown("<div style='margin-top:32px'></div>", unsafe_allow_html=True)
        st.markdown("#### 📋 本月人员变化明细")
        t1, t2 = st.tabs(["本月进场人员", "本月离场归档人员"])
        
        with t1:
            in_list = month_data.get("in_list", [])
            if not in_list:
                st.info("本月暂无进场记录。")
            else:
                df_in = pd.DataFrame([
                    {
                        "姓名": r.get("姓名", ""), 
                        "班组": r.get("班组", ""), 
                        "进场日期": _format_date(r.get("进场日期") or r.get("进场时间") or r.get("入场日期"))
                    }
                    for r in in_list
                ])
                st.dataframe(df_in, use_container_width=True, hide_index=True)
                
        with t2:
            out_list = month_data.get("out_list", [])
            if not out_list:
                st.info("本月暂无完成归档的离场记录。正在离场结算中的人员不计入此处。")
            else:
                df_out = pd.DataFrame([
                    {
                        "姓名": r.get("姓名", ""), 
                        "班组": r.get("班组", ""), 
                        "离场日期": _format_date(r.get("离场日期", ""))
                    }
                    for r in out_list
                ])
                st.dataframe(df_out, use_container_width=True, hide_index=True)
