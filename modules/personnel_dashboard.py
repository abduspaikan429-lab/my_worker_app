# modules/personnel_dashboard.py
# 人员变更面板 - 四源实时联动
from __future__ import annotations
import json, os, re
from datetime import date, datetime
import pandas as pd
import streamlit as st
from modules.master_data import load_master_df
from modules.offboarding_pipeline import load_offboarding_history, offboarding_service

ONBOARDING_FILE = "data/onboarding_data.json"
OFFBOARDING_FILE = "data/offboarding_data.json"
REPORT_CACHE_FILE = "data/personnel_change_report.json"

def _load_json_list(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def _load_json_dict(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def load_onboarding_active():
    return _load_json_dict(ONBOARDING_FILE)

def load_offboarding_active():
    return _load_json_dict(OFFBOARDING_FILE)

def _parse_date(val):
    if not val:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except ValueError:
            continue
    if str(val).isdigit() and len(str(val)) == 5:
        try:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))).date()
        except Exception:
            pass
    return None

def compute_dashboard():
    master = load_master_df()
    onboarding = load_onboarding_active()
    offboarding = load_offboarding_active()
    history = load_offboarding_history()
    
    # 只要人员进入离场结算板块或历史归档，即自动从在场人员中剔除
    on_site_df = offboarding_service.filter_onsite_df(master)
    
    return {
        "on_site_count": len(on_site_df),
        "onboarding_in_progress": len(onboarding),
        "offboarding_in_progress": len(offboarding),
        "on_site_df": on_site_df,
        "monthly_summary": _build_monthly_summary(master, history, offboarding),
        "company_summary": _build_company_summary(on_site_df, onboarding, offboarding),
    }

def _build_monthly_summary(master, history, offboarding=None):
    date_col = next((c for c in ["进场日期","进场时间","入场日期"] if not master.empty and c in master.columns), None)
    onboard_dates = []
    if date_col:
        for v in master[date_col].dropna():
            d = _parse_date(str(v))
            if d:
                onboard_dates.append(d)
    offboard_dates = []
    for rec in history:
        d = _parse_date(rec.get("离场日期",""))
        if d:
            offboard_dates.append(d)
    if offboarding:
        for _, rec in offboarding.items():
            info = rec.get("info", {}) if isinstance(rec, dict) else {}
            d = _parse_date(info.get("离场日期") or rec.get("created_at") or str(date.today()))
            if d:
                offboard_dates.append(d)
    all_dates = onboard_dates + offboard_dates
    if not all_dates:
        return pd.DataFrame(columns=["月份","进场人数","离场人数","累计在场"])
    min_m = min(all_dates).replace(day=1)
    max_m = max(all_dates).replace(day=1)
    months = []
    cur = min_m
    while cur <= max_m:
        months.append(cur.strftime("%Y-%m"))
        cur = cur.replace(month=cur.month+1) if cur.month < 12 else cur.replace(year=cur.year+1,month=1)
    rows = []
    cumulative = 0
    for m in months:
        inn = sum(1 for d in onboard_dates if d.strftime("%Y-%m") == m)
        out = sum(1 for d in offboard_dates if d.strftime("%Y-%m") == m)
        cumulative = max(cumulative + inn - out, 0)
        rows.append({"月份": m, "进场人数": inn, "离场人数": out, "累计在场": cumulative})
    return pd.DataFrame(rows)

def _build_company_summary(on_site_df, onboarding, offboarding):
    team_col = next((c for c in ["班组","工种","人员类型"] if not on_site_df.empty and c in on_site_df.columns), None)
    teams = {}
    
    def _map_company(t_name):
        name = str(t_name).strip()
        if "王宜强" in name or "旭之升" in name:
            return "江苏旭之升"
        elif "汪佩沾" in name or "青海久昌" in name:
            return "青海久昌"
        elif not name or name in ("nan", "None", "未分组"):
            return "未分组"
        else:
            return "其他"

    if team_col and not on_site_df.empty:
        for team, grp in on_site_df.groupby(team_col):
            t = _map_company(team)
            teams.setdefault(t, {"在场":0,"进场手续中":0,"离场结算中":0})
            teams[t]["在场"] += len(grp)
    for _, wd in onboarding.items():
        team = _map_company(wd.get("info",{}).get("班组",""))
        teams.setdefault(team, {"在场":0,"进场手续中":0,"离场结算中":0})
        teams[team]["进场手续中"] += 1
    for _, wd in offboarding.items():
        team = _map_company(wd.get("info",{}).get("班组",""))
        teams.setdefault(team, {"在场":0,"进场手续中":0,"离场结算中":0})
        teams[team]["离场结算中"] += 1
    if not teams:
        return pd.DataFrame(columns=["公司","在场","进场手续中","离场结算中"])
    return pd.DataFrame([{"公司": t, **v} for t,v in sorted(teams.items())])

def parse_change_report_excel(file):
    P_MONTH = re.compile(r"(\d+)月")
    P_IN = re.compile(r"进场[务工人员]*总数[：:]\s*(\d+)")
    P_OUT = re.compile(r"离场[务工人员]*总数[：:]\s*(\d+)")
    P_ON = re.compile(r"现场[务工人员]*总数[：:]\s*(\d+)")
    results = []
    try:
        xl = pd.ExcelFile(file, engine="openpyxl")
    except Exception:
        return []
    yr = date.today().year
    for sn in xl.sheet_names:
        try:
            raw = xl.parse(sn, header=None, nrows=6, dtype=str)
        except Exception:
            continue
        mm = P_MONTH.search(sn)
        if not mm:
            continue
        month_str = f"{yr}-{int(mm.group(1)):02d}"
        is_total = "总" in sn
        co = sn.replace(mm.group(0),"").strip()
        co = co if co not in ("","总") else "全部（汇总）"
        stat = ""
        for ri in range(min(6, len(raw))):
            t = " ".join(str(v) for v in raw.iloc[ri].tolist() if pd.notna(v) and str(v) not in ("nan","None",""))
            if "进场" in t and "离场" in t:
                stat = t
                break
        if not stat:
            continue
        mi = P_IN.search(stat); mo = P_OUT.search(stat); mn = P_ON.search(stat)
        results.append({
            "月份": month_str, "Sheet名": sn, "班组/公司": co,
            "类型": "汇总" if is_total else "分表",
            "进场人数": int(mi.group(1)) if mi else 0,
            "离场人数": int(mo.group(1)) if mo else 0,
            "在场人数": int(mn.group(1)) if mn else 0,
        })
    return results

def load_report_cache():
    return _load_json_list(REPORT_CACHE_FILE)

def save_report_cache(data):
    os.makedirs("data", exist_ok=True)
    with open(REPORT_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def merge_report_cache(new_records):
    existing = load_report_cache()
    km = {(r["月份"],r["Sheet名"]): r for r in existing}
    for rec in new_records:
        km[(rec["月份"],rec["Sheet名"])] = rec
    merged = sorted(km.values(), key=lambda r:(r["月份"],r.get("类型",""),r["Sheet名"]))
    save_report_cache(merged)
    return merged

def _kpi_card(label, value, color="#6366F1"):
    st.markdown(f"""<div class="kpi-card" style="border-left:4px solid {color};">
        <div class="kpi-value" style="color:{color};">{value}</div>
        <div class="kpi-label">{label}</div></div>""", unsafe_allow_html=True)

def _render_overview(data):
    c1,c2,c3,c4 = st.columns(4)
    with c1: _kpi_card("当前在场总人数", data["on_site_count"], "#6366F1")
    with c2: _kpi_card("进场手续办理中", data["onboarding_in_progress"], "#34D399")
    with c3: _kpi_card("离场结算办理中", data["offboarding_in_progress"], "#F87171")
    with c4: _kpi_card("含手续中总在场", data["on_site_count"]+data["onboarding_in_progress"], "#FB923C")
    st.markdown("<div style='margin-top:24px'></div>", unsafe_allow_html=True)
    monthly = data["monthly_summary"]
    if monthly.empty:
        st.info("暂无月度趋势数据。\n\n进场数据来自档案魔法整合中的「进场日期」字段；离场数据来自离场流水线完成5步结算后的自动归档。")
        return
    st.markdown("#### 📈 月度人员变动趋势")
    try:
        import plotly.graph_objects as go
        fig = go.Figure()
        for col,color,dash in [("进场人数","#34D399","solid"),("离场人数","#F87171","solid"),("累计在场","#6366F1","dot")]:
            fig.add_trace(go.Scatter(x=monthly["月份"],y=monthly[col],mode="lines+markers",name=col,
                line=dict(color=color,width=3,dash=dash),marker=dict(size=8)))
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter,sans-serif",size=13),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
            margin=dict(l=0,r=0,t=10,b=0),hovermode="x unified",
            yaxis=dict(gridcolor="rgba(100,116,139,0.15)"),xaxis=dict(gridcolor="rgba(100,116,139,0.08)"))
        st.plotly_chart(fig, use_container_width=True)
    except ImportError:
        st.line_chart(monthly.set_index("月份")[["进场人数","离场人数","累计在场"]])
    st.markdown("#### 📋 月度明细")
    st.dataframe(monthly.rename(columns={"累计在场":"月末在场人数"}), use_container_width=True, hide_index=True)

def _render_company(data):
    company = data["company_summary"]
    on_site_df = data["on_site_df"]
    if company.empty:
        st.info("暂无公司数据，请先在「档案魔法整合」中同步人员信息。")
        return
    st.markdown("#### 🏢 各公司人员分布")
    try:
        import plotly.graph_objects as go
        fig = go.Figure()
        for col,color in [("在场","#6366F1"),("进场手续中","#34D399"),("离场结算中","#F87171")]:
            if col in company.columns:
                fig.add_trace(go.Bar(name=col,x=company["公司"],y=company[col],marker_color=color,text=company[col],textposition="auto"))
        fig.update_layout(barmode="group",plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter,sans-serif",size=13),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
            margin=dict(l=0,r=0,t=10,b=0),yaxis=dict(gridcolor="rgba(100,116,139,0.15)"))
        st.plotly_chart(fig, use_container_width=True)
    except ImportError:
        st.bar_chart(company.set_index("公司")[["在场","进场手续中","离场结算中"]])
    st.markdown("#### 📋 公司汇总")
    st.dataframe(company, use_container_width=True, hide_index=True)
    st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)
    with st.expander("📋 查看在场人员完整名单", expanded=False):
        if on_site_df.empty:
            st.info("暂无在场人员数据。")
        else:
            show_cols = [c for c in ["姓名","工种","班组","进场日期","进场时间","身份证号"] if c in on_site_df.columns]
            search = st.text_input("搜索姓名/工种", key="dash_name_search")
            display = on_site_df[show_cols] if show_cols else on_site_df
            if search:
                mask = pd.Series(False,index=display.index)
                for col in display.columns:
                    mask |= display[col].astype(str).str.contains(search,na=False)
                display = display[mask]
            st.caption(f"共 {len(display)} 人")
            st.dataframe(display, use_container_width=True, hide_index=True, height=400)

def _render_excel_import():
    st.markdown("""<div class="hint-box">📌 此功能用于导入纸质月报Excel，提取进场/离场/在场数字，与系统实时数据进行<b>对比核对</b>。
        月报数据<b>不会覆盖</b>系统档案，仅作辅助记录。</div>""", unsafe_allow_html=True)
    st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("上传人员变更月报Excel（支持多Sheet）", type=["xlsx","xls"], key="dash_excel_upload")
    if uploaded:
        with st.spinner("正在解析Excel..."):
            records = parse_change_report_excel(uploaded)
        if not records:
            st.error("未能从Excel中提取到有效数据，请检查文件格式。")
            return
        st.success(f"解析成功：共识别 {len(records)} 个Sheet的数据。")
        st.dataframe(pd.DataFrame(records), use_container_width=True, hide_index=True)
        if st.button("✅ 确认保存到系统缓存", type="primary", key="dash_save_report"):
            merged = merge_report_cache(records)
            st.success(f"已保存！当前缓存共 {len(merged)} 条月报记录。")
            st.rerun()
    cached = load_report_cache()
    if cached:
        st.markdown("---")
        st.markdown("#### 📦 已缓存的月报数据")
        cached_df = pd.DataFrame(cached)
        total_df = cached_df[cached_df["类型"]=="汇总"].copy() if "类型" in cached_df.columns else cached_df
        if not total_df.empty:
            st.markdown("##### 月度汇总行（来自月报Excel）")
            st.dataframe(total_df, use_container_width=True, hide_index=True)
        with st.expander("查看全部 Sheet 明细"):
            st.dataframe(cached_df, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.markdown("#### 🔍 月报 vs 系统实时数据对比")
        master = load_master_df()
        history = load_offboarding_history()
        offboarding = load_offboarding_active()
        if total_df.empty or master.empty:
            st.info("系统档案或月报汇总数据不足，无法对比。")
            return
        date_col = next((c for c in ["进场日期","进场时间","入场日期"] if c in master.columns), None)
        compare_rows = []
        for _, row in total_df.iterrows():
            m = row["月份"]
            sys_in = sum(1 for v in master[date_col].dropna() if str(v)[:7]==m) if date_col else 0
            hist_out = sum(1 for rec in history if str(rec.get("离场日期",""))[:7]==m)
            active_out = sum(
                1 for _, rec in offboarding.items()
                if str(rec.get("info", {}).get("离场日期") or rec.get("created_at") or str(date.today()))[:7] == m
            )
            sys_out = hist_out + active_out
            compare_rows.append({"月份":m,"月报进场":int(row.get("进场人数",0)),"系统进场":sys_in,
                "进场差异":sys_in-int(row.get("进场人数",0)),"月报离场":int(row.get("离场人数",0)),
                "系统离场":sys_out,"离场差异":sys_out-int(row.get("离场人数",0))})
        compare_df = pd.DataFrame(compare_rows)
        def _hl(val):
            if isinstance(val,(int,float)) and val!=0:
                return "background-color:#FEF3C7;color:#92400E;font-weight:bold;"
            return ""
        diff_cols = [c for c in compare_df.columns if "差异" in c]
        st.dataframe(compare_df.style.applymap(_hl,subset=diff_cols), use_container_width=True, hide_index=True)
        st.caption("差异=系统数据-月报数据；黄色标注表示数据不一致，请人工核查。")

def render():
    st.markdown("""<style>
    .kpi-card{background:#fff;border-radius:14px;padding:20px 22px 16px;box-shadow:0 2px 12px rgba(0,0,0,0.07);margin-bottom:4px;}
    .kpi-value{font-size:2.2rem;font-weight:700;line-height:1.1;}
    .kpi-label{font-size:.82rem;color:#64748B;margin-top:4px;font-weight:500;}
    </style>""", unsafe_allow_html=True)
    st.markdown("""
    <div class="page-header-deco">
        <span class="material-symbols-outlined" style="font-size:32px;color:#6366F1;">bar_chart</span>
        <div class="header-text"><h2>人员变更面板</h2>
        <p>四源实时联动 · 在场/进场/离场数据一站式总览</p></div>
    </div>
    <div class="color-strip" style="background:linear-gradient(90deg,#C7D2FE,#EDE9FE);"></div>
    """, unsafe_allow_html=True)
    if st.button("🔄 刷新数据", key="dash_refresh"):
        st.rerun()
    data = compute_dashboard()
    tab1,tab2,tab3 = st.tabs(["📊 实时总览","🏢 班组/公司明细","📂 导入月报核对"])
    with tab1: _render_overview(data)
    with tab2: _render_company(data)
    with tab3: _render_excel_import()
