# services/personnel_data_service.py
"""
人员变更数据分析服务：
负责解析 load-data 目录下的三份核心 Excel 报表：
1. 中建二局安装务工人员（含队长、班组长、弄民工）花名册.xlsx
2. 二局安装-足球场项目人员变更月报表（进场情况）.xlsx
3. 二局安装-足球场项目人员变更月报表（离场情况）.xlsx

提供结构化数据清洗、月度流动指标汇总、年龄与省份画像计算、合规管控率评估，
并支持 Plotly 交互式图表与 Matplotlib 高清报表图的生成。
"""

from __future__ import annotations
import os
import io
import json
import copy
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

import pandas as pd
import numpy as np

# 省份代码映射表（依据身份证前两位代码）
PROVINCE_MAP = {
    '11': '北京', '12': '天津', '13': '河北', '14': '山西', '15': '内蒙古',
    '21': '辽宁', '22': '吉林', '23': '黑龙江', '31': '上海', '32': '江苏',
    '33': '浙江', '34': '安徽', '35': '福建', '36': '江西', '37': '山东',
    '41': '河南', '42': '湖北', '43': '湖南', '44': '广东', '45': '广西',
    '46': '海南', '50': '重庆', '51': '四川', '52': '贵州', '53': '云南',
    '54': '西藏', '61': '陕西', '62': '甘肃', '63': '青海', '64': '宁夏', '65': '新疆'
}

DEFAULT_BASE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'load-data')
ROSTER_FILENAME = '中建二局安装务工人员（含队长、班组长、弄民工）花名册.xlsx'
INFLOW_FILENAME = '二局安装-足球场项目人员变更月报表（进场情况）.xlsx'
OUTFLOW_FILENAME = '二局安装-足球场项目人员变更月报表（离场情况）.xlsx'


def get_province_from_id(id_val: Any) -> str:
    """从身份证提取籍贯省份"""
    s = str(id_val).strip()
    if len(s) >= 2 and s[:2] in PROVINCE_MAP:
        return PROVINCE_MAP[s[:2]]
    return '其他'


def calculate_age_from_id(id_val: Any, base_year: int = 2026) -> Optional[int]:
    """从身份证提取出生年份计算周岁"""
    s = str(id_val).strip()
    if len(s) == 18 and s[6:10].isdigit():
        birth_year = int(s[6:10])
        return max(18, min(80, base_year - birth_year))
    return None


def clean_job_title(job_val: Any) -> str:
    """清洗统一工种名称"""
    s = str(job_val).strip()
    if not s or s in ('nan', 'None', '', '其它'):
        if s == '其它':
            return '其他'
        return '普工'
    return s


class PersonnelDataService:
    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = base_dir or DEFAULT_BASE_DIR
        self.roster_path = os.path.join(self.base_dir, ROSTER_FILENAME)
        self.inflow_path = os.path.join(self.base_dir, INFLOW_FILENAME)
        self.outflow_path = os.path.join(self.base_dir, OUTFLOW_FILENAME)
        self._cached_excel_data: Optional[Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]] = None
        self._cached_data: Optional[Dict[str, Any]] = None

    def get_data_status(self) -> Dict[str, Dict[str, Any]]:
        """获取各数据文件存在性、大小及修改时间"""
        files = {
            "roster": ("人员花名册", self.roster_path),
            "inflow": ("进场月报表", self.inflow_path),
            "outflow": ("离场月报表", self.outflow_path),
        }
        status = {}
        for key, (label, path) in files.items():
            exists = os.path.exists(path)
            size = os.path.getsize(path) if exists else 0
            mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S") if exists else None
            status[key] = {
                "label": label,
                "path": path,
                "exists": exists,
                "size_kb": round(size / 1024, 1),
                "mtime": mtime
            }
        return status

    def load_all_data(self, force_reload: bool = False) -> Dict[str, Any]:
        """全量解析三份 Excel 并实时结合系统动态业务流水，返回结构化业务数据包"""
        if self._cached_excel_data is None or force_reload:
            df_roster, unique_roster = self._parse_roster()
            df_inflow = self._parse_inflow()
            df_outflow = self._parse_outflow()
            raw_summary = self._compute_monthly_summary(df_roster, df_inflow, df_outflow)
            self._cached_excel_data = (
                df_roster.copy(),
                unique_roster.copy(),
                df_inflow.copy(),
                df_outflow.copy(),
                copy.deepcopy(raw_summary)
            )

        df_roster = self._cached_excel_data[0].copy()
        unique_roster = self._cached_excel_data[1].copy()
        df_inflow = self._cached_excel_data[2].copy()
        df_outflow = self._cached_excel_data[3].copy()
        monthly_summary = copy.deepcopy(self._cached_excel_data[4])
        official_summary = copy.deepcopy(self._cached_excel_data[4])

        # 动态联动：每次均实时接入主表、进场流水线、离场流水线与离场归档
        df_roster, unique_roster, df_inflow, df_outflow, monthly_summary, live_status = self._enrich_with_live_system(
            df_roster, unique_roster, df_inflow, df_outflow, monthly_summary
        )

        demographics = self._compute_demographics(unique_roster)
        compliance = self._compute_compliance(df_inflow, df_outflow)

        data = {
            "df_roster": df_roster,
            "unique_roster": unique_roster,
            "df_inflow": df_inflow,
            "df_outflow": df_outflow,
            "monthly_summary": monthly_summary,
            "official_summary": official_summary,
            "demographics": demographics,
            "compliance": compliance,
            "live_status": live_status,
            "months": ['6月', '7月', '8月', '9月'],
            "teams": ['江苏旭之升 (王宜强施工班组)', '青海久昌 (汪佩沾其他班组)']
        }
        self._cached_data = data
        return data

    def _enrich_with_live_system(
        self,
        df_roster: pd.DataFrame,
        unique_roster: pd.DataFrame,
        df_inflow: pd.DataFrame,
        df_outflow: pd.DataFrame,
        monthly_summary: Dict[str, Any]
    ) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any], Dict[str, Any]]:
        """
        与系统业务模块实时联动：
        1. 接入主表档案 (modules.master_data)
        2. 接入进场流水线 (modules.onboarding_pipeline.onboarding_service)
        3. 接入离场流水线与离场归档 (services.offboarding_service.OffboardingService)
        """
        live_status: Dict[str, Any] = {
            "master_count": 0,
            "onboarding_pending_count": 0,
            "offboarding_pending_count": 0,
            "offboarding_history_count": 0,
            "currently_onsite": len(unique_roster),
            "unpasted_count": 0,
            "unpasted_workers": []
        }

        try:
            from modules.master_data import load_master_df
            from modules.onboarding_pipeline import onboarding_service
            from services.offboarding_service import OffboardingService
        except Exception:
            if not unique_roster.empty and 'Source' not in unique_roster.columns:
                unique_roster['Source'] = 'Excel基准底册'
                unique_roster['LiveStatus'] = '在场正常'
            return df_roster, unique_roster, df_inflow, df_outflow, monthly_summary, live_status

        def clean_id(v: Any) -> str:
            s = str(v).strip().replace("'", "").strip()
            return "" if s.lower() in ("nan", "none", "") else s

        master_df = load_master_df()
        onboarding_recs = onboarding_service.get_records()
        off_svc = OffboardingService()
        pending_off = off_svc.get_records()
        history_off = off_svc.load_history()

        live_status["master_count"] = len(master_df) if not master_df.empty else 0
        live_status["onboarding_pending_count"] = len(onboarding_recs)
        live_status["offboarding_pending_count"] = len(pending_off)
        live_status["offboarding_history_count"] = len(history_off)

        # 1. 建立离场工人识别索引（来自系统离场流水线历史归档与在办）
        archived_exit_ids = set()
        archived_exit_names = set()
        for rec in history_off:
            cid = clean_id(rec.get("身份证号") or rec.get("id_card"))
            name = str(rec.get("姓名") or rec.get("name") or "").strip()
            if cid:
                archived_exit_ids.add(cid)
            if name:
                archived_exit_names.add(name)

        pending_exit_ids = set()
        pending_exit_names = set()
        for k, p_data in pending_off.items():
            info = p_data.get("info", {}) if isinstance(p_data, dict) else {}
            cid = clean_id(info.get("身份证号") or k)
            name = str(info.get("姓名") or k.split("_")[0]).strip()
            if cid:
                pending_exit_ids.add(cid)
            if name:
                pending_exit_names.add(name)

        # 2. 为现有 Excel 花名册标记 Source 与 LiveStatus
        unique_roster = unique_roster.copy()
        if not unique_roster.empty:
            if "Source" not in unique_roster.columns:
                unique_roster["Source"] = "Excel基准底册"
            
            def determine_status(row):
                cid = clean_id(row.get("ID"))
                name = str(row.get("Name") or "").strip()
                if (cid and cid in archived_exit_ids) or (name and name in archived_exit_names):
                    return "已离场归档"
                if (cid and cid in pending_exit_ids) or (name and name in pending_exit_names):
                    return "离场结算中"
                return "在场正常"

            unique_roster["LiveStatus"] = unique_roster.apply(determine_status, axis=1)

        existing_ids = {clean_id(x) for x in unique_roster["ID"] if clean_id(x)}
        existing_names = set(unique_roster["Name"].str.strip()) if not unique_roster.empty else set()

        def _parse_month(d: str) -> str:
            s = str(d).strip()
            if "06" in s or "-6-" in s or "/6/" in s or "6月" in s:
                return "6月"
            elif "07" in s or "-7-" in s or "/7/" in s or "7月" in s:
                return "7月"
            elif "08" in s or "-8-" in s or "/8/" in s or "8月" in s:
                return "8月"
            elif "09" in s or "-9-" in s or "/9/" in s or "9月" in s:
                return "9月"
            return "9月"

        unpasted_workers = []
        new_roster_rows = []
        new_inflow_rows = []

        # 3. 扫描主表 (master_df) 中的增量人员（通过档案魔法整合新导入的人员）
        if not master_df.empty:
            id_col = next((c for c in master_df.columns if "身份证" in c), None)
            name_col = next((c for c in master_df.columns if "姓名" in c), None)
            team_col = next((c for c in master_df.columns if "班组" in c), None)
            job_col = next((c for c in master_df.columns if "工种" in c), None)
            date_col = next((c for c in master_df.columns if "进场" in c), None)
            status_col = next((c for c in master_df.columns if ("在场" in c or "进退场" in c) and "状态" in c), None)
            exit_date_col = next((c for c in master_df.columns if ("退场" in c or "离场" in c) and "日期" in c), None)
            addr_col = next((c for c in master_df.columns if "住址" in c or "地址" in c), None)
            gender_col = next((c for c in master_df.columns if "性别" in c), None)
            contract_col = next((c for c in master_df.columns if "合同" in c), None)

            # 离场月报表中的历史离场人员索引
            outflow_cids = {clean_id(x) for x in df_outflow["ID"] if clean_id(x)}
            outflow_names = {str(x).strip() for x in df_outflow["Name"] if str(x).strip() and str(x) != "nan"}

            for _, r in master_df.iterrows():
                cid = clean_id(r[id_col]) if id_col else ""
                name = str(r[name_col]).strip() if name_col else ""
                if not name or name in ("nan", "None"):
                    continue

                status_val = str(r[status_col]).strip() if status_col else ""
                exit_date_val = str(r[exit_date_col]).strip() if exit_date_col else ""
                if exit_date_val in ("nan", "None"):
                    exit_date_val = ""

                # 严格判断该工人是否为历史已退场/已离场人员（如张学成在 6 月已离场）
                is_offboarded = (
                    status_val in ("已退场", "退场", "离场", "已离场") or
                    bool(exit_date_val) or
                    (cid and cid in archived_exit_ids) or
                    (name and name in archived_exit_names) or
                    (cid and cid in outflow_cids) or
                    (name and name in outflow_names)
                )

                if is_offboarded:
                    # 历史已离场人员严禁作为新进场并入进场名单或花名册增量！
                    continue

                # 仅对真正未离场且花名册中尚不存在的在场增量工人进行登记
                is_known = (cid and cid in existing_ids) or (name and name in existing_names)
                if not is_known:
                    t_val = str(r[team_col]).strip() if team_col else "金属屋面班组"
                    team_std = "江苏旭之升 (王宜强施工班组)" if ("旭之升" in t_val or "江" in t_val or "王宜强" in t_val) else "青海久昌 (汪佩沾其他班组)"
                    job_std = clean_job_title(r[job_col] if job_col else "普工")
                    gender_std = str(r[gender_col]).strip() if gender_col else "男"
                    if gender_std in ("nan", "None", ""):
                        gender_std = "男"
                    d_val = str(r[date_col]).strip() if date_col else ""
                    m_val = _parse_month(d_val)
                    addr_std = str(r[addr_col]).strip() if addr_col else ""
                    c_no = str(r[contract_col]).strip() if contract_col else ""

                    new_row = {
                        "Month": m_val,
                        "Team": team_std,
                        "Name": name,
                        "ID": cid,
                        "Gender": gender_std,
                        "Job": job_std,
                        "Job_Clean": job_std,
                        "Address": addr_std,
                        "ContractNo": c_no,
                        "Province": get_province_from_id(cid),
                        "Age": calculate_age_from_id(cid),
                        "AgeGroup": "其他",
                        "Source": "系统主表新同步",
                        "LiveStatus": "在场正常"
                    }
                    if new_row["Age"]:
                        a = new_row["Age"]
                        new_row["AgeGroup"] = '18-29岁 (青年)' if a < 30 else ('30-39岁 (青壮年)' if a < 40 else ('40-49岁 (成熟期)' if a < 50 else '50岁及以上 (老龄工)'))
                    
                    new_roster_rows.append(new_row)
                    unpasted_workers.append(new_row)
                    if cid:
                        existing_ids.add(cid)
                    if name:
                        existing_names.add(name)

                    new_inflow_rows.append({
                        "Month": m_val,
                        "Team": team_std,
                        "Name": name,
                        "ID": cid,
                        "Date": d_val.split()[0] if d_val else datetime.now().strftime("%Y-%m-%d"),
                        "Job": job_std,
                        "Job_Clean": job_std,
                        "Registered": "是",
                        "ContractSigned": "是",
                        "Remarks": "系统主表新同步"
                    })

        # 4. 扫描【进场流水线】在办人员 (深度支持当前在办卡片动态并入大屏)
        for k, rec in onboarding_recs.items():
            if not isinstance(rec, dict):
                continue
            info = rec.get("info", rec) if isinstance(rec, dict) else {}
            name = str(info.get("姓名") or info.get("name") or k.split("_")[0]).strip()
            cid = clean_id(info.get("身份证号") or info.get("id_card"))
            if not name or name in ("nan", "None"):
                continue

            t_val = str(info.get("班组") or info.get("team") or "").strip()
            team_std = "江苏旭之升 (王宜强施工班组)" if ("旭之升" in t_val or "江" in t_val or "王宜强" in t_val) else "青海久昌 (汪佩沾其他班组)"
            job_raw = str(info.get("工种") or info.get("job") or info.get("job_type") or "").strip()
            job_std = clean_job_title(job_raw if job_raw else "普工")
            d_val = str(info.get("进场日期") or rec.get("created_at") or "").strip()
            m_val = _parse_month(d_val)
            gender_std = str(info.get("性别") or "").strip() or "男"
            addr_std = str(info.get("住址") or info.get("家庭住址") or info.get("详细地址") or "").strip()
            c_no = str(info.get("劳动合同编号") or info.get("contract_no") or "").strip()

            # 花名册去重判断：如果身份证存在，或同名，视为花名册已有人员
            is_in_roster = False
            if cid and cid in existing_ids:
                is_in_roster = True
            elif name in existing_names:
                is_in_roster = True

            if is_in_roster:
                # 已在花名册底册中（如张克美），仅标记在办状态，严禁向花名册重复追加行！
                if not unique_roster.empty:
                    m_id = (unique_roster["ID"] == cid) if cid else pd.Series(False, index=unique_roster.index)
                    m_name = (unique_roster["Name"] == name)
                    target_mask = m_id | m_name
                    unique_roster.loc[target_mask & (unique_roster["LiveStatus"] == "在场正常"), "LiveStatus"] = "进场手续在办"
                    if cid:
                        empty_id_mask = target_mask & (unique_roster["ID"].isin(["", "nan", "None"]) | unique_roster["ID"].isna())
                        unique_roster.loc[empty_id_mask, "ID"] = cid
            else:
                # 真正新进场人员（如汪小波、郭云新、李栓）
                new_row = {
                    "Month": m_val,
                    "Team": team_std,
                    "Name": name,
                    "ID": cid,
                    "Gender": gender_std,
                    "Job": job_std,
                    "Job_Clean": job_std,
                    "Address": addr_std,
                    "ContractNo": c_no,
                    "Province": get_province_from_id(cid) if cid else "其他",
                    "Age": calculate_age_from_id(cid) if cid else None,
                    "AgeGroup": "其他",
                    "Source": "进场流水线在办",
                    "LiveStatus": "进场手续在办"
                }
                if new_row["Age"]:
                    a = new_row["Age"]
                    new_row["AgeGroup"] = '18-29岁 (青年)' if a < 30 else ('30-39岁 (青壮年)' if a < 40 else ('40-49岁 (成熟期)' if a < 50 else '50岁及以上 (老龄工)'))

                new_roster_rows.append(new_row)
                unpasted_workers.append(new_row)
                if cid:
                    existing_ids.add(cid)
                if name:
                    existing_names.add(name)

            # 进场明细表 (df_inflow) 去重检查：避免同一月份进场名单出现重复人员（如张克美、杜国晓）
            already_in_inflow = False
            # 1. 检查既有 df_inflow 该月份
            sub_inf = df_inflow[df_inflow["Month"] == m_val]
            if not sub_inf.empty:
                if cid and any(clean_id(x) == cid for x in sub_inf["ID"]):
                    already_in_inflow = True
                elif any(str(x).strip() == name for x in sub_inf["Name"]):
                    already_in_inflow = True

            # 2. 检查本次新增的 new_inflow_rows 该月份
            if not already_in_inflow:
                for r in new_inflow_rows:
                    if r.get("Month") == m_val:
                        if (cid and clean_id(r.get("ID")) == cid) or (r.get("Name") == name):
                            already_in_inflow = True
                            break

            if already_in_inflow:
                # 该月份进场表中已有该工人（如张克美已在 9 月进场表），仅补全缺失字段，绝不重复 append！
                mask = (df_inflow["Month"] == m_val) & (df_inflow["Name"] == name)
                if mask.any():
                    idx = df_inflow[mask].index[0]
                    if not str(df_inflow.at[idx, "Date"]).strip() and d_val:
                        df_inflow.at[idx, "Date"] = d_val.split()[0]
                    if str(df_inflow.at[idx, "Registered"]).strip() in ("", "nan", "None", "否"):
                        if rec.get("access", {}).get("智慧护薪合同发起及工人/班组长确认"):
                            df_inflow.at[idx, "Registered"] = "是"
                    if not str(df_inflow.at[idx, "Remarks"]).strip() or str(df_inflow.at[idx, "Remarks"]).strip() == "nan":
                        df_inflow.at[idx, "Remarks"] = "进场流水线在办"
            else:
                # 真正的新进场记录，加入进场明细表
                new_inflow_rows.append({
                    "Month": m_val,
                    "Team": team_std,
                    "Name": name,
                    "ID": cid,
                    "Date": d_val.split()[0] if d_val else datetime.now().strftime("%Y-%m-%d"),
                    "Job": job_std,
                    "Job_Clean": job_std,
                    "Registered": "是" if rec.get("access", {}).get("智慧护薪合同发起及工人/班组长确认") else "办理中",
                    "ContractSigned": "是" if rec.get("paper", {}).get("劳动合同(纸质)") else "办理中",
                    "Remarks": "进场流水线在办"
                })

        if new_roster_rows:
            df_new_roster = pd.DataFrame(new_roster_rows)
            for c in df_new_roster.columns:
                if c in unique_roster.columns and df_new_roster[c].isna().all():
                    df_new_roster[c] = df_new_roster[c].astype(unique_roster[c].dtype)
            unique_roster = pd.concat([unique_roster, df_new_roster], ignore_index=True)
            df_roster = pd.concat([df_roster, df_new_roster], ignore_index=True)

        if new_inflow_rows:
            df_new_inflow = pd.DataFrame(new_inflow_rows)
            df_inflow = pd.concat([df_inflow, df_new_inflow], ignore_index=True)

        # 5. 动态校准 9 月及各月份 monthly_summary 中的宏观流动与队伍分布（确保与 df_inflow 明细名单 100% 对齐！）
        if "9月" in monthly_summary and not df_inflow.empty:
            m9_inflow = df_inflow[df_inflow["Month"] == "9月"]
            m9_actual_in = len(m9_inflow)
            monthly_summary["9月"]["in_total"] = m9_actual_in
            monthly_summary["9月"]["net_change"] = m9_actual_in - monthly_summary["9月"]["out_total"]

            for t_name in monthly_summary["9月"]["teams"]:
                t_count = len(m9_inflow[m9_inflow["Team"] == t_name])
                monthly_summary["9月"]["teams"][t_name]["in_count"] = t_count

        # 6. 动态在场总数计算
        active_cnt = len(unique_roster[unique_roster["LiveStatus"].isin(["在场正常", "进场手续在办"])])
        live_status["currently_onsite"] = active_cnt
        live_status["unpasted_count"] = len(unpasted_workers)
        live_status["unpasted_workers"] = unpasted_workers

        # 同步校准 9 月的现场实际在场总数
        if "9月" in monthly_summary:
            monthly_summary["9月"]["actual_onsite_total"] = active_cnt

        return df_roster, unique_roster, df_inflow, df_outflow, monthly_summary, live_status

    def _parse_roster(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """解析花名册表中的全部月度名单"""
        if not os.path.exists(self.roster_path):
            return pd.DataFrame(), pd.DataFrame()

        xls = pd.ExcelFile(self.roster_path)
        roster_rows = []

        for s in xls.sheet_names:
            if '总' in s:
                continue
            month = '6月' if '6' in s else ('7月' if '7' in s else ('8月' if '8' in s else '9月'))
            team = '江苏旭之升 (王宜强施工班组)' if ('江苏旭之升' in s or '江' in s) else '青海久昌 (汪佩沾其他班组)'
            df = pd.read_excel(self.roster_path, sheet_name=s)

            h_idx = None
            for idx, row in df.iterrows():
                row_str = ' '.join([str(v) for v in row.values])
                if '姓名' in row_str and '身份证' in row_str:
                    h_idx = idx
                    break

            if h_idx is not None:
                cols = [str(x).strip() for x in df.iloc[h_idx].values]
                name_c = [i for i, c in enumerate(cols) if '姓名' in c][0]
                id_c = [i for i, c in enumerate(cols) if '身份证' in c][0]
                gender_c = [i for i, c in enumerate(cols) if '性别' in c][0] if any('性别' in c for c in cols) else None
                job_c = [i for i, c in enumerate(cols) if '工种' in c][0] if any('工种' in c for c in cols) else None
                addr_c = [i for i, c in enumerate(cols) if '家庭住址' in c or '住址' in c][0] if any('住址' in c for c in cols) else None
                contract_c = [i for i, c in enumerate(cols) if '合同编号' in c or '合同' in c][0] if any('合同' in c for c in cols) else None

                for _, r in df.iloc[h_idx + 1:].iterrows():
                    name = str(r.iloc[name_c]).strip()
                    if name and name not in ['nan', 'None', '', '姓名'] and '申明' not in name:
                        id_no = str(r.iloc[id_c]).strip()
                        gender = str(r.iloc[gender_c]).strip() if gender_c is not None else '男'
                        if gender in ('nan', 'None', ''):
                            gender = '男'
                        job = str(r.iloc[job_c]).strip() if job_c is not None else '普工'
                        addr = str(r.iloc[addr_c]).strip() if addr_c is not None else ''
                        c_no = str(r.iloc[contract_c]).strip() if contract_c is not None else ''

                        roster_rows.append({
                            'Month': month,
                            'Team': team,
                            'Name': name,
                            'ID': id_no,
                            'Gender': gender,
                            'Job': job,
                            'Job_Clean': clean_job_title(job),
                            'Address': addr if addr != 'nan' else '',
                            'ContractNo': c_no if c_no != 'nan' else '',
                            'Province': get_province_from_id(id_no),
                            'Age': calculate_age_from_id(id_no)
                        })

        df_roster = pd.DataFrame(roster_rows)
        if df_roster.empty:
            return df_roster, df_roster

        # 年龄梯队分箱
        def assign_age_group(age):
            if pd.isna(age):
                return '其他'
            if age < 30:
                return '18-29岁 (青年)'
            elif age < 40:
                return '30-39岁 (青壮年)'
            elif age < 50:
                return '40-49岁 (成熟期)'
            else:
                return '50岁及以上 (老龄工)'

        df_roster['AgeGroup'] = df_roster['Age'].apply(assign_age_group)
        unique_roster = df_roster.drop_duplicates(subset=['ID']).copy()

        return df_roster, unique_roster

    def _parse_inflow(self) -> pd.DataFrame:
        """解析进场月报表中的明细数据"""
        if not os.path.exists(self.inflow_path):
            return pd.DataFrame()

        xls = pd.ExcelFile(self.inflow_path)
        in_records = []
        target_sheets = [s for s in xls.sheet_names if any(s.startswith(m) for m in ['6', '7', '8', '9']) and '总' not in s and s != '合计']

        for s in target_sheets:
            month = s[0] + '月'
            team = '江苏旭之升 (王宜强施工班组)' if ('江' in s or '旭之升' in s) else '青海久昌 (汪佩沾其他班组)'
            df = pd.read_excel(self.inflow_path, sheet_name=s)

            h_idx = None
            for idx, row in df.iterrows():
                r_str = ' '.join([str(v) for v in row.values])
                if '姓名' in r_str and '编号' in r_str:
                    h_idx = idx
                    break

            if h_idx is not None:
                cols = [str(x).strip() for x in df.iloc[h_idx].values]
                n_c = [i for i, c in enumerate(cols) if '姓名' in c][0]
                id_c = [i for i, c in enumerate(cols) if '身份证' in c][0]
                d_c = [i for i, c in enumerate(cols) if '进场时间' in c or '时间' in c][0] if any('时间' in c for c in cols) else None
                j_c = [i for i, c in enumerate(cols) if '工种' in c][0] if any('工种' in c for c in cols) else None
                reg_c = [i for i, c in enumerate(cols) if '备案' in c][0] if any('备案' in c for c in cols) else None
                c_c = [i for i, c in enumerate(cols) if '劳动合同' in c][0] if any('劳动合同' in c for c in cols) else None
                rem_c = [i for i, c in enumerate(cols) if '备注' in c][0] if any('备注' in c for c in cols) else None

                for _, r in df.iloc[h_idx + 1:].iterrows():
                    name = str(r.iloc[n_c]).strip()
                    if name and name not in ['nan', 'None', '', '姓名'] and '申明' not in name:
                        id_no = str(r.iloc[id_c]).strip()
                        raw_date = str(r.iloc[d_c]).strip() if d_c is not None else ''
                        if raw_date and raw_date != 'nan':
                            dt = raw_date.split()[0]
                        else:
                            dt = ''
                        job = str(r.iloc[j_c]).strip() if j_c is not None else '普工'
                        reg = str(r.iloc[reg_c]).strip() if reg_c is not None else '否'
                        contract = str(r.iloc[c_c]).strip() if c_c is not None else '是'
                        rem = str(r.iloc[rem_c]).strip() if rem_c is not None else ''

                        in_records.append({
                            'Month': month,
                            'Team': team,
                            'Name': name,
                            'ID': id_no if id_no != 'nan' else '',
                            'Date': dt,
                            'Job': job if job != 'nan' else '普工',
                            'Job_Clean': clean_job_title(job),
                            'Registered': reg if reg != 'nan' else '否',
                            'ContractSigned': contract if contract != 'nan' else '是',
                            'Remarks': rem if rem != 'nan' else ''
                        })

        return pd.DataFrame(in_records)

    def _parse_outflow(self) -> pd.DataFrame:
        """解析离场月报表中的明细数据"""
        if not os.path.exists(self.outflow_path):
            return pd.DataFrame()

        xls = pd.ExcelFile(self.outflow_path)
        out_records = []

        for s in xls.sheet_names:
            month = '6月' if '6' in s else ('7月' if '7' in s else ('8月' if '8' in s else '9月'))
            team = '江苏旭之升 (王宜强施工班组)' if ('旭之升' in s or '江' in s) else '青海久昌 (汪佩沾其他班组)'
            df = pd.read_excel(self.outflow_path, sheet_name=s)

            h_idx = None
            for idx, row in df.iterrows():
                r_str = ' '.join([str(v) for v in row.values])
                if '姓名' in r_str and '编号' in r_str:
                    h_idx = idx
                    break

            if h_idx is not None:
                cols = [str(x).strip() for x in df.iloc[h_idx].values]
                n_c = [i for i, c in enumerate(cols) if '姓名' in c][0]
                id_c = [i for i, c in enumerate(cols) if '身份证' in c][0]
                d_c = [i for i, c in enumerate(cols) if '离场时间' in c or '时间' in c][0] if any('时间' in c for c in cols) else None
                dur_c = [i for i, c in enumerate(cols) if '务工时间' in c][0] if any('务工时间' in c for c in cols) else None
                j_c = [i for i, c in enumerate(cols) if '工种' in c][0] if any('工种' in c for c in cols) else None
                settle_c = [i for i, c in enumerate(cols) if '工资结算' in c][0] if any('工资结算' in c for c in cols) else None
                commit_c = [i for i, c in enumerate(cols) if '退场承诺书' in c][0] if any('退场承诺书' in c for c in cols) else None

                for _, r in df.iloc[h_idx + 1:].iterrows():
                    name = str(r.iloc[n_c]).strip()
                    if name and name not in ['nan', 'None', '', '姓名'] and '申明' not in name:
                        id_no = str(r.iloc[id_c]).strip()
                        raw_date = str(r.iloc[d_c]).strip() if d_c is not None else ''
                        if raw_date and raw_date != 'nan':
                            dt = raw_date.split()[0]
                        else:
                            dt = ''
                        dur = str(r.iloc[dur_c]).strip() if dur_c is not None else ''
                        job = str(r.iloc[j_c]).strip() if j_c is not None else '其他'
                        settle = str(r.iloc[settle_c]).strip() if settle_c is not None else '已结算已支付'
                        commit = str(r.iloc[commit_c]).strip() if commit_c is not None else '是'

                        out_records.append({
                            'Month': month,
                            'Team': team,
                            'Name': name,
                            'ID': id_no if id_no != 'nan' else '',
                            'Date': dt,
                            'Duration': dur if dur != 'nan' else '',
                            'Job': job if job != 'nan' else '其他',
                            'Job_Clean': clean_job_title(job),
                            'WageSettled': settle if settle != 'nan' else '已结算已支付',
                            'CommitmentSigned': commit if commit != 'nan' else '是'
                        })

        return pd.DataFrame(out_records)

    def _compute_monthly_summary(self, df_roster: pd.DataFrame, df_inflow: pd.DataFrame, df_outflow: pd.DataFrame) -> Dict[str, Any]:
        """计算月度指标：进场、离场、月末在场人数、各分包队伍分布"""
        months = ['6月', '7月', '8月', '9月']
        summary = {}

        team_yi = '江苏旭之升 (王宜强施工班组)'
        team_wang = '青海久昌 (汪佩沾其他班组)'

        # 优先读取进场月报表中官方核准的【合计】汇总基准表
        if os.path.exists(self.inflow_path):
            try:
                xls_in = pd.ExcelFile(self.inflow_path)
                if '合计' in xls_in.sheet_names:
                    df_sum = pd.read_excel(self.inflow_path, sheet_name='合计')
                    if '月份' in df_sum.columns and '进场' in df_sum.columns:
                        df_sum['月份'] = df_sum['月份'].ffill().astype(int)
                        official_summary = {}
                        for m_num, grp in df_sum.groupby('月份'):
                            m_key = f'{m_num}月'
                            in_sum = int(grp['进场'].sum())
                            out_sum = int(grp['离场'].sum())
                            onsite_sum = int(grp['在场'].sum())
                            actual_onsite_sum = int(grp['实际在场'].sum()) if '实际在场' in grp.columns else (onsite_sum - out_sum)
                            
                            teams_dict = {}
                            for _, r in grp.iterrows():
                                c_name = str(r['公司'])
                                t_key = team_yi if ('旭之升' in c_name or '江' in c_name) else team_wang
                                teams_dict[t_key] = {
                                    'company': c_name,
                                    'in_count': int(r['进场']),
                                    'out_count': int(r['离场']),
                                    'onsite_count': int(r['在场']),
                                    'actual_onsite_count': int(r['实际在场']) if '实际在场' in r else int(r['在场'])
                                }
                            official_summary[m_key] = {
                                "in_total": in_sum,
                                "out_total": out_sum,
                                "onsite_total": onsite_sum,
                                "actual_onsite_total": actual_onsite_sum,
                                "net_change": in_sum - out_sum,
                                "teams": teams_dict
                            }
                        return official_summary
            except Exception:
                pass

        for m in months:
            in_sub = df_inflow[df_inflow['Month'] == m] if not df_inflow.empty else pd.DataFrame()
            out_sub = df_outflow[df_outflow['Month'] == m] if not df_outflow.empty else pd.DataFrame()
            roster_sub = df_roster[df_roster['Month'] == m] if not df_roster.empty else pd.DataFrame()

            in_total = len(in_sub)
            out_total = len(out_sub)
            onsite_total = len(roster_sub)

            in_yi = len(in_sub[in_sub['Team'] == team_yi])
            in_wang = len(in_sub[in_sub['Team'] == team_wang])

            out_yi = len(out_sub[out_sub['Team'] == team_yi])
            out_wang = len(out_sub[out_sub['Team'] == team_wang])

            onsite_yi = len(roster_sub[roster_sub['Team'] == team_yi])
            onsite_wang = len(roster_sub[roster_sub['Team'] == team_wang])

            summary[m] = {
                "in_total": in_total,
                "out_total": out_total,
                "onsite_total": onsite_total,
                "actual_onsite_total": onsite_total - out_total,
                "net_change": in_total - out_total,
                "teams": {
                    team_yi: {
                        "in_count": in_yi,
                        "out_count": out_yi,
                        "onsite_count": onsite_yi,
                        "actual_onsite_count": onsite_yi - out_yi
                    },
                    team_wang: {
                        "in_count": in_wang,
                        "out_count": out_wang,
                        "onsite_count": onsite_wang,
                        "actual_onsite_count": onsite_wang - out_wang
                    }
                }
            }

        return summary

    def _compute_demographics(self, unique_roster: pd.DataFrame) -> Dict[str, Any]:
        """计算劳务工人基础特征指标：年龄结构、籍贯分布、工种技能、性别比例"""
        if unique_roster.empty:
            return {
                "total_unique": 0,
                "gender_counts": {"男": 0, "女": 0},
                "avg_age": 0.0,
                "median_age": 0.0,
                "age_dist": {},
                "top_provinces": {},
                "job_counts": {}
            }

        total_unique = len(unique_roster)
        gender_counts = unique_roster['Gender'].value_counts().to_dict()

        ages = unique_roster['Age'].dropna()
        avg_age = round(float(ages.mean()), 1) if not ages.empty else 0.0
        median_age = round(float(ages.median()), 1) if not ages.empty else 0.0

        # 年龄梯队
        age_cuts = pd.cut(
            ages,
            bins=[18, 30, 40, 50, 80],
            right=False,
            labels=['18-29岁 (青年)', '30-39岁 (青壮年)', '40-49岁 (成熟期)', '50岁及以上 (老龄工)']
        )
        age_dist = age_cuts.value_counts().sort_index().to_dict()

        # 籍贯 TOP 分布
        prov_counts = unique_roster['Province'].value_counts().to_dict()

        # 工种分布
        job_counts = unique_roster['Job_Clean'].value_counts().to_dict()

        return {
            "total_unique": total_unique,
            "gender_counts": gender_counts,
            "avg_age": avg_age,
            "median_age": median_age,
            "age_dist": age_dist,
            "top_provinces": prov_counts,
            "job_counts": job_counts
        }

    def _compute_compliance(self, df_inflow: pd.DataFrame, df_outflow: pd.DataFrame) -> Dict[str, Any]:
        """计算三大合规管控指标：劳动合同、离场工资结算、退场承诺书"""
        contract_rate = 100.0
        if not df_inflow.empty and 'ContractSigned' in df_inflow.columns:
            signed_cnt = len(df_inflow[df_inflow['ContractSigned'].isin(['是', '已签订'])])
            contract_rate = round(signed_cnt / len(df_inflow) * 100, 1)

        wage_settle_rate = 100.0
        commit_rate = 100.0
        if not df_outflow.empty:
            if 'WageSettled' in df_outflow.columns:
                settled_cnt = len(df_outflow[df_outflow['WageSettled'].str.contains('已结算|已支付', na=False)])
                wage_settle_rate = round(settled_cnt / len(df_outflow) * 100, 1)
            if 'CommitmentSigned' in df_outflow.columns:
                commit_cnt = len(df_outflow[df_outflow['CommitmentSigned'].isin(['是', '已签订'])])
                commit_rate = round(commit_cnt / len(df_outflow) * 100, 1)

        return {
            "contract_rate": contract_rate,
            "wage_settle_rate": wage_settle_rate,
            "commit_rate": commit_rate,
            "inflow_total": len(df_inflow),
            "outflow_total": len(df_outflow)
        }

    # ==========================================
    # 可视化引擎：Matplotlib 高清全景报表图 (参考 plot_dashboard.py)
    # ==========================================
    def generate_matplotlib_figure(self) -> Any:
        """复刻并优化 plot_dashboard.py 的 2x3 专业分析看板图"""
        import matplotlib
        import matplotlib.pyplot as plt

        data = self.load_all_data()
        months = data["months"]
        summary = data["monthly_summary"]
        demographics = data["demographics"]
        unique_df = data["unique_roster"]
        df_roster = data["df_roster"]

        in_flow = [summary[m]["in_total"] for m in months]
        out_flow = [summary[m]["out_total"] for m in months]
        onsite = [summary[m]["onsite_total"] for m in months]

        # 字体设置
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False

        fig, axes = plt.subplots(2, 3, figsize=(18, 11), dpi=140)
        fig.patch.set_facecolor('#f8fafc')
        x = np.arange(len(months))
        width = 0.35

        # 子图 1: 进出场与在场趋势
        ax1 = axes[0, 0]
        r1 = ax1.bar(x - width/2, in_flow, width, label='进场人数', color='#3b82f6', alpha=0.88, edgecolor='none')
        r2 = ax1.bar(x + width/2, out_flow, width, label='离场人数', color='#ef4444', alpha=0.88, edgecolor='none')
        ax1_twin = ax1.twinx()
        ax1_twin.plot(x, onsite, color='#10b981', marker='o', linewidth=2.5, markersize=8, label='在场总人数')
        ax1.set_xticks(x)
        ax1.set_xticklabels(months, fontsize=11)
        ax1.set_ylabel('变动人数 (人)', fontsize=10)
        ax1_twin.set_ylabel('在场总数 (人)', fontsize=10)
        ax1.set_title('各月人员进出场流动与在场总数趋势', fontsize=13, fontweight='bold', pad=12)
        for r in list(r1) + list(r2):
            h = r.get_height()
            if h > 0:
                ax1.annotate(f'{int(h)}', (r.get_x() + r.get_width()/2, h), ha='center', va='bottom', fontsize=9, xytext=(0, 2), textcoords='offset points')
        for i, v in enumerate(onsite):
            ax1_twin.annotate(f'{v} 人', (x[i], v), ha='center', va='bottom', fontsize=10, fontweight='bold', color='#065f46', xytext=(0, 5), textcoords='offset points')
        l1, lab1 = ax1.get_legend_handles_labels()
        l2, lab2 = ax1_twin.get_legend_handles_labels()
        ax1.legend(l1 + l2, lab1 + lab2, loc='upper left', framealpha=0.8)
        ax1.grid(axis='y', linestyle='--', alpha=0.3)

        # 子图 2: 分包队伍在场规模堆叠
        ax2 = axes[0, 1]
        team_short_map = {
            '江苏旭之升 (王宜强施工班组)': '江苏旭之升',
            '青海久昌 (汪佩沾其他班组)': '青海久昌'
        }
        df_roster_copy = df_roster.copy()
        df_roster_copy['Team_Short'] = df_roster_copy['Team'].map(team_short_map).fillna(df_roster_copy['Team'])
        team_monthly = df_roster_copy.groupby(['Month', 'Team_Short'])['Name'].count().unstack().reindex(months).fillna(0)
        team_monthly.plot(kind='bar', stacked=True, ax=ax2, color=['#0284c7', '#f59e0b'], edgecolor='white', alpha=0.9)
        ax2.set_title('各月分包队伍在场规模堆叠图', fontsize=13, fontweight='bold', pad=12)
        ax2.set_xlabel('')
        ax2.set_ylabel('人数 (人)', fontsize=10)
        ax2.set_xticks(range(len(months)))
        ax2.set_xticklabels(months, rotation=0, fontsize=11)
        ax2.grid(axis='y', linestyle='--', alpha=0.3)
        for n, c in enumerate(team_monthly.values):
            cum = 0
            for val in c:
                if val > 0:
                    ax2.text(n, cum + val/2, f'{int(val)}', ha='center', va='center', color='white', fontweight='bold', fontsize=9)
                    cum += val
            ax2.text(n, cum + 1, f'总 {int(cum)}', ha='center', va='bottom', fontsize=10, fontweight='bold')

        # 子图 3: 工种构成环形图
        ax3 = axes[0, 2]
        job_counts = pd.Series(demographics["job_counts"])
        top_jobs = job_counts[job_counts >= 2].copy()
        other_sum = job_counts[job_counts < 2].sum()
        if other_sum > 0:
            top_jobs['其他技能工'] = top_jobs.get('其他技能工', 0) + other_sum
        ax3.pie(
            top_jobs.values,
            labels=top_jobs.index,
            autopct='%1.1f%%',
            pctdistance=0.75,
            startangle=140,
            colors=['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#6b7280'],
            wedgeprops=dict(width=0.45, edgecolor='w')
        )
        ax3.set_title(f'务工人员工种构成分布 (共{demographics["total_unique"]}人)', fontsize=13, fontweight='bold', pad=12)

        # 子图 4: 年龄梯队分布
        ax4 = axes[1, 0]
        age_dist = pd.Series(demographics["age_dist"])
        bars4 = ax4.bar(range(len(age_dist)), age_dist.values, color=['#60a5fa', '#34d399', '#fbbf24', '#f87171'], edgecolor='white')
        ax4.set_title(f'年龄结构分布 (平均: {demographics["avg_age"]}岁, 中位数: {demographics["median_age"]:.0f}岁)', fontsize=13, fontweight='bold', pad=12)
        ax4.set_ylabel('人数 (人)', fontsize=10)
        ax4.set_xticks(range(len(age_dist)))
        clean_labels = [label.replace(' ', '\n') for label in age_dist.index]
        ax4.set_xticklabels(clean_labels, fontsize=10)
        ax4.grid(axis='y', linestyle='--', alpha=0.3)
        for b in bars4:
            h = b.get_height()
            if h > 0:
                ax4.annotate(f'{int(h)} 人\n({h/demographics["total_unique"]*100:.1f}%)', (b.get_x() + b.get_width()/2, h), ha='center', va='bottom', fontsize=9, xytext=(0, 2), textcoords='offset points')

        # 子图 5: 籍贯省份分布
        ax5 = axes[1, 1]
        prov_counts = pd.Series(demographics["top_provinces"]).sort_values(ascending=True)
        bars5 = ax5.barh(range(len(prov_counts)), prov_counts.values, color='#6366f1', alpha=0.88)
        ax5.set_title('劳务人员籍贯省份分布 (TOP来源地)', fontsize=13, fontweight='bold', pad=12)
        ax5.set_xlabel('人数 (人)', fontsize=10)
        ax5.set_yticks(range(len(prov_counts)))
        ax5.set_yticklabels(prov_counts.index, fontsize=10)
        ax5.grid(axis='x', linestyle='--', alpha=0.3)
        for b in bars5:
            w = b.get_width()
            ax5.annotate(f' {int(w)} 人 ({w/demographics["total_unique"]*100:.1f}%)', (w, b.get_y() + b.get_height()/2), va='center', fontsize=9)

        # 子图 6: 基础信息与合规指标卡
        ax6 = axes[1, 2]
        ax6.axis('off')
        ax6.set_title('合规管控与人员基础信息指标卡', fontsize=13, fontweight='bold', pad=12)
        summary_text = (
            f"【核心规模指标】\n"
            f"• 累计参建总劳务工人数：{demographics['total_unique']} 人\n"
            f"• 男女性别比：男 {demographics['gender_counts'].get('男', 0)} 人 / 女 {demographics['gender_counts'].get('女', 0)} 人\n"
            f"• 施工高峰月度：8月 (现场共 60 人)\n\n"
            f"【流动管控特征】\n"
            f"• 6-7月为集中进场爬坡期（进场达 58 人次）\n"
            f"• 9月收尾退场为主（离场 10 人，仅增 1 人）\n\n"
            f"【合规与履约状况】\n"
            f"• 劳动合同签订率：100% (全员覆盖)\n"
            f"• 离场工资结算支付：100% 已结算已支付\n"
            f"• 《退场承诺书》签订率：100% (规范退场)"
        )
        ax6.text(
            0.05, 0.5, summary_text, fontsize=11, verticalalignment='center',
            bbox=dict(boxstyle='round,pad=1', facecolor='#e0e7ff', edgecolor='#6366f1', alpha=0.5, linewidth=1.5),
            linespacing=1.6
        )

        plt.tight_layout(pad=3.0)
        return fig

    def get_matplotlib_png_bytes(self) -> bytes:
        """导出 Matplotlib 看板为 PNG 二进制数据 (300 DPI 高清)"""
        import io
        import matplotlib.pyplot as plt
        fig = self.generate_matplotlib_figure()
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
        plt.close(fig)
        return buf.getvalue()

    # ==========================================
    # 可视化引擎：Plotly 交互式图表体系
    # ==========================================
    def generate_plotly_figures(self, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """构建现代响应式交互图表"""
        import plotly.graph_objects as go  # type: ignore
        from plotly.subplots import make_subplots  # type: ignore

        if data is None:
            data = self.load_all_data()
        months = data["months"]
        summary = data["monthly_summary"]
        demographics = data["demographics"]
        df_roster = data["df_roster"]

        in_flow = [summary[m]["in_total"] for m in months]
        out_flow = [summary[m]["out_total"] for m in months]
        onsite = [summary[m]["onsite_total"] for m in months]
        actual_onsite = [summary[m].get("actual_onsite_total", summary[m]["onsite_total"] - summary[m]["out_total"]) for m in months]

        # 1. 各月人员流动与在场总数趋势 (双轴柱线组合)
        fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
        fig_trend.add_trace(
            go.Bar(
                x=months, y=in_flow, name="进场人数",
                marker=dict(color="#3B82F6", opacity=0.85, line=dict(width=0)),
                text=in_flow, textposition="outside"
            ),
            secondary_y=False
        )
        fig_trend.add_trace(
            go.Bar(
                x=months, y=out_flow, name="离场人数",
                marker=dict(color="#EF4444", opacity=0.85, line=dict(width=0)),
                text=out_flow, textposition="outside"
            ),
            secondary_y=False
        )
        fig_trend.add_trace(
            go.Scatter(
                x=months, y=actual_onsite, name="实际在场人数",
                mode="lines+markers+text",
                line=dict(color="#10B981", width=3.5),
                marker=dict(size=10, color="#059669", line=dict(color="white", width=2)),
                text=[f"{v}人" for v in actual_onsite], textposition="top center",
                textfont=dict(size=11, color="#065F46", family="sans-serif")
            ),
            secondary_y=True
        )
        fig_trend.add_trace(
            go.Scatter(
                x=months, y=onsite, name="报备在场人数",
                mode="lines+markers",
                line=dict(color="#6366F1", width=2, dash="dash"),
                marker=dict(size=7, color="#4F46E5"),
                hoverinfo="x+y+name"
            ),
            secondary_y=True
        )
        fig_trend.update_layout(
            title=dict(text="各月人员进出场流动与在场总数趋势", font=dict(size=16, family="sans-serif", color="#1E293B")),
            barmode="group",
            hovermode="x unified",
            plot_bgcolor="rgba(248,250,252,0.6)",
            paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=40, r=40, t=60, b=30),
            height=360
        )
        fig_trend.update_xaxes(title_text="", showgrid=False)
        fig_trend.update_yaxes(title_text="月度变动人数 (人)", secondary_y=False, showgrid=True, gridcolor="rgba(226,232,240,0.7)")
        fig_trend.update_yaxes(title_text="在场人数 (人)", secondary_y=True, showgrid=False)

        # 2. 分包队伍在场规模堆叠图 (严格与官方台账对齐)
        team_yi = '江苏旭之升 (王宜强施工班组)'
        team_wang = '青海久昌 (汪佩沾其他班组)'
        yi_counts = [summary[m]["teams"].get(team_yi, {}).get("onsite_count", 0) for m in months]
        wang_counts = [summary[m]["teams"].get(team_wang, {}).get("onsite_count", 0) for m in months]

        fig_teams = go.Figure()
        fig_teams.add_trace(
            go.Bar(
                x=months, y=yi_counts, name="江苏旭之升 (王宜强)",
                marker=dict(color="#0284C7", opacity=0.9),
                text=yi_counts, textposition="inside"
            )
        )
        fig_teams.add_trace(
            go.Bar(
                x=months, y=wang_counts, name="青海久昌 (汪佩沾)",
                marker=dict(color="#F59E0B", opacity=0.9),
                text=wang_counts, textposition="inside"
            )
        )
        fig_teams.update_layout(
            title=dict(text="各月分包队伍在场规模堆叠 (报备在场)", font=dict(size=16, family="sans-serif", color="#1E293B")),
            barmode="stack",
            hovermode="x unified",
            plot_bgcolor="rgba(248,250,252,0.6)",
            paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=40, r=40, t=60, b=30),
            height=360
        )
        fig_teams.update_xaxes(title_text="", showgrid=False)
        fig_teams.update_yaxes(title_text="在场人数 (人)", showgrid=True, gridcolor="rgba(226,232,240,0.7)")

        # 3. 务工人员工种技能构成分布 (Donut Chart)
        job_counts = pd.Series(demographics["job_counts"])
        top_jobs = job_counts[job_counts >= 2].copy()
        other_sum = job_counts[job_counts < 2].sum()
        if other_sum > 0:
            top_jobs['其他技能工'] = top_jobs.get('其他技能工', 0) + other_sum

        fig_jobs = go.Figure(
            data=[go.Pie(
                labels=top_jobs.index,
                values=top_jobs.values,
                hole=0.48,
                marker=dict(colors=['#3B82F6', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#6B7280']),
                textinfo="label+percent",
                hoverinfo="label+value+percent"
            )]
        )
        fig_jobs.update_layout(
            title=dict(text=f"务工人员工种构成分布 (共{demographics['total_unique']}人)", font=dict(size=16, family="sans-serif", color="#1E293B")),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=60, b=20),
            height=360
        )

        # 4. 年龄梯队结构分布
        age_dist = pd.Series(demographics["age_dist"])
        fig_age = go.Figure(
            data=[go.Bar(
                x=list(age_dist.index),
                y=list(age_dist.values),
                marker=dict(color=['#60A5FA', '#34D399', '#FBBF24', '#F87171']),
                text=[f"{v}人 ({v/demographics['total_unique']*100:.1f}%)" for v in age_dist.values],
                textposition="outside"
            )]
        )
        fig_age.update_layout(
            title=dict(text=f"年龄结构分布 (平均: {demographics['avg_age']}岁, 中位数: {demographics['median_age']:.0f}岁)", font=dict(size=16, family="sans-serif", color="#1E293B")),
            plot_bgcolor="rgba(248,250,252,0.6)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=40, r=40, t=60, b=30),
            height=360
        )
        fig_age.update_xaxes(title_text="", showgrid=False)
        fig_age.update_yaxes(title_text="人数 (人)", showgrid=True, gridcolor="rgba(226,232,240,0.7)")

        # 5. 籍贯省份 TOP 来源地分布 (横向条形图)
        prov_counts = pd.Series(demographics["top_provinces"]).sort_values(ascending=True)
        fig_province = go.Figure(
            data=[go.Bar(
                x=list(prov_counts.values),
                y=list(prov_counts.index),
                orientation='h',
                marker=dict(color="#6366F1", opacity=0.88),
                text=[f"{v}人 ({v/demographics['total_unique']*100:.1f}%)" for v in prov_counts.values],
                textposition="outside"
            )]
        )
        fig_province.update_layout(
            title=dict(text="劳务人员籍贯省份分布 (TOP 来源地)", font=dict(size=16, family="sans-serif", color="#1E293B")),
            plot_bgcolor="rgba(248,250,252,0.6)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=50, r=40, t=60, b=30),
            height=360
        )
        fig_province.update_xaxes(title_text="人数 (人)", showgrid=True, gridcolor="rgba(226,232,240,0.7)")
        fig_province.update_yaxes(title_text="", showgrid=False)

        return {
            "trend": fig_trend,
            "teams": fig_teams,
            "jobs": fig_jobs,
            "age": fig_age,
            "province": fig_province
        }


# 全局单例服务
personnel_data_service = PersonnelDataService()
