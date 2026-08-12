from __future__ import annotations

import re
from typing import Any, Dict, List, Optional
import pandas as pd

_NAME_CANDS = ["姓名", "名", "名字", "人员姓名", "Name"]
_ID_CANDS = ["身份证号", "身份证号码", "身份证", "证件号码", "证件号"]
_COMPANY_CANDS = ["分包/所属企业", "所属企业", "分包单位", "单位", "劳务单位"]
_DAYS_CANDS = [
    "考勤天数",
    "出勤天数",
    "实际天数",
    "天数",
    "实出勤",
    "实际出勤天数",
    "当月出勤天数",
    "小计",
    "合计",
    "合计天数",
    "总计",
    "天",
    "出勤合计",
    "出勤",
    "Days",
]
_SERIAL_CANDS = ["序号", "编号", "序", "No.", "NO."]
_NAME_EXCLUDED_HINTS = (
    "签字",
    "签名",
    "申明",
    "声明",
    "项目",
    "工程",
    "班组",
    "负责人",
    "名称",
    "编号",
    "序号",
    "日期",
    "备注",
    "单位",
    "工资",
    "考勤",
    "人员",
)


class AttendanceService:
    """Service layer for Attendance Excel parsing, cleaning, and day extraction."""

    @staticmethod
    def normalize_header(value: Any) -> str:
        return re.sub(r"[\s\u3000]+", "", str(value).strip())

    @staticmethod
    def clean_identity(value: Any) -> str:
        text = AttendanceService.raw_attendance_value(value).replace(" ", "").upper()
        return text[:-2] if text.endswith(".0") else text

    @staticmethod
    def is_valid_identity(value: Any) -> bool:
        return len(AttendanceService.clean_identity(value)) >= 15

    @staticmethod
    def raw_attendance_value(value: Any) -> str:
        if value is None or pd.isna(value):
            return ""
        text = str(value).strip()
        if text.lower() in {"nan", "none", "null", "<na>"}:
            return ""
        return text

    @staticmethod
    def attendance_state(value: Any) -> str:
        text = AttendanceService.raw_attendance_value(value)
        if not text or text in {"--", "—", "-", "/"}:
            return "缺勤/无记录"
        if any(mark in text for mark in ["√", "✓", "✔", "出勤", "正常"]):
            return "有考勤"
        if re.search(r"\d{1,2}:\d{2}", text):
            return "有考勤"
        return "待确认"

    @staticmethod
    def day_number(column: Any) -> Optional[int]:
        text = str(column).strip().replace("号", "").replace("日", "")
        if text.isdigit():
            day = int(text)
            if 1 <= day <= 31:
                return day
        return None

    @staticmethod
    def resolve_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
        for c in candidates:
            if c in df.columns:
                return c
        normalized_columns = {
            column: AttendanceService.normalize_header(column)
            for column in df.columns
        }
        for candidate in candidates:
            normalized_candidate = AttendanceService.normalize_header(candidate)
            for column, normalized_column in normalized_columns.items():
                if normalized_candidate == "名":
                    if normalized_column == "名":
                        return column
                elif normalized_candidate and normalized_candidate in normalized_column:
                    return column
        return None

    @staticmethod
    def detect_header_row(file_obj: Any) -> int:
        try:
            raw = pd.read_excel(file_obj, dtype=str, header=None, nrows=12)
            for i in range(min(12, len(raw))):
                row_cells = [
                    AttendanceService.normalize_header(v)
                    for v in raw.iloc[i].tolist()
                ]
                if any(
                    (
                        "姓名" in cell
                        or "人员姓名" in cell
                        or "名字" in cell
                        or cell == "名"
                        or cell.lower() == "name"
                    )
                    for cell in row_cells
                ):
                    return i
        except Exception:
            pass
        return 0

    @staticmethod
    def safe_read_excel(file_obj: Any, label: str) -> pd.DataFrame:
        if file_obj is None:
            return pd.DataFrame()
        try:
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)
            header_row = AttendanceService.detect_header_row(file_obj)

            if hasattr(file_obj, "seek"):
                file_obj.seek(0)
            df = pd.read_excel(file_obj, dtype=str, header=header_row)
            df.columns = [
                re.sub(r"[\s\u3000]+", "", str(c).strip()) for c in df.columns
            ]

            name_col = AttendanceService.resolve_col(df, _NAME_CANDS)
            if name_col:
                id_col = AttendanceService.resolve_col(df, _ID_CANDS)
                serial_col = AttendanceService.resolve_col(df, _SERIAL_CANDS)

                def is_person_row(row):
                    name = row.get(name_col, "")
                    if (
                        AttendanceService.is_valid_identity(row.get(id_col, ""))
                        if id_col
                        else False
                    ):
                        return True
                    serial = (
                        str(row.get(serial_col, "")).strip() if serial_col else ""
                    )
                    compact = AttendanceService.normalize_header(name)
                    looks_like_name = bool(
                        compact
                        and not any(
                            hint in compact for hint in _NAME_EXCLUDED_HINTS
                        )
                        and re.fullmatch(r"[\u4e00-\u9fff·]{2,12}", compact)
                    )
                    return looks_like_name and (
                        bool(re.fullmatch(r"\d+(?:\.0)?", serial))
                        or not serial_col
                    )

                mask = (
                    df[name_col].notna()
                    & (df[name_col].astype(str).str.strip() != "")
                    & df.apply(is_person_row, axis=1)
                )
                df = df[mask].reset_index(drop=True)

            df["__来源__"] = label
            id_col = AttendanceService.resolve_col(df, _ID_CANDS)
            if id_col:
                df["身份证号"] = df[id_col].apply(AttendanceService.clean_identity)
            company_col = AttendanceService.resolve_col(df, _COMPANY_CANDS)
            if company_col and company_col != "分包/所属企业":
                df["分包/所属企业"] = df[company_col].astype(str).str.strip()
            return df
        except Exception:
            return pd.DataFrame()

    @staticmethod
    def extract_daily_source(
        df: pd.DataFrame, source_name: str
    ) -> Dict[str, Any]:
        if df is None or df.empty:
            return {}
        name_col = AttendanceService.resolve_col(df, _NAME_CANDS)
        id_col = AttendanceService.resolve_col(df, _ID_CANDS)
        if name_col is None:
            return {}

        day_cols = {
            day: col
            for col in df.columns
            if (day := AttendanceService.day_number(col)) is not None
        }
        records = {}
        for _, row in df.iterrows():
            name = AttendanceService.raw_attendance_value(row.get(name_col, ""))
            identity = (
                AttendanceService.raw_attendance_value(row.get(id_col, ""))
                if id_col
                else ""
            )
            key = identity if len(identity) >= 15 else f"姓名:{name}"
            if not name or not key:
                continue
            item = records.setdefault(
                key,
                {
                    "姓名": name,
                    "身份证号": identity,
                    "来源": source_name,
                    "days": {},
                },
            )
            if not item["身份证号"] and identity:
                item["身份证号"] = identity
            for day, col in day_cols.items():
                value = AttendanceService.raw_attendance_value(row.get(col, ""))
                if value:
                    item["days"][day] = value
        return records
