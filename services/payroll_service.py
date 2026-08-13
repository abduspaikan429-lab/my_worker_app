from __future__ import annotations

import re
from typing import Any, Dict, Optional
import pandas as pd


class PayrollService:
    """Service layer handling daily rate parsing, wage calculation, and summary aggregation."""

    @staticmethod
    def parse_daily_rate(rate_str: Any) -> float:
        """Parse daily rate from strings like '基本工资：450元/天' or '450'."""
        if rate_str is None or pd.isna(rate_str):
            return 0.0
        text = str(rate_str).strip()
        if not text:
            return 0.0

        match = re.search(r"(\d+(?:\.\d+)?)", text)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return 0.0
        return 0.0

    @staticmethod
    def calculate_payroll(
        workers_df: pd.DataFrame,
        days_col: str = "考勤天数",
        rate_col: str = "结算单价/标准",
        override_rate_dict: Optional[Dict[str, float]] = None,
    ) -> pd.DataFrame:
        """Calculate payroll data for workers DataFrame."""
        if workers_df is None or workers_df.empty:
            return pd.DataFrame()

        df = workers_df.copy()
        rates = []
        total_wages = []

        override_dict = override_rate_dict or {}

        for _, row in df.iterrows():
            id_card = str(row.get("身份证号", "")).strip()
            raw_rate = row.get(rate_col, "")

            if id_card and id_card in override_dict:
                daily_rate = float(override_dict[id_card])
            else:
                daily_rate = PayrollService.parse_daily_rate(raw_rate)

            days_val = row.get(days_col, 0)
            try:
                days = (
                    float(days_val)
                    if days_val is not None and not pd.isna(days_val)
                    else 0.0
                )
            except (ValueError, TypeError):
                days = 0.0

            total_wage = round(daily_rate * days, 2)

            rates.append(daily_rate)
            total_wages.append(total_wage)

        df["日薪单价"] = rates
        df["应发工资"] = total_wages
        return df

    @staticmethod
    def summarize_payroll(payroll_df: pd.DataFrame) -> Dict[str, Any]:
        """Summarize total workers count, total days, and total payroll wages."""
        if payroll_df is None or payroll_df.empty:
            return {
                "total_workers": 0,
                "total_days": 0.0,
                "total_wages": 0.0,
            }

        total_workers = len(payroll_df)

        days_series = (
            payroll_df["考勤天数"]
            if "考勤天数" in payroll_df.columns
            else pd.Series([0])
        )
        wages_series = (
            payroll_df["应发工资"]
            if "应发工资" in payroll_df.columns
            else pd.Series([0])
        )

        total_days = pd.to_numeric(days_series, errors="coerce").fillna(0).sum()
        total_wages = pd.to_numeric(wages_series, errors="coerce").fillna(0).sum()

        return {
            "total_workers": total_workers,
            "total_days": float(round(total_days, 2)),
            "total_wages": float(round(total_wages, 2)),
        }
