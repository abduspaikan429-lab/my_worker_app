"""项目人员主数据存储与版本变化识别。

官网系统是人员基础信息的权威来源，本模块只负责把官网导出的合并结果
保存为项目主表，并记录每次同步产生的新增、变更和未出现在本次导出中的
人员。进场流程数据不写入这里。
"""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from typing import Any

import pandas as pd

from repositories.worker_repository import JsonWorkerRepository
from services.worker_service import WorkerService


MASTER_DIR = Path("data/master")
STATE_FILE = Path("data/master_state.json")
HISTORY_DIR = Path("data/master_history")
ID_COL = "身份证号"

json_repository = JsonWorkerRepository(file_path=lambda: STATE_FILE)
worker_service = WorkerService(repository=json_repository)


def clean_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null", "<na>"}:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    if "e+" in text.lower() or "e-" in text.lower():
        try:
            text = f"{float(text):.0f}"
        except (TypeError, ValueError):
            pass
    return text


def normalize_df(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()
    out.columns = [str(c).replace("\n", "").strip() for c in out.columns]
    out = out.loc[:, [c for c in out.columns if not c.lower().startswith("unnamed")]]
    for col in out.columns:
        out[col] = out[col].map(clean_value)

    if ID_COL not in out.columns:
        return pd.DataFrame()

    out = out[out[ID_COL].str.len() >= 15].copy()
    out = out.drop_duplicates(subset=[ID_COL], keep="last").reset_index(drop=True)
    return out


def _read_excel_master(path: Path) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    try:
        book = pd.ExcelFile(path)
        for sheet in book.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=30, dtype=str)
            header_row = None
            for idx in range(len(raw)):
                values = {clean_value(v) for v in raw.iloc[idx].tolist()}
                if "姓名" in values and ID_COL in values:
                    header_row = idx
                    break
            if header_row is None:
                continue
            part = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype=str)
            part.columns = [str(c).replace("\n", "").strip() for c in part.columns]
            frames.append(part)
    except Exception:
        return pd.DataFrame()

    if not frames:
        return pd.DataFrame()
    return normalize_df(pd.concat(frames, ignore_index=True))


def _empty_state() -> dict[str, Any]:
    return {
        "version": "",
        "updated_at": "",
        "rows": [],
        "last_changes": {
            "new_rows": [],
            "updated_rows": [],
            "missing_from_import": [],
            "source_files": [],
        },
    }


def _read_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return _empty_state()
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        state = _empty_state()
        state.update(data)
        if "rows" not in data or data["rows"] is None:
            state["rows"] = json_repository.get_all_workers()
        else:
            state["rows"] = data.get("rows", [])
        state["last_changes"] = {
            **_empty_state()["last_changes"],
            **(data.get("last_changes") or {}),
        }
        return state
    except (OSError, json.JSONDecodeError, TypeError):
        return _empty_state()


def load_master_df() -> pd.DataFrame:
    """加载当前项目主表；首次使用时从 data/master 下的现有 Excel 初始化。"""
    state = _read_state()
    if state.get("version") or state.get("rows"):
        return normalize_df(pd.DataFrame(state.get("rows") or []))

    frames = []
    for path in sorted(MASTER_DIR.glob("*.xlsx")):
        frame = _read_excel_master(path)
        if not frame.empty:
            frames.append(frame)
    if not frames:
        return pd.DataFrame()
    return normalize_df(pd.concat(frames, ignore_index=True))


def get_last_changes() -> dict[str, Any]:
    return _read_state().get("last_changes", _empty_state()["last_changes"])


def _row_values(row: pd.Series, columns: list[str]) -> dict[str, str]:
    return {col: clean_value(row.get(col, "")) for col in columns}


def preview_update(current_df: pd.DataFrame | None, incoming_df: pd.DataFrame | None) -> dict[str, Any]:
    """预览官网本次导出对主表的影响，不写磁盘。"""
    current = normalize_df(current_df)
    incoming = normalize_df(incoming_df)
    if incoming.empty:
        return {
            "current_df": current,
            "incoming_df": incoming,
            "merged_df": current,
            "new_rows": pd.DataFrame(),
            "updated_rows": pd.DataFrame(),
            "missing_from_import": pd.DataFrame(),
            "error": f"官网导入数据缺少有效的{ID_COL}，无法更新主表。",
        }

    columns: list[str] = []
    for col in list(current.columns) + list(incoming.columns):
        if col not in columns:
            columns.append(col)
    if ID_COL not in columns:
        return {"error": f"数据中没有{ID_COL}列。"}

    old = current.set_index(ID_COL, drop=False) if not current.empty else pd.DataFrame(index=[])
    new = incoming.set_index(ID_COL, drop=False)
    old_keys = set(old.index)
    new_keys = set(new.index)

    new_keys_only = sorted(new_keys - old_keys)
    common_keys = sorted(old_keys & new_keys)
    updated_keys: list[str] = []
    merged_rows: list[dict[str, str]] = []

    for key in sorted(old_keys | new_keys):
        old_row = old.loc[key] if key in old.index else None
        new_row = new.loc[key] if key in new.index else None
        row = {}
        for col in columns:
            old_val = clean_value(old_row.get(col, "")) if old_row is not None else ""
            new_val = clean_value(new_row.get(col, "")) if new_row is not None else ""
            # 官网新值非空时更新；空值不覆盖原有资料，避免导出缺列造成数据丢失。
            row[col] = new_val if new_val else old_val
        merged_rows.append(row)
        if key in common_keys:
            old_values = _row_values(old_row, columns)
            new_values = _row_values(new_row, columns)
            if any(new_values[c] and new_values[c] != old_values.get(c, "") for c in columns if c != ID_COL):
                updated_keys.append(key)

    merged = pd.DataFrame(merged_rows, columns=columns)
    new_rows = incoming[incoming[ID_COL].isin(new_keys_only)].copy()
    updated_rows = merged[merged[ID_COL].isin(updated_keys)].copy()
    missing = (
        current[current[ID_COL].isin(sorted(old_keys - new_keys))].copy()
        if ID_COL in current.columns
        else pd.DataFrame(columns=columns)
    )

    return {
        "current_df": current,
        "incoming_df": incoming,
        "merged_df": normalize_df(merged),
        "new_rows": normalize_df(new_rows),
        "updated_rows": normalize_df(updated_rows),
        "missing_from_import": normalize_df(missing),
        "error": "",
    }


def commit_update(incoming_df: pd.DataFrame, source_files: list[str] | None = None) -> dict[str, Any]:
    """保存主表新版本和本次变化快照。"""
    preview = preview_update(load_master_df(), incoming_df)
    if preview.get("error"):
        return preview

    now = datetime.now()
    version = now.strftime("%Y%m%d_%H%M%S_%f")
    current_df = preview["merged_df"]
    changes = {
        "new_rows": preview["new_rows"].to_dict(orient="records"),
        "updated_rows": preview["updated_rows"].to_dict(orient="records"),
        "missing_from_import": preview["missing_from_import"].to_dict(orient="records"),
        "source_files": source_files or [],
    }
    workers = current_df.to_dict(orient="records")
    state = {
        "version": version,
        "updated_at": now.isoformat(timespec="seconds"),
        "rows": workers,
        "last_changes": changes,
    }

    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    worker_service.save_workers(workers)
    (HISTORY_DIR / f"{version}.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    preview["version"] = version
    preview["updated_at"] = now.isoformat(timespec="seconds")
    preview["changes"] = changes
    return preview


