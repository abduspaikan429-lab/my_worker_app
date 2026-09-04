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

from repositories.shadow_worker_repository import ShadowWorkerRepository
from repositories.worker_repository import JsonWorkerRepository
from services.worker_service import WorkerService


MASTER_DIR = Path("data/master")
STATE_FILE = Path("data/master_state.json")
HISTORY_DIR = Path("data/master_history")
ID_COL = "身份证号"

json_repository = JsonWorkerRepository(file_path=lambda: STATE_FILE)
shadow_repository = ShadowWorkerRepository(json_repo=json_repository)
worker_service = WorkerService(repository=shadow_repository)


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


def _write_master_excel(df: pd.DataFrame) -> None:
    """将最新全量人员主表同步写入 data/master 目录下的 Excel 文件中，保持磁盘物理文件与系统状态实时一致。"""
    if df is None or df.empty:
        return
    try:
        MASTER_DIR.mkdir(parents=True, exist_ok=True)
        target_path = MASTER_DIR / "全量劳务人员主表.xlsx"
        with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="全量人员主表")

        for f in MASTER_DIR.glob("*.xlsx"):
            if f.name != "全量劳务人员主表.xlsx" and not f.name.startswith("~$"):
                try:
                    with pd.ExcelWriter(f, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="劳务人员汇总表")
                except Exception:
                    pass
    except Exception as e:
        print(f"Error writing master Excel: {e}")


def commit_update(incoming_df: pd.DataFrame, source_files: list[str] | None = None) -> dict[str, Any]:
    """保存主表新版本和本次变化快照，并同步更新 data/master 物理 Excel 文件。"""
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
    _write_master_excel(current_df)

    preview["version"] = version
    preview["updated_at"] = now.isoformat(timespec="seconds")
    preview["changes"] = changes
    return preview


def upsert_master_workers(
    workers_list: list[dict[str, Any]], source: str = "onboarding"
) -> dict[str, Any]:
    """
    增量插入或更新项目主表中的人员档案（支持进场流程归档人员与手动录入人员的自动同步）。
    写入 master_state.json，同步更新 SQLite 镜像，并保存版本快照。
    """
    if not workers_list:
        return {
            "added": 0,
            "updated": 0,
            "total": 0,
            "version": "",
            "error": "工人列表为空",
        }

    current_df = load_master_df()
    state = _read_state()
    existing_rows = state.get("rows") or []

    # 建立现有主表记录的索引查找表
    # 1. 身份证号索引
    id_map: dict[str, int] = {}
    for idx, r in enumerate(existing_rows):
        sid = clean_value(r.get(ID_COL, ""))
        if sid and len(sid) >= 15:
            id_map[sid] = idx

    # 2. 姓名+班组 / 姓名 索引
    name_team_map: dict[str, int] = {}
    name_map: dict[str, int] = {}
    for idx, r in enumerate(existing_rows):
        nm = clean_value(r.get("姓名", ""))
        tm = clean_value(r.get("班组", ""))
        if nm and tm:
            name_team_map[f"{nm}_{tm}"] = idx
        if nm and nm not in name_map:
            name_map[nm] = idx

    added_count = 0
    updated_count = 0
    new_records = []
    updated_records = []

    # 基础列集合
    existing_cols = (
        list(current_df.columns)
        if not current_df.empty
        else [
            "姓名",
            "班组",
            "身份证号",
            "手机号",
            "工种",
            "结算单价/标准",
            "工资卡号",
            "开户银行",
            "进场日期",
            "在场/进退场状态",
            "人员类型",
            "分包/所属企业",
            "合同签订状态",
        ]
    )

    for raw_worker in workers_list:
        if not isinstance(raw_worker, dict):
            continue
        w = {k: clean_value(v) for k, v in raw_worker.items()}
        nm = w.get("姓名", "")
        if not nm:
            continue
        sid = w.get("身份证号", "")
        tm = w.get("班组", "")

        # 统一同义字段
        if "银行卡号" in w and not w.get("工资卡号"):
            w["工资卡号"] = w["银行卡号"]
        if "日薪" in w and not w.get("结算单价/标准") and w["日薪"]:
            wage_val = w["日薪"]
            w["结算单价/标准"] = (
                f"基本工资：{wage_val}元/天"
                if not str(wage_val).endswith("元/天")
                else str(wage_val)
            )

        # 查找现有记录（身份证优先，其次 姓名+班组，最后 姓名）
        target_idx = None
        if sid and len(sid) >= 15 and sid in id_map:
            target_idx = id_map[sid]
        elif f"{nm}_{tm}" in name_team_map:
            target_idx = name_team_map[f"{nm}_{tm}"]
        elif nm in name_map:
            target_idx = name_map[nm]

        if target_idx is not None:
            # 更新已有记录（增量更新非空字段）
            target = existing_rows[target_idx]
            has_diff = False
            for k, v in w.items():
                if v and target.get(k) != v:
                    target[k] = v
                    has_diff = True
            if has_diff:
                updated_count += 1
                updated_records.append(target)
        else:
            # 新增人员记录
            new_row = {col: "" for col in existing_cols}
            new_row.update(w)
            if not new_row.get("在场/进退场状态"):
                new_row["在场/进退场状态"] = "在场"
            if not new_row.get("人员类型"):
                new_row["人员类型"] = "劳务人员"
            if not new_row.get("合同签订状态"):
                new_row["合同签订状态"] = "已签订"

            existing_rows.append(new_row)
            added_count += 1
            new_records.append(new_row)

            # 更新索引
            new_idx = len(existing_rows) - 1
            if sid and len(sid) >= 15:
                id_map[sid] = new_idx
            if nm and tm:
                name_team_map[f"{nm}_{tm}"] = new_idx
            if nm and nm not in name_map:
                name_map[nm] = new_idx

    now = datetime.now()
    version = now.strftime("%Y%m%d_%H%M%S_%f")
    changes = {
        "new_rows": new_records,
        "updated_rows": updated_records,
        "missing_from_import": [],
        "source_files": [f"{source}_sync"],
    }
    state["version"] = version
    state["updated_at"] = now.isoformat(timespec="seconds")
    state["rows"] = existing_rows
    state["last_changes"] = changes

    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    worker_service.save_workers(existing_rows)
    (HISTORY_DIR / f"{version}.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    _write_master_excel(pd.DataFrame(existing_rows))

    return {
        "added": added_count,
        "updated": updated_count,
        "total": len(existing_rows),
        "version": version,
        "updated_at": state["updated_at"],
        "changes": changes,
        "error": "",
    }



