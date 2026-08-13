import json
from pathlib import Path
from unittest.mock import patch
import pandas as pd
import pytest

from config import settings
from modules import master_data
from repositories.worker_repository import JsonWorkerRepository


def test_load_master_df_does_not_depend_on_sqlite(tmp_path: Path):
    """1. Test load_master_df loads from master_state.json and does not depend on SQLite."""
    test_json_file = tmp_path / "master_state.json"
    test_history_dir = tmp_path / "master_history"

    sample_data = {
        "version": "20260813_120000_000000",
        "updated_at": "2026-08-13T12:00:00",
        "rows": [
            {
                "姓名": "张三",
                "身份证号": "110101199001011234",
                "手机号": "13800138000",
                "工种": "木工",
            }
        ],
        "last_changes": {},
    }
    test_json_file.write_text(json.dumps(sample_data, ensure_ascii=False), encoding="utf-8")

    test_json_repo = JsonWorkerRepository(file_path=test_json_file)

    with patch.object(master_data, "STATE_FILE", test_json_file), \
         patch.object(master_data, "HISTORY_DIR", test_history_dir), \
         patch.object(master_data, "json_repository", test_json_repo), \
         patch.object(settings, "WORKER_STORAGE", "json"):
        df = master_data.load_master_df()
        assert not df.empty
        assert len(df) == 1
        assert df.iloc[0]["姓名"] == "张三"
        assert df.iloc[0]["身份证号"] == "110101199001011234"


def test_load_master_df_retains_extra_fields_not_in_sqlite(tmp_path: Path):
    """2. Test load_master_df retains extra fields in JSON master data not supported by SQLite."""
    test_json_file = tmp_path / "master_state.json"
    test_history_dir = tmp_path / "master_history"

    sample_data = {
        "version": "20260813_120000_000000",
        "updated_at": "2026-08-13T12:00:00",
        "rows": [
            {
                "姓名": "李四",
                "身份证号": "110101199505054321",
                "手机号": "13900139000",
                "工种": "钢筋工",
                "出生日期": "1995-05-05",
                "民族": "汉",
                "家庭住址": "北京市朝阳区某某街道",
                "紧急联系人": "王五",
            }
        ],
        "last_changes": {},
    }
    test_json_file.write_text(json.dumps(sample_data, ensure_ascii=False), encoding="utf-8")

    test_json_repo = JsonWorkerRepository(file_path=test_json_file)

    with patch.object(master_data, "STATE_FILE", test_json_file), \
         patch.object(master_data, "HISTORY_DIR", test_history_dir), \
         patch.object(master_data, "json_repository", test_json_repo), \
         patch.object(settings, "WORKER_STORAGE", "json"):
        df = master_data.load_master_df()
        assert "出生日期" in df.columns
        assert "民族" in df.columns
        assert "家庭住址" in df.columns
        assert "紧急联系人" in df.columns
        assert df.iloc[0]["出生日期"] == "1995-05-05"
        assert df.iloc[0]["民族"] == "汉"
        assert df.iloc[0]["家庭住址"] == "北京市朝阳区某某街道"
        assert df.iloc[0]["紧急联系人"] == "王五"


def test_commit_update_writes_master_state_json(tmp_path: Path):
    """3. Test commit_update writes updated master_state.json correctly."""
    test_json_file = tmp_path / "master_state.json"
    test_history_dir = tmp_path / "master_history"

    initial_data = {
        "version": "20260813_100000_000000",
        "updated_at": "2026-08-13T10:00:00",
        "rows": [
            {
                "姓名": "赵六",
                "身份证号": "110101199202025555",
                "手机号": "13700137000",
                "工种": "电工",
            }
        ],
        "last_changes": {},
    }
    test_json_file.write_text(json.dumps(initial_data, ensure_ascii=False), encoding="utf-8")
    test_json_repo = JsonWorkerRepository(file_path=test_json_file)

    incoming_df = pd.DataFrame([
        {
            "姓名": "赵六",
            "身份证号": "110101199202025555",
            "手机号": "13700137999",
            "工种": "高级电工",
        },
        {
            "姓名": "孙七",
            "身份证号": "110101199303036666",
            "手机号": "13600136000",
            "工种": "水暖工",
        }
    ])

    with patch.object(master_data, "STATE_FILE", test_json_file), \
         patch.object(master_data, "HISTORY_DIR", test_history_dir), \
         patch.object(master_data, "json_repository", test_json_repo), \
         patch.object(settings, "WORKER_STORAGE", "json"):
        result = master_data.commit_update(incoming_df, source_files=["test_import.xlsx"])
        assert result.get("error") == ""
        assert test_json_file.exists()

        saved_content = json.loads(test_json_file.read_text(encoding="utf-8"))
        assert "rows" in saved_content
        assert len(saved_content["rows"]) == 2

        ids = {r["身份证号"] for r in saved_content["rows"]}
        assert "110101199202025555" in ids
        assert "110101199303036666" in ids

        history_files = list(test_history_dir.glob("*.json"))
        assert len(history_files) == 1
