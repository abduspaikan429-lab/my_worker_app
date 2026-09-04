import json
from pathlib import Path
from unittest.mock import patch
import pandas as pd
import pytest

from modules import master_data, info_merge
from modules.attendance_payroll import _load_master_data, enrich_with_master
from services.onboarding_service import OnboardingService


def test_upsert_master_workers_and_onboarding_sync(tmp_path: Path):
    """测试将新进场人员写入主表，并验证 load_master_df 与 attendance_payroll 能正确读取日薪。"""
    test_json_file = tmp_path / "master_state.json"
    test_history_dir = tmp_path / "master_history"
    test_onboarding_file = tmp_path / "onboarding_data.json"

    initial_state = {
        "version": "20260830_000000_000000",
        "updated_at": "2026-08-30T00:00:00",
        "rows": [
            {
                "姓名": "老员工A",
                "班组": "汪佩沾",
                "身份证号": "110101198001011111",
                "结算单价/标准": "基本工资：400元/天",
                "工种": "木工",
            }
        ],
        "last_changes": {},
    }
    test_json_file.write_text(json.dumps(initial_state, ensure_ascii=False), encoding="utf-8")

    # 模拟从官网导出的 6 个新进场人员
    new_incoming_workers = [
        {
            "姓名": f"新工人{i}",
            "班组": "王宜强施工班组",
            "身份证号": f"34010119900101000{i}",
            "手机号": f"1380000000{i}",
            "工种": "焊工",
            "结算单价/标准": "基本工资：450元/天",
            "工资卡号": f"622202000000000{i}",
            "开户银行": "中国农业银行",
            "进场日期": "2026-09-01",
        }
        for i in range(1, 7)
    ]

    with patch.object(master_data, "STATE_FILE", test_json_file), \
         patch.object(master_data, "HISTORY_DIR", test_history_dir):
        # 1. 增量写入主表
        result = master_data.upsert_master_workers(new_incoming_workers, source="test")
        assert result["added"] == 6
        assert result["total"] == 7

        # 2. 验证 load_master_df 包含 7 个人
        master_df = master_data.load_master_df()
        assert len(master_df) == 7
        assert "新工人1" in master_df["姓名"].values
        assert "新工人6" in master_df["姓名"].values

        # 3. 验证 _load_master_data 正确提取有效日薪 450
        loaded_master = _load_master_data()
        assert loaded_master is not None
        assert len(loaded_master) == 7
        assert "提取日薪" in loaded_master.columns
        new_w1 = loaded_master[loaded_master["姓名"] == "新工人1"].iloc[0]
        assert new_w1["提取日薪"] == 450.0

        # 4. 验证考勤定稿做工资表时能正确匹配日薪，无缺失
        attendance_df = pd.DataFrame([
            {"姓名": f"新工人{i}", "身份证号": f"34010119900101000{i}", "最终核定天数": 20}
            for i in range(1, 7)
        ])
        enriched = enrich_with_master(attendance_df)
        assert len(enriched) == 6
        assert (enriched["提取日薪"] == 450.0).all()


def test_info_merge_auto_sync_to_master(tmp_path: Path):
    """测试档案整合板块清洗后自动全量提交写入主表，并在考勤工资模块正常提取日薪。"""
    test_json_file = tmp_path / "master_state.json"
    test_history_dir = tmp_path / "master_history"

    # 模拟官网导出的 6 个新进场人员整合结果
    incoming_df = pd.DataFrame([
        {
            "姓名": f"官网进场工人{i}",
            "班组": "汪佩沾",
            "身份证号": f"42010119920101000{i}",
            "手机号": f"1390000000{i}",
            "工种": "安装工",
            "结算单价/标准": "基本工资：380元/天",
            "工资卡号": f"622848000000000{i}",
            "开户银行": "中国建设银行",
            "分包/所属企业": "江苏旭之升",
        }
        for i in range(1, 7)
    ])

    test_master_dir = tmp_path / "master"
    test_master_dir.mkdir(parents=True, exist_ok=True)

    with patch.object(master_data, "STATE_FILE", test_json_file), \
         patch.object(master_data, "HISTORY_DIR", test_history_dir), \
         patch.object(master_data, "MASTER_DIR", test_master_dir):
        saved = master_data.commit_update(incoming_df, source_files=["官方导出.xlsx"])
        assert saved.get("error") == ""
        assert len(saved["new_rows"]) == 6

        # 考勤工资模块读取主表并核算
        loaded_master = _load_master_data()
        assert len(loaded_master) == 6
        assert (loaded_master["提取日薪"] == 380.0).all()
        assert "官网进场工人1" in loaded_master["姓名"].values



def test_process_and_merge_with_synonyms(tmp_path: Path):
    """测试多平台导出的不同表头同义字段（日薪、银行账号、开户行等）清洗合并与主表全量更新。"""
    # 模拟三局系统导出 Excel
    df_sanju = pd.DataFrame([
        {
            "姓名": "张三",
            "所属班组": "汪佩沾",
            "身份证号码": "340101198501011234",
            "联系电话": "13811112222",
            "岗位": "木工",
            "劳务公司": "江苏旭之升",
        }
    ])
    f_sanju = tmp_path / "sanju.xlsx"
    df_sanju.to_excel(f_sanju, index=False)

    # 模拟智慧护薪导出 Excel（含日薪和银行卡）
    df_huxin = pd.DataFrame([
        {
            "姓名": "张三",
            "身份证号": "340101198501011234",
            "日薪": "420",
            "银行卡号": "6228480000001111",
            "开户行": "中国农业银行",
            "合同状态": "已签订",
        }
    ])
    f_huxin = tmp_path / "huxin.xlsx"
    df_huxin.to_excel(f_huxin, index=False)

    with open(f_sanju, "rb") as f_a, open(f_huxin, "rb") as f_b:
        merged_df, count_a, overlap, count_b = info_merge.process_and_merge([f_a], [f_b])

    assert merged_df is not None
    assert len(merged_df) == 1
    row = merged_df.iloc[0]
    assert row["姓名"] == "张三"
    assert row["班组"] == "汪佩沾"
    assert row["身份证号"] == "340101198501011234"
    assert row["手机号"] == "13811112222"
    assert row["工种"] == "木工"
    assert row["结算单价/标准"] == "420"
    assert row["工资卡号"] == "6228480000001111"
    assert row["开户银行"] == "中国农业银行"
    assert row["合同签订状态"] == "已签订"


def test_enrich_with_master_without_id_or_cached_zeros():
    """验证当考勤表缺失身份证号、或者此前定稿日薪为0/空时，也能100%按姓名/班组权威补齐日薪。"""
    # 模拟考勤表：只有姓名和天数，甚至没有身份证号
    attendance_df = pd.DataFrame([
        {"姓名": "张建平", "最终核定天数": 25},
        {"姓名": "聂超峰", "最终核定天数": 20},
        {"姓名": "王明双", "最终核定天数": 18},
        {"姓名": "王平", "最终核定天数": 22},
        {"姓名": "王义全", "最终核定天数": 20},
        {"姓名": "赵敏", "最终核定天数": 26},
    ])

    enriched = enrich_with_master(attendance_df)
    assert len(enriched) == 6
    assert (enriched["提取日薪"] > 0).all()
    assert (enriched["身份证号"].str.len() >= 15).all()

    # 验证提取到的具体日薪
    wages = dict(zip(enriched["姓名"], enriched["提取日薪"]))
    assert wages["张建平"] == 450.0
    assert wages["聂超峰"] == 380.0
    assert wages["王明双"] == 450.0
    assert wages["王平"] == 450.0
    assert wages["王义全"] == 500.0
    assert wages["赵敏"] == 450.0


