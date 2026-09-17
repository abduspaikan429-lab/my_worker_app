import json
from pathlib import Path
from unittest.mock import patch
import pytest

from services.onboarding_service import (
    OnboardingService,
    PAPER_ITEMS,
    ACCESS_ITEMS,
    PHOTO_ITEMS,
    SYSTEM_ITEMS,
    TOTAL_ITEMS,
    is_item_done,
    are_all_items_done,
)
from services.task_engine import get_onboarding_status


def test_constants_definitions():
    """验证核对清单项数量与内容是否与用户工作规约一致。"""
    assert len(PAPER_ITEMS) == 10
    assert "简易合同" in PAPER_ITEMS
    assert "体检单" in PAPER_ITEMS
    assert "三级教育" in PAPER_ITEMS
    assert "承诺书" in PAPER_ITEMS
    assert "进场承诺书" in PAPER_ITEMS
    assert "岗前培训试题" in PAPER_ITEMS
    assert "签到表按手印(2张)" in PAPER_ITEMS
    assert "花名册" in PAPER_ITEMS
    assert "退场承诺书" in PAPER_ITEMS
    assert "离场结算单" in PAPER_ITEMS

    assert len(ACCESS_ITEMS) == 5
    assert "门禁录入完成" in ACCESS_ITEMS
    assert "扫进场码" in ACCESS_ITEMS
    assert "浙里办-工人保障在线" in ACCESS_ITEMS
    assert "百工聚加卡签合同" in ACCESS_ITEMS
    assert "人社小灵光签合同" in ACCESS_ITEMS

    assert len(PHOTO_ITEMS) == 4
    assert "手持身份证+工资卡" in PHOTO_ITEMS
    assert "手持合同封面照" in PHOTO_ITEMS
    assert "手持合同日工资页照" in PHOTO_ITEMS
    assert "手持合同签字页照" in PHOTO_ITEMS

    assert len(SYSTEM_ITEMS) == 3
    assert TOTAL_ITEMS == 22


def test_backward_compatibility_with_legacy_keys():
    """验证旧版本字段的向下兼容与别名映射。"""
    legacy_data = {
        "paper": {
            "劳动合同(纸质)": True,  # 别名映射到 "简易合同"
            "岗前培训": True,        # 别名映射到 "岗前培训试题"
            "进场告知书(纸质)": True, # 别名映射到 "进场承诺书"
            "体检单": True,
        },
        "access": {
            "百工聚合同/告知书签署及加卡": True, # 别名映射到 "百工聚加卡签合同"
            "智慧护薪合同发起及工人/班组长确认": True, # 别名映射到 "人社小灵光签合同"
        },
        "system": {
            "更新月更报表": True, # 别名映射到 "更新变更月报"
        },
    }

    assert is_item_done(legacy_data["paper"], "简易合同") is True
    assert is_item_done(legacy_data["paper"], "岗前培训试题") is True
    assert is_item_done(legacy_data["paper"], "进场承诺书") is True
    assert is_item_done(legacy_data["paper"], "体检单") is True
    assert is_item_done(legacy_data["paper"], "离场结算单") is False

    assert is_item_done(legacy_data["access"], "百工聚加卡签合同") is True
    assert is_item_done(legacy_data["access"], "人社小灵光签合同") is True

    assert is_item_done(legacy_data["system"], "更新变更月报") is True

    service = OnboardingService()
    c, t = service.get_progress(legacy_data)
    assert t == 22
    assert c == 7  # 4 paper + 2 access + 1 system


def test_create_new_worker_structure(tmp_path: Path):
    """验证新工人创建时字段结构完整性。"""
    test_file = tmp_path / "onboarding_test.json"
    service = OnboardingService(file_path=test_file)

    rec = service.create_onboarding({
        "info": {
            "姓名": "测试工人A",
            "班组": "汪佩沾",
            "身份证号": "330101199001011234",
            "手机号": "13800001111",
        }
    })

    assert rec["info"]["姓名"] == "测试工人A"
    assert len(rec["paper"]) == 10
    assert len(rec["access"]) == 5
    assert len(rec["photo"]) == 4
    assert len(rec["system"]) == 3

    c, t = service.get_progress(rec)
    assert t == 22
    assert c == 0


def test_task_engine_onboarding_next_action():
    """验证任务引擎在面对未齐备材料时的唯一步骤提示。"""
    master_ids = set()
    master_name_phones = set()
    master_name_teams = set()

    worker_data = {
        "info": {"姓名": "李四", "班组": "王宜强"},
        "paper": {k: False for k in PAPER_ITEMS},
        "access": {k: False for k in ACCESS_ITEMS},
        "photo": {k: False for k in PHOTO_ITEMS},
        "system": {k: False for k in SYSTEM_ITEMS},
    }

    # 1. 缺纸质资料，第一步提示缺资料
    res = get_onboarding_status("李四_王宜强", worker_data, master_ids, master_name_phones, master_name_teams)
    assert res["category"] == "red"
    assert "补齐资料: 简易合同" in res["action"]

    # 2. 补齐纸质资料后，提示门禁
    for k in PAPER_ITEMS:
        worker_data["paper"][k] = True
    res = get_onboarding_status("李四_王宜强", worker_data, master_ids, master_name_phones, master_name_teams)
    assert res["category"] == "orange"
    assert res["action"] == "等待录入门禁"
