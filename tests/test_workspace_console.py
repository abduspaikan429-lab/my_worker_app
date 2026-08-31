# tests/test_workspace_console.py
import pytest
import pandas as pd
from services.task_engine import get_all_tasks, get_onboarding_status, get_offboarding_status
from modules.workspace_console import get_tasks

def test_get_onboarding_status_red():
    data = {
        "info": {"姓名": "张三", "班组": "木工班", "身份证号": "110101199001011234"},
        "paper": {"身份证复印件": False},
        "system": {},
        "access": {},
        "status": "active"
    }
    status = get_onboarding_status("worker_1", data, set(), set(), set())
    assert status["category"] == "red"
    assert "补齐资料" in status["action"]
    assert status["type"] == "onboarding"

def test_get_offboarding_status():
    data = {
        "info": {"姓名": "李四", "班组": "钢筋班"},
        "steps": {"1. 工人在百工聚发起离场": False},
        "paper": {},
        "system": {},
        "status": "active"
    }
    status = get_offboarding_status("worker_2", data)
    assert status["category"] == "orange"
    assert status["type"] == "offboarding"

def test_get_all_tasks():
    onboarding_data = {
        "w1": {
            "info": {"姓名": "张三", "班组": "木工班"},
            "paper": {"身份证复印件": False},
            "system": {},
            "access": {},
            "status": "active"
        }
    }
    offboarding_data = {}
    master_df = pd.DataFrame()
    red, orange, yellow, anomalies = get_all_tasks(onboarding_data, offboarding_data, master_df)
    assert len(red) == 1
    assert red[0]["name"] == "张三"
