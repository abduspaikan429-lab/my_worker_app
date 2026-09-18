# tests/test_payroll_full_pipeline.py
import os
import shutil
import tempfile
import unittest
import pandas as pd
from pathlib import Path

from services.payroll_calculation_service import (
    filter_top_working_days,
    calculate_budget_allocation,
    number_to_rmb_capital,
    add_to_backlog,
    get_backlog_for_company,
    extract_from_backlog,
)


class TestPayrollFullPipeline(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_number_to_rmb_capital(self):
        self.assertEqual(number_to_rmb_capital(0), "零元整")
        self.assertEqual(number_to_rmb_capital(125000), "壹拾贰万伍仟元整")
        self.assertEqual(number_to_rmb_capital(135000.5), "壹拾叁万伍仟元伍角")
        self.assertEqual(number_to_rmb_capital(1002.34), "壹仟零贰元叁角肆分")

    def test_strict_half_day_allocation(self):
        workers = pd.DataFrame([
            {'姓名': '张三', '提取日薪': 300.0, '最终核定天数': 25.0, '锁定': False},
            {'姓名': '李四', '提取日薪': 350.0, '最终核定天数': 22.0, '锁定': False},
            {'姓名': '王五', '提取日薪': 400.0, '最终核定天数': 20.0, '锁定': True},
        ])
        # 原总额: 300*25 + 350*22 + 400*20 = 7500 + 7700 + 8000 = 23200
        # 目标额度: 18000
        alloc = calculate_budget_allocation(workers, 18000.0, strict_half_days=True)

        # 验证王五被锁定，天数和金额保持原样
        wangwu = alloc[alloc['姓名'] == '王五'].iloc[0]
        self.assertEqual(wangwu['实发天数'], 20.0)
        self.assertEqual(wangwu['实发工资'], 8000.0)

        # 验证所有人的天数都是 0.5 的整数倍
        for _, r in alloc.iterrows():
            self.assertTrue((r['实发天数'] * 2) % 1 == 0, f"天数非0.5整倍数: {r['实发天数']}")

        # 验证总实发不超过总额度
        self.assertLessEqual(alloc['实发工资'].sum(), 18000.0)



    def test_backlog_lifecycle(self):
        comp = "青海久昌建筑装饰工程有限公司"
        period = "6、7月"

        # 添加差额
        add_to_backlog(
            period=period,
            company=comp,
            worker_id="110101199001011234",
            worker_name="李四",
            days_deducted=3.0,
            amount_deducted=900.0,
            rejected_days={15: 8.0, 16: 7.5, 17: 8.0},
            daily_rate=300.0
        )

        backlog = get_backlog_for_company(comp)
        self.assertIn(period, backlog)
        self.assertIn("110101199001011234", backlog[period])
        self.assertEqual(backlog[period]["110101199001011234"]["欠额"], 900.0)

        # 提取核销
        extracted = extract_from_backlog(comp, period, ["110101199001011234"])
        self.assertEqual(len(extracted), 1)
        self.assertEqual(extracted[0]['姓名'], "李四")

        # 再次查询应已核销
        after_backlog = get_backlog_for_company(comp)
        self.assertNotIn(period, after_backlog)


if __name__ == '__main__':
    unittest.main()
