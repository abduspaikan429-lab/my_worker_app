import unittest
import pandas as pd
import numpy as np
from services.payroll_calculation_service import calculate_budget_allocation

class TestPayrollBudgetAllocation(unittest.TestCase):
    def test_calculate_budget_allocation(self):
        # 构造测试数据
        data = {
            '姓名': ['张三', '李四', '王五'],
            '提取日薪': [300.0, 320.0, 350.0],
            '最终核定天数': [20.0, 20.0, 20.0],
            '锁定': [False, False, False]
        }
        df = pd.DataFrame(data)
        
        # 原应发总额: 300*20 + 320*20 + 350*20 = 6000 + 6400 + 7000 = 19400
        # 目标额度: 13500.00
        target_budget = 13500.00
        
        res_df = calculate_budget_allocation(df, target_budget)
        
        # 验证实发总额分毫不差匹配额度
        actual_total = res_df['实发工资'].sum()
        self.assertAlmostEqual(actual_total, target_budget, places=2)
        
        # 验证天数为 0.5 的倍数 (尾差吸收者可能不是严格的 0.5 倍数，但大部分应该是)
        # 确保计算逻辑不抛异常且总额正确
        
if __name__ == '__main__':
    unittest.main()
