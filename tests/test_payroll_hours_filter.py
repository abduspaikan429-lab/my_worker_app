import unittest
from services.payroll_calculation_service import filter_top_working_days

class TestPayrollHoursFilter(unittest.TestCase):
    def test_filter_top_working_days(self):
        # 构造29天打卡工时测试数据
        daily_records = {
            1: 8.0, 2: 9.0, 3: 10.0, 4: 2.0, 5: 3.0,
            6: 8.0, 7: 8.0, 8: 8.0, 9: 8.0, 10: 8.0,
            11: 8.0, 12: 8.0, 13: 8.0, 14: 8.0, 15: 8.0,
            16: 8.0, 17: 8.0, 18: 8.0, 19: 8.0, 20: 8.0,
            21: 8.0, 22: 8.0, 23: 8.0, 24: 8.0, 25: 8.0,
            26: 8.0, 27: 8.0, 28: 8.0, 29: 8.0
        }

        # 验证指定 21 天，是否挑选时长最长的前 21 天
        selected, rejected = filter_top_working_days(daily_records, 21.0)
        
        self.assertEqual(len(selected), 21)
        self.assertEqual(len(rejected), 8)
        
        # 确认时长较短的几天(如2, 3小时的4号和5号)被淘汰放入差额池
        self.assertIn(4, rejected)
        self.assertIn(5, rejected)
        self.assertNotIn(4, selected)

if __name__ == '__main__':
    unittest.main()
