# tests/test_payroll_sanju_reconciliation.py
import os
import unittest
import openpyxl
import pandas as pd
from services.payroll_calculation_service import (
    parse_sanju_punch_excel,
    reconcile_attendance,
    filter_top_working_days,
    calculate_budget_allocation,
)
from modules.attendance_payroll import (
    _populate_company_attendance_sheet,
    _populate_zongbao_attendance_sheet,
)


class TestPayrollSanjuReconciliation(unittest.TestCase):
    def setUp(self):
        self.excel_file = r'D:\code\工作脚本\my_worker_app\load-data\工人出勤明细-20260917.xlsx'

    def test_parse_sanju_punch_excel(self):
        if not os.path.exists(self.excel_file):
            self.skipTest("Sample file not found")

        df, period = parse_sanju_punch_excel(self.excel_file)
        self.assertIn("08月", period)
        self.assertGreater(len(df), 30)

        # 验证包含刘灯高
        ldg = df[df['姓名'] == '刘灯高']
        self.assertEqual(len(ldg), 1)
        ldg_row = ldg.iloc[0]
        self.assertEqual(ldg_row['总包系统天数'], 20.0)
        self.assertGreater(ldg_row['总工时'], 160.0)
        self.assertIn(1, ldg_row['__考勤明细__'])

    def test_reconcile_attendance_rules(self):
        # 构造测试数据
        system_df = pd.DataFrame([
            {
                '姓名': '张三',
                '身份证号': '110101199001011234',
                '分包/所属企业': '江苏旭之升建筑工程有限公司',
                '总包系统天数': 21.0,
                '总工时': 180.0,
                '平均工时': 8.57,
                '__考勤明细__': {i: 8.5 for i in range(1, 22)},
            },
            {
                '姓名': '李四',
                '身份证号': '110101199001015678',
                '分包/所属企业': '江苏旭之升建筑工程有限公司',
                '总包系统天数': 28.0,
                '总工时': 230.0,
                '平均工时': 8.21,
                # 前18天工时8~10h，后10天为混打卡0.5~2h
                '__考勤明细__': {**{i: 9.0 for i in range(1, 19)}, **{i: 1.0 for i in range(19, 29)}},
            }
        ])

        # 模拟：张三在我这里29天，班组长25天，系统21天 -> 应裁定21天
        # 李四在我这里29天，班组长18天，系统28天 -> 应裁定18天
        foreman_map = {
            '张三': 25.0,
            '李四': 18.0
        }

        reconciled = reconcile_attendance(system_df, foreman_days_map=foreman_map)

        # 1. 验证张三
        zs = reconciled[reconciled['姓名'] == '张三'].iloc[0]
        self.assertEqual(zs['最终裁定工日'], 21.0)
        self.assertIn("受限系统打卡", zs['裁定状态'])
        self.assertEqual(len(zs['__考勤明细__']), 21)

        # 2. 验证李四
        ls = reconciled[reconciled['姓名'] == '李四'].iloc[0]
        self.assertEqual(ls['最终裁定工日'], 18.0)
        self.assertIn("班组长核减10.0天", ls['裁定状态'])
        selected_dates = set(ls['__考勤明细__'].keys())
        self.assertEqual(selected_dates, set(range(1, 19)))
        rejected_dates = set(ls['_rejected_foreman'].keys())
        self.assertEqual(rejected_dates, set(range(19, 29)))

    def test_reconcile_with_my_df_containing_dict_column(self):
        """验证 my_df 包含字典列（如 __考勤明细__）时不会触发 TypeError: float() argument must be a string or a real number, not 'dict'"""
        system_df = pd.DataFrame([{
            '姓名': '张三',
            '身份证号': '110101199001011234',
            '分包/所属企业': '江苏旭之升建筑工程有限公司',
            '总包系统天数': 20.0,
            '总工时': 160.0,
            '平均工时': 8.0,
            '__考勤明细__': {i: 8.0 for i in range(1, 21)},
        }])

        my_df = pd.DataFrame([{
            '分包/所属企业': '江苏旭之升建筑工程有限公司',
            '姓名': '张三',
            '身份证号': '110101199001011234',
            '月份': '8月',
            '__考勤明细__': {1: '✓', 2: '✓', 3: '✓'},  # 字典列
            '解析出勤天数': 25.0,
            '最终核定天数': 25.0,
        }])

        reconciled = reconcile_attendance(system_df, my_df=my_df)
        zs = reconciled.iloc[0]
        self.assertEqual(zs['我的考勤天数'], 25.0)
        # min(25, 20) = 20
        self.assertEqual(zs['最终裁定工日'], 20.0)

    def test_reconcile_preserves_raw_punch_hours_and_drawer_safe(self):
        """验证 _raw_punch_hours 得到保留，且工时抽屉解析不会因 '✓' 或 '0.5' 等字符串抛出 ValueError"""
        system_df = pd.DataFrame([{
            '姓名': '李四',
            '身份证号': '110101199001015678',
            '分包/所属企业': '青海久昌建筑装饰工程有限公司',
            '总包系统天数': 3.0,
            '总工时': 24.0,
            '平均工时': 8.0,
            '_raw_punch_hours': {1: 9.5, 2: 10.0, 3: 0.5},
            '__考勤明细__': {1: 9.5, 2: 10.0, 3: 0.5},
        }])

        reconciled = reconcile_attendance(system_df, foreman_days_map={'李四': 2.0})
        ls = reconciled.iloc[0]
        self.assertEqual(ls['最终裁定工日'], 2.0)
        # 1号和2号工时长被选中，值是 '✓'
        self.assertEqual(ls['__考勤明细__'][1], '✓')
        self.assertEqual(ls['__考勤明细__'][2], '✓')
        # _raw_punch_hours 仍保留真实小时数
        self.assertIn(1, ls['_raw_punch_hours'])
        self.assertEqual(ls['_raw_punch_hours'][1], 9.5)

    def test_filter_zero_attendance_from_sheets(self):
        """测试 1: 考勤为 0 的工人绝对不被写入公司与总包考勤表中"""
        df = pd.DataFrame([
            {
                '姓名': '张三', '工种': '木工', '性别': '男',
                '__考勤明细__': {1: '✓', 2: '✓'},
                '最终核定天数': 2.0,
            },
            {
                '姓名': '李四(0考勤)', '工种': '杂工', '性别': '男',
                '__考勤明细__': {},
                '最终核定天数': 0.0,
            },
            {
                '姓名': '王五', '工种': '瓦工', '性别': '男',
                '__考勤明细__': {3: '✓'},
                '最终核定天数': 1.0,
            }
        ])

        batch = {
            'company': '青海久昌建筑装饰工程有限公司',
            'period': '8月',
            'df': df
        }

        # 1. 公司标准考勤表
        wb_comp = openpyxl.Workbook()
        ws_comp = wb_comp.active
        ws_comp.title = "8月青海久昌"
        for c in range(1, 36):
            ws_comp.cell(row=6, column=c, value='')
        res_comp = _populate_company_attendance_sheet(ws_comp, batch)
        workers_in_comp = [ws_comp.cell(r, 2).value for r in range(res_comp['start_row'], res_comp['start_row'] + 2)]
        self.assertIn('张三', workers_in_comp)
        self.assertIn('王五', workers_in_comp)
        self.assertNotIn('李四(0考勤)', workers_in_comp)

        # 2. 总包标准考勤表
        wb_zb = openpyxl.Workbook()
        ws_zb = wb_zb.active
        ws_zb.title = "Sheet1"
        res_zb = _populate_zongbao_attendance_sheet(ws_zb, batch)
        workers_in_zb = [ws_zb.cell(r, 2).value for r in range(res_zb['start_row'], res_zb['start_row'] + 2)]
        self.assertIn('张三', workers_in_zb)
        self.assertIn('王五', workers_in_zb)
        self.assertNotIn('李四(0考勤)', workers_in_zb)

    def test_allowance_in_calculation_and_budget(self):
        """测试 2: 工资中增加津贴列，联动应发与实发计算"""
        system_df = pd.DataFrame([{
            '姓名': '班组长老赵',
            '身份证号': '110101198501011111',
            '分包/所属企业': '江苏旭之升',
            '工种': '班组长',
            '总包系统天数': 20.0,
            '总工时': 180.0,
            '平均工时': 9.0,
            '提取日薪': 400.0,
            '津贴': 1500.0,
            '__考勤明细__': {i: 9.0 for i in range(1, 21)},
        }])

        reconciled = reconcile_attendance(system_df, allowance_map={'班组长老赵': 1500.0})
        row = reconciled.iloc[0]
        self.assertEqual(row['应发工资'], 9500.0)
        self.assertEqual(row['津贴'], 1500.0)

        alloc_df = calculate_budget_allocation(reconciled, total_budget=5500.0, strict_half_days=True)
        r_alloc = alloc_df.iloc[0]
        self.assertEqual(r_alloc['实发工资'], round(r_alloc['实发天数'] * 400.0 + 1500.0, 2))

    def test_sufficient_budget_no_reduction(self):
        """测试 3: 增加额度充足时，不需要反推缩减，全额足额发放"""
        workers_df = pd.DataFrame([
            {'姓名': '甲', '提取日薪': 300.0, '最终核定天数': 20.0, '津贴': 500.0},
            {'姓名': '乙', '提取日薪': 400.0, '最终核定天数': 15.0, '津贴': 0.0},
        ])

        alloc_df = calculate_budget_allocation(workers_df, total_budget=13000.0)
        self.assertEqual(alloc_df.loc[0, '实发天数'], 20.0)
        self.assertEqual(alloc_df.loc[0, '实发工资'], 6500.0)
        self.assertEqual(alloc_df.loc[1, '实发天数'], 15.0)
        self.assertEqual(alloc_df.loc[1, '实发工资'], 6000.0)

    def test_final_days_override_and_special_status(self):
        """测试 4: 最终裁定工日支持人工特批修改，不写死；且超额特批时自动在空缺日补齐打钩"""
        system_df = pd.DataFrame([{
            '姓名': '张师傅',
            '身份证号': '110101198001012222',
            '分包/所属企业': '青海久昌',
            '总包系统天数': 25.0,
            '总工时': 225.0,
            '平均工时': 9.0,
            '_raw_punch_hours': {i: 9.0 for i in range(1, 26)},
            '__考勤明细__': {i: 9.0 for i in range(1, 26)},
        }])

        reconciled = reconcile_attendance(
            system_df,
            foreman_days_map={'张师傅': 28.0},
            final_days_map={'张师傅': 28.0}
        )
        row = reconciled.iloc[0]
        self.assertEqual(row['最终裁定工日'], 28.0)
        self.assertIn('特殊特批工日', row['裁定状态'])
        self.assertIn('超出系统3.0天', row['裁定状态'])

        selected_marks = row['__考勤明细__']
        self.assertEqual(len(selected_marks), 28)
        self.assertTrue(all(v == '✓' for v in selected_marks.values()))


if __name__ == '__main__':
    unittest.main()


