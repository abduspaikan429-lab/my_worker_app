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
                '有效工天': 21.0,
                '总工时': 180.0,
                '平均工时': 8.57,
                '__考勤明细__': {i: 8.5 for i in range(1, 22)},
            },
            {
                '姓名': '李四',
                '身份证号': '110101199001015678',
                '分包/所属企业': '江苏旭之升建筑工程有限公司',
                '总包系统天数': 28.0,
                '有效工天': 28.0,
                '总工时': 230.0,
                '平均工时': 8.21,
                # 前18天工时8~10h，后10天为混打卡0.5~2h
                '__考勤明细__': {**{i: 9.0 for i in range(1, 19)}, **{i: 1.0 for i in range(19, 29)}},
            }
        ])

        # 模拟：以水印照片/班组长为主
        # 张三：水印/班组长25天，系统有效工天21天 -> 以水印为主裁定25天，正常打卡21天，阶段性结算4天
        # 李四：系统有效工天28天，水印/班组长18天 -> 以水印为主核减为18天，淘汰短工时
        foreman_map = {
            '张三': 25.0,
            '李四': 18.0
        }

        reconciled = reconcile_attendance(system_df, foreman_days_map=foreman_map)

        # 1. 验证张三（系统小于水印 -> 以水印为主，走阶段性结算）
        zs = reconciled[reconciled['姓名'] == '张三'].iloc[0]
        self.assertEqual(zs['最终裁定工日'], 25.0)
        self.assertEqual(zs['正常打卡天数'], 21.0)
        self.assertEqual(zs['阶段性结算天数'], 4.0)
        self.assertIn("4.0", zs['裁定状态'])
        self.assertIn("走阶段性结算", zs['裁定状态'])
        self.assertEqual(len(zs['__考勤明细__']), 25)

        # 2. 验证李四（系统大于水印 -> 以水印为主核减，剔除短工时）
        ls = reconciled[reconciled['姓名'] == '李四'].iloc[0]
        self.assertEqual(ls['最终裁定工日'], 18.0)
        self.assertEqual(ls['正常打卡天数'], 18.0)
        self.assertEqual(ls['阶段性结算天数'], 0.0)
        self.assertIn("水印核减 10.0 天", ls['裁定状态'])
        selected_dates = set(ls['__考勤明细__'].keys())
        self.assertEqual(selected_dates, set(range(1, 19)))
        rejected_dates = set(ls['_rejected_foreman'].keys())
        self.assertEqual(rejected_dates, set(range(19, 29)))

    def test_reconcile_with_my_df_containing_dict_column(self):
        """验证 my_df 包含字典列（如 __考勤明细__）且系统小于水印时以水印为主走阶段性结算"""
        system_df = pd.DataFrame([{
            '姓名': '张三',
            '身份证号': '110101199001011234',
            '分包/所属企业': '江苏旭之升建筑工程有限公司',
            '总包系统天数': 20.0,
            '有效工天': 20.0,
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
        self.assertEqual(zs['最终裁定工日'], 25.0)
        self.assertEqual(zs['正常打卡天数'], 20.0)
        self.assertEqual(zs['阶段性结算天数'], 5.0)
        self.assertIn("5.0", zs['裁定状态'])
        self.assertIn("走阶段性结算", zs['裁定状态'])


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
        """测试 4: 最终裁定工日支持人工特批修改，超出系统打卡时自动在空缺日补齐打钩并标记阶段性结算"""
        system_df = pd.DataFrame([{
            '姓名': '张师傅',
            '身份证号': '110101198001012222',
            '分包/所属企业': '青海久昌',
            '总包系统天数': 25.0,
            '有效工天': 25.0,
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
        self.assertEqual(row['正常打卡天数'], 25.0)
        self.assertEqual(row['阶段性结算天数'], 3.0)
        self.assertIn('3.0', row['裁定状态'])
        self.assertIn('走阶段性结算', row['裁定状态'])

        selected_marks = row['__考勤明细__']
        self.assertEqual(len(selected_marks), 28)
        self.assertTrue(all(v == '✓' for v in selected_marks.values()))

    def test_wide_table_effective_days_and_stage_settlement(self):
        """测试 5: 用户最新横向宽表样式解析（严格按有效工天算）以及水印照片核对与阶段性结算"""
        import io
        csv_data = '''2026年09月01 - 2026年09月30出勤情况  ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
姓名,性别,年龄,身份证号,参建单位,2026年09月,2日,3日,4日,5日,6日,7日,8日,9日,10日,11日,12日,13日,14日,15日,16日,17日,18日,19日,20日,21日,22日,23日,24日,25日,26日,27日,28日,29日,30日,出勤天数,有效工天,总工时
姓名,性别,年龄,身份证号,参建单位,1日,2日,3日,4日,5日,6日,7日,8日,9日,10日,11日,12日,13日,14日,15日,16日,17日,18日,19日,20日,21日,22日,23日,24日,25日,26日,27日,28日,29日,30日,出勤天数,有效工天,总工时
李瑞香,女,48,140723197808180047,青海久昌建筑装饰工程有限公司,10.37,5.31,5.33,10.30,10.67,5.09,10.01,10.05,10.09,10.25,10.25,9.78,9.95,10.03,10.23,9.96,9.90,9.89,10.06,10.35,10.14,10.04,10.34,12.79,0.01,8.17,9.12,9.90,11.50,10.06,30.0,28.0,279.94
王明凤,女,23,500222200306206647,青海久昌建筑装饰工程有限公司,9.95,9.67,10.05,9.87,8.96,请假,8.79,--,请假,9.62,10.05,请假,9.83,10.00,9.95,10.00,10.04,10.35,9.81,请假,9.86,9.84,8.63,9.93,请假,--,请假,10.27,10.23,请假,21.0,21.0,205.70
杨科明,男,48,513027197804142632,青海久昌建筑装饰工程有限公司,10.16,5.36,2.90,13.19,10.12,5.34,10.34,11.36,12.30,13.08,1.42,10.06,10.41,10.29,10.01,10.05,10.25,请假,10.59,10.67,10.58,10.73,5.65,12.91,0.01,8.25,9.58,0.01,0.63,9.60,29.0,24.5,245.85
'''
        sys_df, period = parse_sanju_punch_excel(io.StringIO(csv_data.strip()))
        self.assertIn("09月", period)
        self.assertEqual(len(sys_df), 3)

        # 严格验证：总包系统天数按有效工天算！
        lrx = sys_df[sys_df['姓名'] == '李瑞香'].iloc[0]
        self.assertEqual(lrx['有效工天'], 28.0)
        self.assertEqual(lrx['总包系统天数'], 28.0)
        self.assertEqual(lrx['出勤天数'], 30.0)

        ykm = sys_df[sys_df['姓名'] == '杨科明'].iloc[0]
        self.assertEqual(ykm['有效工天'], 24.5)
        self.assertEqual(ykm['总包系统天数'], 24.5)

        # 构造水印照片表：
        # 1. 李瑞香：出勤30天，有效28天，水印30天 -> 水印未超打卡天数，阶段性结算0天，正常发放28天，最终核定28天
        # 2. 杨科明：出勤29天，有效24.5天，水印31天 -> 水印超打卡天数2天，阶段性结算2天，正常发放24.5天，最终核定26.5天
        # 3. 王明凤：出勤21天，有效21天，水印18天 -> 水印核减3天，正常发放18天，阶段性结算0天，最终核定18天
        my_df = pd.DataFrame([
            {'姓名': '李瑞香', '身份证号': '140723197808180047', '最终核定天数': 30.0},
            {'姓名': '杨科明', '身份证号': '513027197804142632', '最终核定天数': 31.0},
            {'姓名': '王明凤', '身份证号': '500222200306206647', '最终核定天数': 18.0},
        ])

        reconciled = reconcile_attendance(sys_df, my_df=my_df)
        
        # 验证李瑞香（水印 <= 打卡天数 -> 无阶段性结算，按有效工天28天正常发放）
        r_lrx = reconciled[reconciled['姓名'] == '李瑞香'].iloc[0]
        self.assertEqual(r_lrx['最终核定天数'], 28.0)
        self.assertEqual(r_lrx['正常打卡天数'], 28.0)
        self.assertEqual(r_lrx['阶段性结算天数'], 0.0)
        self.assertEqual(r_lrx['裁定状态'], "✔️ 正常有效工天发放")

        # 验证杨科明（水印 31 > 打卡天数 29 -> 超出2天走阶段性结算，有效工天24.5天正常发放）
        r_ykm = reconciled[reconciled['姓名'] == '杨科明'].iloc[0]
        self.assertEqual(r_ykm['阶段性结算天数'], 2.0)
        self.assertEqual(r_ykm['正常打卡天数'], 24.5)
        self.assertEqual(r_ykm['最终核定天数'], 26.5)
        self.assertIn("水印超打卡 2.0 天", r_ykm['裁定状态'])
        self.assertIn("走阶段性结算", r_ykm['裁定状态'])

        # 验证王明凤（水印 18 < 有效工天 21 -> 核减3天）
        r_wmf = reconciled[reconciled['姓名'] == '王明凤'].iloc[0]
        self.assertEqual(r_wmf['最终核定天数'], 18.0)
        self.assertEqual(r_wmf['正常打卡天数'], 18.0)
        self.assertEqual(r_wmf['阶段性结算天数'], 0.0)
        self.assertIn("水印核减 3.0 天", r_wmf['裁定状态'])


if __name__ == '__main__':
    unittest.main()



