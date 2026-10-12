import pandas as pd
import pytest
from io import BytesIO
from zipfile import ZipFile
from openpyxl import load_workbook

from modules.attendance_payroll import (
    parse_watermark_attendance,
    _attendance_state,
    build_all_exports_zip
)


def test_watermark_attendance_parsing_and_days_calculation():
    """Test that parse_watermark_attendance correctly calculates actual attendance days from daily columns."""
    df_raw = pd.DataFrame([
        {"姓名": "张三", "身份证号": "110101199001011234", "1日": "√", "2日": "√", "3日": "", "4日": "--"},
        {"姓名": "李四", "身份证号": "110101199505054321", "1日": "08:00-18:00", "2日": "请假", "3日": "出勤", "4日": "/"},
    ])
    
    res = parse_watermark_attendance(df_raw, "测试公司")
    
    assert len(res) == 2
    
    # 张三 should have 2 days ("√", "√")
    zhang_san = res[res['姓名'] == '张三'].iloc[0]
    assert zhang_san['解析出勤天数'] == 2.0
    assert zhang_san['最终核定天数'] == 2.0
    assert zhang_san['月份'] == ''
    assert zhang_san['__考勤明细__'][1] == '√'
    assert zhang_san['__考勤明细__'][2] == '√'
    
    # 李四 should have 2 days ("08:00-18:00", "出勤")
    li_si = res[res['姓名'] == '李四'].iloc[0]
    assert li_si['解析出勤天数'] == 2.0


def test_watermark_attendance_fallback_total_days():
    """Test that parse_watermark_attendance falls back to the total days column if daily records are not present."""
    df_raw = pd.DataFrame([
        {"姓名": "王五", "身份证号": "110101199202029876", "考勤天数": "25.5"},
        {"姓名": "赵六", "身份证号": "110101199303031111", "考勤天数": "15"},
    ])
    
    res = parse_watermark_attendance(df_raw, "测试公司")
    
    assert len(res) == 2
    
    wang_wu = res[res['姓名'] == '王五'].iloc[0]
    assert wang_wu['解析出勤天数'] == 25.5
    
    zhao_liu = res[res['姓名'] == '赵六'].iloc[0]
    assert zhao_liu['解析出勤天数'] == 15.0


def test_personnel_identification():
    """Test that personnel identification and deduplication works."""
    df_raw = pd.DataFrame([
        {"姓名": "孙七", "身份证号": "110101199404042222", "1日": "√"},
        {"姓名": "孙七", "身份证号": "110101199404042222", "1日": "√", "2日": "√"},  # Duplicate ID
        {"姓名": "周八", "身份证号": "", "1日": "√"},
        {"姓名": "周八", "身份证号": "", "1日": "√", "2日": "√"},  # Duplicate Name without ID
    ])
    
    res = parse_watermark_attendance(df_raw, "测试公司")
    
    # Should deduplicate based on ID or Name
    assert len(res) == 2
    
    sun_qi = res[res['姓名'] == '孙七'].iloc[0]
    assert sun_qi['解析出勤天数'] == 2.0  # Daily records are merged from both rows


def test_workflow_status_concept():
    """Test conceptual representation of draft and finalized states and payroll gating."""
    # This acts as a mock representation of the UI logic
    
    state = {'_att_status': None, 'final_attendance': None}
    
    # Step 1: Upload and parse -> draft
    state['_att_status'] = 'draft'
    state['_att_parsed_df'] = pd.DataFrame([{"姓名": "张三", "最终核定天数": 20}])
    
    assert state['_att_status'] == 'draft'
    
    # Check if payroll is allowed (simulating the gating logic in Tab 2)
    def can_calculate_payroll(st_state):
        return st_state.get('_att_status') == 'finalized' and st_state.get('final_attendance') is not None
    
    assert not can_calculate_payroll(state)
    
    # Step 2: Finalize
    state['_att_status'] = 'finalized'
    state['final_attendance'] = state['_att_parsed_df']
    
    assert state['_att_status'] == 'finalized'
    assert can_calculate_payroll(state)


def test_export_attendance_and_payroll():
    """Two companies are consolidated into four workbooks with daily marks preserved."""
    salary_df = pd.DataFrame([
        {
            "分包/所属企业": "江苏旭之升建筑工程有限公司",
            "姓名": "张三",
            "身份证号": "110101199001011234",
            "月份": "6月",
            "__考勤明细__": {1: "✓", 2: "✓"},
            "最终核定天数": 20,
            "日薪": 300,
            "应发工资": 6000,
            "工种": "木工",
            "联系电话": "13800000000",
            "开户银行": "工商银行",
            "银行卡号": "6222020000000000000",
            "性别": "男",
        },
        {
            "分包/所属企业": "青海久昌建筑装饰工程有限公司",
            "姓名": "李四",
            "身份证号": "110101199505054321",
            "月份": "6月",
            "__考勤明细__": {1: "8", 3: "出勤"},
            "最终核定天数": 18,
            "日薪": 400,
            "应发工资": 7200,
            "工种": "焊工",
            "联系电话": "13900000000",
            "开户银行": "农业银行",
            "银行卡号": "6222020000000000001",
            "性别": "男",
        },
    ])

    zip_data = build_all_exports_zip(salary_df, include_ledger=False)

    assert isinstance(zip_data, bytes)
    assert len(zip_data) > 0
    with ZipFile(BytesIO(zip_data)) as archive:
        names = archive.namelist()
        # 2个批次各生成1个包含考勤与工资确认表的Excel文件，共2个文件，不再包含总表
        assert len(names) == 2
        assert all(name.endswith('.xlsx') and '/' not in name for name in names)
        assert any('旭之升' in name and '-考勤与工资确认表.xlsx' in name for name in names)
        assert any('青海久昌' in name and '-考勤与工资确认表.xlsx' in name for name in names)
        # 严格验证：彻底删除以_连接的总包文件，且不包含工资卡台账
        assert not any('_考勤表.xlsx' in name for name in names)
        assert not any('_工资确认表.xlsx' in name for name in names)
        assert '农民工工资卡登记台账.xlsx' not in names

        combined_wb = load_workbook(
            BytesIO(archive.read(next(name for name in names if '旭之升' in name and '-考勤与工资确认表.xlsx' in name))),
            data_only=False,
        )
        # 1、工资表和考勤表放在一起，不要总表
        assert combined_wb.sheetnames == ['考勤表', '工资确认表']
        att_ws = combined_wb['考勤表']
        wage_ws = combined_wb['工资确认表']

        # 验证考勤表中的打钩已转换为数字8
        assert att_ws['D6'].value == 8
        # 3、考勤表在合计处填上公式以便合计=COUNTIF(D6:AH6,8)+COUNTIF(D6:AH6,4)*0.5
        assert att_ws['AI6'].value == '=COUNTIF(D6:AH6,8)+COUNTIF(D6:AH6,4)*0.5'
        # 2、考勤的总计填上公式
        assert att_ws.cell(row=7, column=35).value == '=SUM(AI6:AI6)'

        # 4、工资确认表中的出勤日期用VLOOKUP公式从考勤表引用
        assert wage_ws['D5'].value == '=VLOOKUP(B5,考勤表!$B$6:$AI$6,34,FALSE)'
        # 工资总额公式：出勤工日*工资标准
        assert wage_ws['F5'].value == '=D5*E5'
        # 新增津贴列与应支付公式：工资总额+津贴
        assert wage_ws['G5'].value == 0
        assert wage_ws['L5'].value == '=F5+G5'

        # 合计行 (row 6)
        assert wage_ws.cell(row=6, column=1).value == '合计'
        assert wage_ws.cell(row=6, column=6).value == '=SUM(F5:F5)'
        assert wage_ws.cell(row=6, column=7).value == '=SUM(G5:G5)'
        assert wage_ws.cell(row=6, column=12).value == '=SUM(L5:L5)'

        # 底部申明签字区 (row 7)
        assert '申明' in str(wage_ws.cell(row=7, column=1).value)

        card_col = next(cell.column for cell in wage_ws[3] if cell.value == '银行卡号')
        assert card_col == 15
        assert wage_ws.cell(row=5, column=card_col).data_type == 's'
        assert wage_ws.cell(row=5, column=card_col).value == '6222020000000000000'

        # 验证考勤与工资表人员顺序对应
        assert att_ws.cell(row=6, column=2).value == wage_ws.cell(row=5, column=2).value == '张三'
        assert att_ws.cell(row=6, column=1).value == wage_ws.cell(row=5, column=1).value == 1


def test_attendance_and_wage_order_and_zero_day_alignment():
    """验证多人员场景下考勤表与工资表人员名单及顺序严格一一对应，考勤标记为数字8，0工日人员统一剔除。"""
    salary_df = pd.DataFrame([
        {
            "分包/所属企业": "江苏旭之升建筑工程有限公司",
            "姓名": "张三",
            "身份证号": "110101199001011234",
            "月份": "8月",
            "__考勤明细__": {1: 8.0, 2: 8.0},
            "最终核定天数": 2.0,
            "日薪": 300,
            "津贴": 500,
            "应发工资": 1100,
            "工种": "木工",
            "联系电话": "13800000000",
            "开户银行": "工商银行",
            "银行卡号": "6222020000000000000",
            "性别": "男",
        },
        {
            "分包/所属企业": "江苏旭之升建筑工程有限公司",
            "姓名": "李四",
            "身份证号": "110101199001011235",
            "月份": "8月",
            "__考勤明细__": {},
            "最终核定天数": 0.0,
            "日薪": 300,
            "津贴": 0,
            "应发工资": 500,  # 仅津贴或全额延发到储备池
            "工种": "木工",
            "联系电话": "13800000001",
            "开户银行": "工商银行",
            "银行卡号": "6222020000000000001",
            "性别": "男",
        },
        {
            "分包/所属企业": "江苏旭之升建筑工程有限公司",
            "姓名": "王五",
            "身份证号": "110101199001011236",
            "月份": "8月",
            "__考勤明细__": {1: 8.0},
            "最终核定天数": 1.0,
            "日薪": 300,
            "津贴": 0,
            "应发工资": 300,
            "工种": "钢筋工",
            "联系电话": "13800000002",
            "开户银行": "工商银行",
            "银行卡号": "6222020000000000002",
            "性别": "男",
        },
    ])

    zip_data = build_all_exports_zip(salary_df, include_ledger=False)
    with ZipFile(BytesIO(zip_data)) as archive:
        wb = load_workbook(BytesIO(archive.read('江苏旭之升-8月-考勤与工资确认表.xlsx')), data_only=False)
        assert wb.sheetnames == ['考勤表', '工资确认表']

        att_ws = wb['考勤表']
        wage_ws = wb['工资确认表']

        # 李四考勤天数为0，已被统一剔除，只保留张三和王五（共2人）
        # 验证表头logo图片完整保留（2张logo图片）
        assert len(att_ws._images) == 2

        # 验证考勤表：
        # row 6: 张三
        assert att_ws.cell(row=6, column=1).value == 1
        assert att_ws.cell(row=6, column=2).value == '张三'
        assert att_ws.cell(row=6, column=4).value == 8  # 数字8
        assert att_ws.cell(row=6, column=35).value == '=COUNTIF(D6:AH6,8)+COUNTIF(D6:AH6,4)*0.5'
        # row 7: 王五
        assert att_ws.cell(row=7, column=1).value == 2
        assert att_ws.cell(row=7, column=2).value == '王五'
        assert att_ws.cell(row=7, column=4).value == 8  # 数字8
        assert att_ws.cell(row=7, column=35).value == '=COUNTIF(D7:AH7,8)+COUNTIF(D7:AH7,4)*0.5'

        # 验证数字8格式严格统一（字体、字号、通用数字格式、居中对齐）
        assert att_ws.cell(row=6, column=4).font.name == att_ws.cell(row=7, column=4).font.name == '仿宋_GB2312'
        assert att_ws.cell(row=6, column=4).font.size == att_ws.cell(row=7, column=4).font.size == 9.0
        assert att_ws.cell(row=6, column=4).number_format == att_ws.cell(row=7, column=4).number_format == 'General'
        assert att_ws.cell(row=6, column=4).alignment.horizontal == att_ws.cell(row=7, column=4).alignment.horizontal == 'center'

        # 验证考勤总计已填上公式，且末位工人后紧接着就是底部申明签字区
        assert '申明' in str(att_ws.cell(row=8, column=1).value)
        assert att_ws.cell(row=8, column=35).value == '=SUM(AI6:AI7)'

        # 验证工资确认表：
        # 1. 验证字段公式与计算
        # row 5: 张三
        assert wage_ws.cell(row=5, column=1).value == 1
        assert wage_ws.cell(row=5, column=2).value == '张三'
        assert wage_ws.cell(row=5, column=4).value == '=VLOOKUP(B5,考勤表!$B$6:$AI$7,34,FALSE)'
        assert wage_ws.cell(row=5, column=6).value == '=D5*E5'
        assert wage_ws.cell(row=5, column=7).value == 500
        assert wage_ws.cell(row=5, column=12).value == '=F5+G5'

        # row 6: 王五
        assert wage_ws.cell(row=6, column=1).value == 2
        assert wage_ws.cell(row=6, column=2).value == '王五'
        assert wage_ws.cell(row=6, column=4).value == '=VLOOKUP(B6,考勤表!$B$6:$AI$7,34,FALSE)'
        assert wage_ws.cell(row=6, column=6).value == '=D6*E6'
        assert wage_ws.cell(row=6, column=7).value == 0
        assert wage_ws.cell(row=6, column=12).value == '=F6+G6'

        # row 7: 合计行
        assert wage_ws.cell(row=7, column=1).value == '合计'
        assert wage_ws.cell(row=7, column=6).value == '=SUM(F5:F6)'
        assert wage_ws.cell(row=7, column=7).value == '=SUM(G5:G6)'
        assert wage_ws.cell(row=7, column=12).value == '=SUM(L5:L6)'

        # 工资表第8行紧接底部申明区，无多余空行
        assert '申明' in str(wage_ws.cell(row=8, column=1).value)

