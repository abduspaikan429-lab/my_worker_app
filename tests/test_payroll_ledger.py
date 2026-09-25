# -*- coding: utf-8 -*-
"""
tests/test_payroll_ledger.py
测试工资发放动态总台账服务（支持多分包公司）：
1. 单批次文本解析（自动提取分包商名称）
2. 多批次扫描汇总（单分包隔离与全部分包汇总）
3. 工人最新状态及多批次发薪追踪
4. 导出到目标文件夹（单分包专属与全盘汇总）
"""
import os
import shutil
import pytest
from services.payroll_ledger_service import (
    parse_single_batch_text,
    load_all_payroll_data,
    export_ledger_to_folder,
    get_excel_bytes,
    get_batches_dir,
    get_all_companies
)

SAMPLE_BATCH_COMP_A = """科技文化中心-国际体育中心项目（专业足球场）2026年09月第1批次工资表,,,,,,,,,,,,,,,,,,,
工资单编号：GZ-20260930111111,,,,,,,,,,,,,,,,,,,
分包商名称：江苏旭之升建筑工程有限公司,,,,,,,,,,,,,,,,,,,
班组：二局安装江苏旭之升工人班组,,,,,,,,,,,,,,,,,,,
所属月份：2026.09.01-2026.09.30,,,,,,,,,,,,,,,,,,,
序号,姓名,身份证号,电话号码,实名制银行卡号,开户银行名称,开户银行支行名称,班组,进退场状态,工资标准,工种,实际出勤天数,实际工作量,合同工资,合同工资,支出,实发工资,实发工资,津贴,合计
序号,姓名,身份证号,电话号码,实名制银行卡号,开户银行名称,开户银行支行名称,班组,进退场状态,工资标准,工种,实际出勤天数,实际工作量,基本工资,入账绩效工资,支出,基本工资,实际入账绩效工资,津贴,合计
1,测试工人A,110101199001011234,13800000001,6222000000000001,中国工商银行,北京支行,二局安装江苏旭之升工人班组,已进场,基本工资：500元/天,电焊工,10天,0,5000,0,0,5000,0,0,5000
合计：,,,,,,,,,,,,,5000,0,0,5000,0,0,5000
"""

SAMPLE_BATCH_COMP_B = """科技文化中心-国际体育中心项目（专业足球场）2026年09月第1批次工资表,,,,,,,,,,,,,,,,,,,
工资单编号：GZ-20260930222222,,,,,,,,,,,,,,,,,,,
分包商名称：青海久昌建筑装饰工程有限公司,,,,,,,,,,,,,,,,,,,
班组：二局装饰青海久昌工人班组,,,,,,,,,,,,,,,,,,,
所属月份：2026.09.01-2026.09.30,,,,,,,,,,,,,,,,,,,
序号,姓名,身份证号,电话号码,实名制银行卡号,开户银行名称,开户银行支行名称,班组,进退场状态,工资标准,工种,实际出勤天数,实际工作量,合同工资,合同工资,支出,实发工资,实发工资,津贴,合计
序号,姓名,身份证号,电话号码,实名制银行卡号,开户银行名称,开户银行支行名称,班组,进退场状态,工资标准,工种,实际出勤天数,实际工作量,基本工资,入账绩效工资,支出,基本工资,实际入账绩效工资,津贴,合计
1,测试久昌工友,110101199001019999,13800000099,6222000000000099,中国农业银行,西宁支行,二局装饰青海久昌工人班组,已进场,基本工资：400元/天,油漆工,8天,0,3200,0,0,3200,0,0,3200
合计：,,,,,,,,,,,,,3200,0,0,3200,0,0,3200
"""

def test_parse_single_batch_subcontractor():
    b_a = parse_single_batch_text(SAMPLE_BATCH_COMP_A)
    assert b_a["subcontractor"] == "江苏旭之升建筑工程有限公司"
    assert len(b_a["records"]) == 1
    assert b_a["calc_total"] == 5000.0

    b_b = parse_single_batch_text(SAMPLE_BATCH_COMP_B)
    assert b_b["subcontractor"] == "青海久昌建筑装饰工程有限公司"
    assert len(b_b["records"]) == 1
    assert b_b["calc_total"] == 3200.0

def test_get_all_companies():
    comps = get_all_companies()
    assert "江苏旭之升建筑工程有限公司" in comps
    assert "青海久昌建筑装饰工程有限公司" in comps

def test_load_all_payroll_data_company_filter():
    # 测试单分包商过滤
    data_xzs = load_all_payroll_data(company_filter="江苏旭之升建筑工程有限公司")
    assert data_xzs["company_scope"] == "江苏旭之升建筑工程有限公司"
    assert len(data_xzs["sorted_workers"]) >= 35
    assert data_xzs["stats"]["total_pay"] >= 436450.0

    # 测试多分包全盘汇总
    data_all = load_all_payroll_data(company_filter="全部劳务公司")
    assert data_all["company_scope"] == "全部劳务公司"
    assert len(data_all["sorted_workers"]) >= 35

def test_export_ledger_to_folder_multi_company(tmp_path):
    temp_dir = str(tmp_path)
    
    # 导出指定单公司
    ok1, path1, err1 = export_ledger_to_folder(
        temp_dir,
        filename="江苏旭之升_工资台账.xlsx",
        company_scope="江苏旭之升建筑工程有限公司"
    )
    assert ok1 is True
    assert os.path.exists(path1)

    # 导出全部分包汇总
    ok2, path2, err2 = export_ledger_to_folder(
        temp_dir,
        filename="全盘汇总_工资台账.xlsx",
        company_scope="全部劳务公司"
    )
    assert ok2 is True
    assert os.path.exists(path2)

def test_parse_and_upload_excel_batch():
    import openpyxl
    import io
    from services.payroll_ledger_service import save_uploaded_batch, delete_batch_file
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['科技文化中心-国际体育中心项目（专业足球场）2026年10月第1批次工资表'])
    ws.append(['工资单编号：GZ-20261015888888'])
    ws.append(['分包商名称：青海久昌建筑装饰工程有限公司'])
    ws.append(['班组：二局装饰青海久昌工人班组'])
    ws.append(['所属月份：2026.10.01-2026.10.31'])
    ws.append(['序号', '姓名', '身份证号', '电话号码', '实名制银行卡号', '开户银行名称', '开户银行支行名称', '班组', '进退场状态', '工资标准', '工种', '实际出勤天数', '实际工作量', '基本工资', '入账绩效工资', '支出', '基本工资', '实际入账绩效工资', '津贴', '合计'])
    ws.append([1, 'Excel上传测试工人', '630101199501017777', '13900008888', '6228489999999999', '中国农业银行', '西宁城中支行', '二局装饰青海久昌工人班组', '已进场', '基本工资：450元/天', '电焊工', '15天', '0', 6750, 0, 0, 6750, 0, 0, 6750])
    
    bio = io.BytesIO()
    wb.save(bio)
    excel_bytes = bio.getvalue()
    
    cnt, msgs = save_uploaded_batch(excel_bytes, assigned_company="青海久昌建筑装饰工程有限公司", filename="青海久昌_10月第1批次.xlsx")
    assert cnt == 1
    assert "成功归档至【青海久昌建筑装饰工程有限公司】" in msgs[0]

    # 清理产生的测试批次
    delete_batch_file("青海久昌建筑装饰工程有限公司_2026年10月第1批_888888.csv")

