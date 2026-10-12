import pandas as pd
import pytest
from modules.archive_export import (
    ARCHIVE_COLUMNS,
    extract_cscec2_archive,
    ensure_excel_text as cscec_ensure_text,
)
from modules.report_generator import (
    FULL_INFO_COLUMNS,
    extract_full_info,
    extract_roster,
    extract_monthly_report,
    paste_text,
    ensure_excel_text,
)


@pytest.fixture
def sample_worker_df():
    return pd.DataFrame([
        {
            "姓名": "张三",
            "性别": "男",
            "民族": "汉",
            "年龄": "35",
            "身份证号": "110101199001011234",
            "手机号": "13812345678",
            "详细地址": "四川省成都市金牛区某街道",
            "家庭住址": "四川省成都市金牛区某街道",
            "紧急联系人": "李四",
            "紧急联系电话": "13987654321",
            "班组": "钢筋班组",
            "工种": "电工",
            "人员类型": "班组长",
            "进场日期": "2026-05-01",
            "进场时间": "2026-05-01",
            "在场状态": "在场",
            "进退场状态": "在场",
            "在场/进退场状态": "在场",
            "银行卡号": "6222021234567890123",
            "工资卡号": "6222021234567890123",
            "开户银行": "中国工商银行",
            "劳动合同编号": "HT-2026-001",
            "合同签订状态": "已签署完成",
            "劳动合同": "已签署",
            "是否在市建委": "是",
            "分包/所属企业": "中建二局劳务分包公司",
        },
        {
            "姓名": "王五",
            "性别": "男",
            "民族": "土家族",
            "年龄": "",
            "身份证号": "420101199508085678",
            "手机号": "13700001111",
            "详细地址": "湖北省恩施州某村",
            "班组": "木工班组",
            "工种": "普通木工",
            "人员类型": "工人",
            "进场日期": "2026-06-15",
            "在场状态": "在场",
            "银行卡号": "6228480000000001234",
            "开户银行": "中国农业银行",
            "劳动合同编号": "",
            "合同签订状态": "否",
            "是否在市建委": "否",
            "分包/所属企业": "中建二局劳务分包公司",
        }
    ])


def test_ensure_excel_text():
    assert ensure_excel_text("110101199001011234") == "'110101199001011234"
    assert ensure_excel_text("'110101199001011234") == "'110101199001011234"
    assert ensure_excel_text("") == ""
    assert ensure_excel_text(None) == ""
    assert ensure_excel_text("nan") == ""
    assert ensure_excel_text("123.0") == "'123"


def test_extract_full_info(sample_worker_df):
    full_df = extract_full_info(sample_worker_df)
    assert not full_df.empty
    assert len(full_df) == 2
    
    # 严格验证38列标准表头顺序与列数
    assert list(full_df.columns) == FULL_INFO_COLUMNS
    assert len(full_df.columns) == 38

    # 验证用户明确要求的「序号留空」
    assert full_df.iloc[0]["序号"] == ""
    assert full_df.iloc[1]["序号"] == ""
    
    # 验证关键列名与内容存在
    assert full_df.iloc[0]["姓名"] == "张三"
    assert full_df.iloc[0]["性别"] == "男"
    assert full_df.iloc[0]["民族"] == "汉"
    assert full_df.iloc[0]["工资卡号"].startswith("'6222021234567890123")
    assert full_df.iloc[0]["银行卡号"].startswith("'6222021234567890123")

    # 验证身份证和银行卡号被添加了单引号保护
    assert full_df.iloc[0]["身份证号"].startswith("'")
    assert full_df.iloc[0]["银行卡号"].startswith("'")
    assert full_df.iloc[0]["手机号"].startswith("'")
    assert "110101199001011234" in full_df.iloc[0]["身份证号"]


def test_extract_cscec2_archive(sample_worker_df):
    archive_df = extract_cscec2_archive(sample_worker_df)
    assert not archive_df.empty
    assert len(archive_df) == 2
    assert list(archive_df.columns) == ARCHIVE_COLUMNS
    
    # 验证第一行数据
    row1 = archive_df.iloc[0]
    assert row1["姓名"] == "张三"
    assert row1["性别"] == "男"
    assert row1["身份证号"].startswith("'110101199001011234")
    assert row1["银行卡号"].startswith("'6222021234567890123")
    assert row1["手机"].startswith("'13812345678")
    assert row1["是否为特殊工种"] == "是"  # 电工属于特殊工种
    assert "联系人：李四" in row1["家庭住址、家庭联系人及电话"]
    assert "电话：13987654321" in row1["家庭住址、家庭联系人及电话"]

    # 验证第二行数据
    row2 = archive_df.iloc[1]
    assert row2["姓名"] == "王五"
    assert row2["是否为特殊工种"] == "否"  # 普通木工不属于特殊工种
    assert row2["身份证号"].startswith("'420101199508085678")
    # 年龄未填时应从身份证自动推算
    assert row2["年龄"] != ""


def test_paste_text_format(sample_worker_df):
    full_df = extract_full_info(sample_worker_df)
    tsv_text = paste_text(full_df)
    
    lines = [l for l in tsv_text.splitlines() if l]
    assert len(lines) == 2  # 仅2行数据，无表头
    
    # 每行应使用 tab 分隔
    cols_row1 = lines[0].split("\t")
    assert len(cols_row1) == len(full_df.columns)
    assert len(cols_row1) == 38
    # 序号留空：第一列为空字符串，第二列为姓名"张三"
    assert cols_row1[0] == ""
    assert cols_row1[1] == "张三"
    assert "'110101199001011234" in cols_row1


def test_empty_dataframe_handling():
    empty_df = pd.DataFrame()
    
    full_res = extract_full_info(empty_df)
    assert full_res.empty
    assert list(full_res.columns) == FULL_INFO_COLUMNS
    
    archive_res = extract_cscec2_archive(empty_df)
    assert archive_res.empty
    assert list(archive_res.columns) == ARCHIVE_COLUMNS

    assert paste_text(empty_df) == ""
    assert paste_text(None) == ""


def test_extract_roster_and_monthly(sample_worker_df):
    roster_df = extract_roster(sample_worker_df)
    assert len(roster_df) == 2
    assert "工种(或岗位)" in roster_df.columns
    assert roster_df.iloc[0]["身份证号"].startswith("'")

    monthly_df = extract_monthly_report(sample_worker_df)
    assert len(monthly_df) == 2
    assert monthly_df.iloc[0]["是否签订《劳动合同》"] == "是"
    assert monthly_df.iloc[1]["是否签订《劳动合同》"] == "否"


def test_extract_full_info_partial_official_import():
    # 模拟官网导出仅含少量字段的新增人员
    raw_incoming = pd.DataFrame([
        {
            "姓名": "李四",
            "身份证号": "510101200001011234",
            "电话": "13800002222",
            "工种名称": "钢筋工",
            "所属班组": "钢筋一组",
            "入场日期": "2026-09-01",
            "银行账号": "6222000011112222333",
        }
    ])
    full_df = extract_full_info(raw_incoming)
    assert list(full_df.columns) == FULL_INFO_COLUMNS
    assert len(full_df.columns) == 38

    row = full_df.iloc[0]
    # 序号留空
    assert row["序号"] == ""
    # 基础信息
    assert row["姓名"] == "李四"
    assert row["身份证号"] == "'510101200001011234"
    assert row["手机号"] == "'13800002222"
    assert row["工种"] == "钢筋工"
    assert row["班组"] == "钢筋一组"
    assert row["进场日期"] == "2026-09-01"
    # 同义词回退
    assert row["工资卡号"] == "'6222000011112222333"
    assert row["银行卡号"] == "'6222000011112222333"
    # 身份证推算性别与出生日期
    assert row["性别"] == "男"
    assert row["出生日期"] == "2000-01-01"
    assert row["年龄"] != ""
    # 官网未提供但本地Excel要求的字段：补空字符串
    assert row["进场告知书"] == ""
    assert row["是否参与统计"] == ""
    assert row["协议签订"] == ""

    # TSV粘贴文本测试
    tsv = paste_text(full_df)
    lines = [l for l in tsv.splitlines() if l]
    cols = lines[0].split("\t")
    assert len(cols) == 38
    assert cols[0] == ""  # 序号留空
    assert cols[1] == "李四"
