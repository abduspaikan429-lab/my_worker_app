"""
验证脚本：用实际考勤文件测试修复后的 _safe_read 和 build_diff_table
"""
import sys
import warnings
import io
warnings.filterwarnings('ignore')

sys.path.insert(0, '.')
from modules.attendance_payroll import _safe_read, _detect_header_row, build_diff_table, _NAME_CANDS, _DAYS_CANDS, _resolve_col
import pandas as pd

files = {
    '系统A（旭之升）':  'load-data/2026年07月考勤表_科技文化中心-国际体育中心（专业足球场）工程.xlsx',
    '系统B（旭之升）':  'load-data/考勤表_科技文化中心-国际体育中心项目（专业足球场）_江苏旭之升建筑工程有限公司_20260731.xlsx',
    '纸质（旭之升）':   'load-data/7月水印照片考勤表.xlsx',
    '系统A（久昌）':   'load-data/2026年07月考勤表_科技文化中心-国际体育中心（专业足球场）工程 (1).xlsx',
    '系统B（久昌）':   'load-data/考勤表_科技文化中心-国际体育中心项目（专业足球场）_青海久昌建筑装饰工程有限公司_20260731.xlsx',
    '纸质（久昌）':    'load-data/7月水印照片-考勤表（一式两份本人签字摁手印）.xlsx',
}

print('=' * 60)
print('逐文件读取测试')
print('=' * 60)

dfs = {}
for label, path in files.items():
    with open(path, 'rb') as f:
        raw_bytes = f.read()
    file_obj = io.BytesIO(raw_bytes)
    header_row = _detect_header_row(file_obj)
    file_obj.seek(0)
    df = _safe_read(file_obj, label)
    name_col = next((c for c in _NAME_CANDS if c in df.columns), None)
    days_col = next((c for c in _DAYS_CANDS if c in df.columns), None)
    print(f'\n[{label}]')
    print(f'  检测表头行: {header_row}')
    print(f'  有效行数:   {len(df)}')
    print(f'  姓名列:     {name_col}')
    print(f'  天数列:     {days_col}')
    if name_col:
        names = df[name_col].head(5).tolist()
        print(f'  前5人:      {names}')
    if days_col:
        days = df[days_col].head(5).tolist()
        print(f'  前5天数:    {days}')
    dfs[label] = df

print('\n' + '=' * 60)
print('build_diff_table 测试（旭之升三方）')
print('=' * 60)

try:
    with open('load-data/2026年07月考勤表_科技文化中心-国际体育中心（专业足球场）工程.xlsx', 'rb') as f:
        fa = io.BytesIO(f.read())
    with open('load-data/考勤表_科技文化中心-国际体育中心项目（专业足球场）_江苏旭之升建筑工程有限公司_20260731.xlsx', 'rb') as f:
        fb = io.BytesIO(f.read())
    with open('load-data/7月水印照片考勤表.xlsx', 'rb') as f:
        fc = io.BytesIO(f.read())

    df_a    = _safe_read(fa, '系统A')
    fb.seek(0)
    df_b    = _safe_read(fb, '系统B')
    fc.seek(0)
    df_p    = _safe_read(fc, '纸质')

    result = build_diff_table(df_a, df_b, df_p)
    print(f'  对比结果行数: {len(result)}')
    print(f'  三方一致:     {(result["差异状态"] == "✔ 一致").sum()}')
    print(f'  存在差异:     {result["差异状态"].str.startswith("⚠").sum()}')
    print(f'  数据缺失:     {(result["差异状态"] == "— 数据缺失").sum()}')
    print()
    print(result.head(10).to_string(index=False))
except Exception as e:
    import traceback
    print(f'ERROR: {e}')
    traceback.print_exc()
