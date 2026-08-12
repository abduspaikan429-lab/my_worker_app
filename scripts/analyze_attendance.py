import pandas as pd
import warnings
warnings.filterwarnings('ignore')
import os

load_dir = 'load-data'
files = [os.path.join(load_dir, f) for f in os.listdir(load_dir) if f.endswith('.xlsx')]

for filepath in files:
    fname = os.path.basename(filepath)
    print(f'\n{"="*60}')
    print(f'FILE: {fname}')
    try:
        xl = pd.ExcelFile(filepath)
        print(f'Sheets: {xl.sheet_names}')
        for sheet in xl.sheet_names:
            print(f'\n  -- Sheet: [{sheet}] --')
            # Read raw without header to see structure
            raw = pd.read_excel(filepath, sheet_name=sheet, dtype=str, header=None, nrows=12)
            print(f'  Shape (raw 12 rows): {raw.shape}')
            for i in range(min(8, len(raw))):
                row_vals = [str(v)[:20] if str(v) != 'nan' else 'NaN' for v in raw.iloc[i].tolist()]
                print(f'  row[{i}]: {row_vals[:10]}')

            # Auto-detect header row: find row containing common name patterns
            header_row = None
            for i in range(min(8, len(raw))):
                row_str = ' '.join(str(v) for v in raw.iloc[i].tolist())
                if any(kw in row_str for kw in ['姓名', '名字', '人员']):
                    header_row = i
                    print(f'  >> DETECTED header_row={i}')
                    break

            if header_row is not None:
                df = pd.read_excel(filepath, sheet_name=sheet, dtype=str, header=header_row)
                # Clean columns
                df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
                print(f'  Columns: {list(df.columns)}')
                # Find name col and days col
                name_candidates = ['姓名', '名字', '人员姓名']
                days_candidates = ['小计', '合计', '出勤天数', '天数', '出勤', '实出勤']
                name_col = next((c for c in name_candidates if c in df.columns), None)
                days_col = next((c for c in days_candidates if c in df.columns), None)
                print(f'  name_col={name_col}, days_col={days_col}')
                if name_col:
                    sample_names = df[name_col].dropna().head(5).tolist()
                    print(f'  Sample names: {sample_names}')
                if days_col:
                    sample_days = df[days_col].dropna().head(5).tolist()
                    print(f'  Sample days: {sample_days}')
    except Exception as e:
        print(f'ERROR: {e}')
