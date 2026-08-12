import pandas as pd
import json

file_path = r'D:\code\工作脚本\my_worker_app\load-data\二局安装-足球场项目人员变更月报表（进场情况）.xlsx'
xl = pd.ExcelFile(file_path)

output = []
for sheet in xl.sheet_names:
    df = xl.parse(sheet, nrows=5)
    
    # Try to grab some text from the first few rows to understand what this sheet is about
    # Often, row 1 or 2 contains the title, e.g., "某某公司6月人员进场表"
    # We will convert the first few rows to strings and put them in a list
    sample_data = []
    for row in df.head(5).values.tolist():
        row_clean = [str(x) for x in row if pd.notna(x)]
        if row_clean:
            sample_data.append(row_clean)
            
    output.append({
        "sheet_name": sheet,
        "sample_rows": sample_data
    })

with open('excel_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)
print("Analysis complete.")
