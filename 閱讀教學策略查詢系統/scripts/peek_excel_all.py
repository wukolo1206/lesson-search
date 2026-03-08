import pandas as pd
import json

file_path = r"d:\test ch\閱讀教學策略查詢系統\學習扶助補充文本(教育部)\學扶補充文章1-6年級.xlsx"

try:
    # Read all sheets instead of just the first one
    all_dfs = pd.read_excel(file_path, sheet_name=None)
    
    sheet_info = {}
    for sheet_name, df in all_dfs.items():
        sheet_info[sheet_name] = {
            "columns": df.columns.tolist(),
            "row_count": len(df)
        }
    
    with open('peek_excel_all_sheets.json', 'w', encoding='utf-8') as f:
        json.dump(sheet_info, f, ensure_ascii=False, indent=2)
        
    print("All sheets data written to peek_excel_all_sheets.json")
except Exception as e:
    print(f"Error reading file: {e}")
