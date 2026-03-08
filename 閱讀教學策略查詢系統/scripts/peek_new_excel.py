import pandas as pd
import json

file_path = r"d:\test ch\閱讀教學策略查詢系統\學習扶助補充文本(教育部)\學扶補充文章1-6年級.xlsx"

try:
    df = pd.read_excel(file_path)
    
    # Convert first few rows to dict
    sample_data = df.head(5).to_dict(orient='records')
    
    with open('peek_excel.json', 'w', encoding='utf-8') as f:
        json.dump({
            "columns": df.columns.tolist(),
            "sample": sample_data,
            "total_rows": len(df)
        }, f, ensure_ascii=False, indent=2)
        
    print("Data written to peek_excel.json")
except Exception as e:
    print(f"Error reading file: {e}")
