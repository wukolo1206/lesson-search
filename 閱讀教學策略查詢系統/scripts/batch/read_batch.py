import pandas as pd
import json

file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
df_trans = pd.read_excel(file_path, sheet_name='學習遷移題目')

# Extract first 5 rows - all fields, no truncation except text
results = []
for idx, row in df_trans.iterrows():
    if idx >= 5:
        break
    results.append({
        'row': idx + 2,
        '原始案例標題': str(row.get('原始案例標題', '')),
        '原始題目': str(row.get('原始題目', '')),
        '認知歷程': str(row.get('認知歷程', '')),
        '教學策略': str(row.get('教學策略', '')),
        '遷移文本內容': str(row.get('遷移文本內容', ''))[:400],
        '遷移題目': str(row.get('遷移題目', ''))
    })

print(json.dumps(results, ensure_ascii=False, indent=2))
