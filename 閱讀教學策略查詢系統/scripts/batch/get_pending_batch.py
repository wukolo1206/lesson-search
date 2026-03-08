import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
df = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')

# Print rows 3-6 (index 1-4)
pending = df[df['審查狀態'] == '待審查']
batch = pending.iloc[1:5]  # Skip row 2 (already reviewed), get next 4

for i, (idx, row) in enumerate(batch.iterrows()):
    print(f"\n{'='*60}")
    print(f"[題目 {i+2}/5] Row {idx+2} | 原始文章: {row.get('原始案例標題', '')}")
    print(f"認知歷程: {row.get('認知歷程', '')} | 教學策略: {row.get('教學策略', '')}")
    print(f"--- 原始題目 ---")
    print(str(row.get('原始題目', ''))[:200])
    print(f"--- 遷移文本(節錄400字) ---")
    print(str(row.get('遷移文本內容', ''))[:400])
    print(f"--- 遷移題目(完整) ---")
    print(str(row.get('遷移題目', '')))
