import pandas as pd
import os

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
OUTPUT_DIR = r'd:\test ch\閱讀教學策略查詢系統\review_batch'
os.makedirs(OUTPUT_DIR, exist_ok=True)

df = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')
pending = df[df['審查狀態'] == '待審查'].head(5)

for i, (idx, row) in enumerate(pending.iterrows()):
    filename = os.path.join(OUTPUT_DIR, f'q{i+1}_row{idx+2}.txt')
    content = f"""行號(Excel Row): {idx+2}
原始案例標題: {row.get('原始案例標題', '')}
認知歷程: {row.get('認知歷程', '')}
教學策略: {row.get('教學策略', '')}

==== 原始題目 ====
{row.get('原始題目', '')}

==== 遷移文本（前400字） ====
{str(row.get('遷移文本內容', ''))[:400]}

==== 遷移題目（完整） ====
{row.get('遷移題目', '')}
"""
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"已輸出: {filename}")

print("完成！共輸出 5 個檔案。")
