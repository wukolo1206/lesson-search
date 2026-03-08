import pandas as pd
file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
df = pd.read_excel(file_path, sheet_name='學習遷移題目')

missing_count = 0
for idx, row in df.iterrows():
    q_str = str(row['遷移題目'])
    # Check if there are NO options like A, B, 1, 2, or brackets
    if q_str != 'nan' and not any(x in q_str for x in ['A', 'B', '1', '2', '(', '選項']):
        print(f"Row {idx+2}: {row['原始案例標題']}")
        print(f"Raw Text: {q_str}")
        print("-" * 20)
        missing_count += 1
        if missing_count >= 5:
            break
