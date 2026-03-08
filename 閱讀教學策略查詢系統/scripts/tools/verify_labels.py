import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
checks = [
    ('三年級', '獅子蚊子',       '21', '推論訊息'),
    ('三年級', '水鴨',           '23', '推論訊息'),
    ('四年級', '動物放煙火',     '22', '比較評估'),
    ('五年級', '林昀儒的故事',   '19', '詮釋整合'),
    ('六年級', '水都威尼斯',     '20', '提取訊息'),
    ('六年級', '白冷圳採果趣',   '25', '提取訊息'),
]

all_pass = True
for sheet, title, qnum, expected in checks:
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
    row = df[(df['文本標題'].astype(str) == title) & (df['題號'].astype(str) == qnum)]
    if len(row) > 0:
        actual = str(row.iloc[0]['認知歷程'])
        ok = '✅' if actual == expected else '❌'
        if actual != expected:
            all_pass = False
        print(f"{ok} {sheet} | {title} Q{qnum} => 【{actual}】（預期：{expected}）")
    else:
        print(f"❓ 找不到：{sheet} | {title} Q{qnum}")
        all_pass = False

print()
print("✅ 所有修正均已確認！" if all_pass else "❌ 有標籤尚未正確更新，請檢查。")
