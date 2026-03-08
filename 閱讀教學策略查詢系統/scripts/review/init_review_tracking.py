"""
初始化審查追蹤欄位腳本
執行一次即可，為 Excel 的「學習遷移題目」分頁新增追蹤欄位。
"""
import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

print("載入 Excel...")
df = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')

# 只有在欄位不存在的情況下才新增
if '審查狀態' not in df.columns:
    df.insert(len(df.columns), '審查狀態', '待審查')
    print("新增「審查狀態」欄位。")
else:
    print("「審查狀態」欄位已存在，跳過。")

if '審查備註' not in df.columns:
    df.insert(len(df.columns), '審查備註', '')
    print("新增「審查備註」欄位。")
else:
    print("「審查備註」欄位已存在，跳過。")

if '修補後題目' not in df.columns:
    df.insert(len(df.columns), '修補後題目', '')
    print("新增「修補後題目」欄位。")
else:
    print("「修補後題目」欄位已存在，跳過。")

# 用 openpyxl 寫回，保留其他分頁
print("寫回 Excel（保留所有其他分頁）...")
with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='學習遷移題目', index=False)

print("✅ 完成！現在 Excel 已有追蹤欄位，可以開始分批審查。")
print(f"總計 {len(df)} 題待審查。")
