"""
修正後儲存腳本 - 逐個分頁讀取並寫回
"""
import pandas as pd
import os

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
BACKUP_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析_備份_原題標籤修正前.xlsx'
GRADE_SHEETS = ['三年級', '四年級', '五年級', '六年級']

CORRECTIONS = [
    ("三年級", "獅子蚊子",       "21", "提取訊息", "推論訊息"),
    ("三年級", "水鴨",           "23", "提取訊息", "推論訊息"),
    ("四年級", "動物放煙火",     "22", "詮釋整合", "比較評估"),
    ("五年級", "林昀儒的故事",   "19", "推論訊息", "詮釋整合"),
    ("六年級", "水都威尼斯",     "20", "推論訊息", "提取訊息"),
    ("六年級", "白冷圳採果趣",   "25", "比較評估", "提取訊息"),
    # 臺灣俠醫林杰樑 Q21 已確認是"提取訊息"，無需更改
]

# 先備份原始檔案
import shutil
if not os.path.exists(BACKUP_PATH):
    shutil.copy2(EXCEL_PATH, BACKUP_PATH)
    print(f"✅ 備份已儲存至：{BACKUP_PATH}")

print("載入所有分頁...")
all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)

# 套用修正
update_count = 0
for (sheet_name, title, q_num, old_label, new_label) in CORRECTIONS:
    if sheet_name not in all_sheets:
        print(f"⚠️  找不到分頁：{sheet_name}")
        continue
    df = all_sheets[sheet_name]
    mask = (
        (df['文本標題'].astype(str) == title) &
        (df['題號'].astype(str) == q_num)
    )
    matched = df[mask]
    if len(matched) == 0:
        print(f"❌  找不到：{sheet_name} | {title} Q{q_num}")
        continue
    
    current = matched.iloc[0]['認知歷程']
    all_sheets[sheet_name].loc[matched.index, '認知歷程'] = new_label
    print(f"✅  {sheet_name} | {title} Q{q_num} ：【{current}】→【{new_label}】")
    update_count += 1

# 寫回 Excel（全部分頁）
print(f"\n💾 寫回 Excel 中（{update_count} 筆修正）...")
with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    for name, df in all_sheets.items():
        df.to_excel(writer, sheet_name=name, index=False)

print("✅ 全部完成！Excel 已更新。")
print(f"（原始備份：{BACKUP_PATH}）")
