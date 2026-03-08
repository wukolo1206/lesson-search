"""
批次儲存審查結果：Row 7-11
"""
import json

results = [
    {
        "row": 7,
        "status": "通過",
        "notes": "認知歷程（詮釋整合）與遷移題型一致，選項誘答性佳，主旨正確為(A)。"
    },
    {
        "row": 8,
        "status": "通過",
        "notes": "認知歷程（提取訊息）一致，問文章描寫的時間，月夜明確可提取，誘答分佈良好。"
    },
    {
        "row": 9,
        "status": "通過",
        "notes": "認知歷程（推論訊息）一致，微風隱喻需有輕度推論，選項誘答合理。"
    },
    {
        "row": 10,
        "status": "修補",
        "repaired": """這篇文章主要想告訴我們什麼？
(A) 月夜讓人回想起美好的家庭時光
(B) 月亮形狀說明了什麼是翻船
(C) 晚上一定要早點回到屋裡
(D) 欣賞月夜是一種奇怪的習慣
答案：(A)""",
        "notes": "題目本身可通過，但教學策略「人物放大鏡」不適合無明顯人物的抒情文《月夜》。已維持題目，備註教學策略配對問題。"
    },
    {
        "row": 11,
        "status": "修補",
        "repaired": """關於這篇文章的篇名「默默行善的阿嬤」，下面哪個說法最適當？
(A) 篇名點出了阿嬤低調但持續做善事的特質
(B) 篇名說明阿嬤很安靜、不喜歡說話
(C) 篇名是描述阿嬤在市場默默賣菜的樣子
(D) 篇名表示阿嬤的善行還沒有被人發現
答案：(A)""",
        "notes": "原題型為「找圖文未涵蓋的項目」，遷移題卻改為普通事實確認題，且與教學策略「篇名有意思」完全無關。已修補為讓學生理解篇名深意的題型，符合「篇名有意思」策略目標。"
    }
]

# 用 save_review_results.py 相容格式輸出
import sys
sys.path.insert(0, r'd:\test ch\閱讀教學策略查詢系統')

import pandas as pd
EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

print("載入遷移題目分頁...")
df = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')

for r in results:
    idx = r["row"] - 2  # Excel Row 轉 DataFrame index (header=0, row2=index0)
    df.at[idx, '審查狀態'] = r["status"]
    df.at[idx, '審查備註'] = r["notes"]
    if r["status"] == "修補" and "repaired" in r:
        df.at[idx, '修補後題目'] = r["repaired"]
    print(f"✅ Row {r['row']} → {r['status']}")

# 寫回（保留所有分頁）
all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
all_sheets['學習遷移題目'] = df

with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    for name, sheet_df in all_sheets.items():
        sheet_df.to_excel(writer, sheet_name=name, index=False)

print("\n✅ 審查結果已寫回 Excel！")
print("已完成：Row 7~11（本批 5 題，累計 10/585）")
