"""
批次儲存審查結果：Row 12-16
"""
import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

results = [
    {
        "row": 12,
        "status": "修補",
        "repaired": """關於這篇文章的篇名「拜訪綠色博物館」，下面哪個說法最適當？
(A) 篇名把植物園比喻成博物館，強調它保存了豐富的自然生態
(B) 篇名說明作者去拜訪了一間真正的博物館
(C) 篇名是說植物園的顏色是綠色的
(D) 篇名告訴我們博物館裡有很多植物標本
答案：(A)""",
        "notes": "教學策略「篇名有意思」要求學生理解篇名的比喻義或深意，但遷移題問的是一般事實確認，與策略不符。已修補為理解篇名比喻義的題型。"
    },
    {
        "row": 13,
        "status": "修補",
        "repaired": """關於這篇文章的篇名「快樂的一天」，下面哪個說法最適當？
(A) 篇名概括了全家出遊享受自然美景的一天
(B) 篇名說明那天天氣很好所以很快樂
(C) 篇名表示只有快樂的人才會去武陵農場
(D) 篇名是說這一天非常短暫轉眼就結束了
答案：(A)""",
        "notes": "教學策略「篇名有意思」，遷移題問的是普通旅行事實，未體現篇名理解。已修補為讓學生分析篇名與全文主旨的關係。"
    },
    {
        "row": 14,
        "status": "修補",
        "repaired": """如果關山鎮的氣候不再溫和，改成又乾又冷，最可能造成什麼影響？
(A) 稻米品質下降，不再香氣十足
(B) 油菜花田會長滿整個公園
(C) 遊客改去騎腳踏車就好了
(D) 親水公園的花草會變得更美
答案：(A)""",
        "notes": "教學策略「如果吹泡泡」要求學生推論假設情境（「如果...會怎樣？」），但遷移題是普通事實確認題。已修補為假設情境推論題型。"
    },
    {
        "row": 15,
        "status": "通過",
        "notes": "認知歷程（詮釋整合）一致，選最佳標題題型正確，「自然段說什麼」策略與詮釋整合配合合理。答案(A)「美麗的關山鎮」能概括全文。"
    },
    {
        "row": 16,
        "status": "修補",
        "repaired": """如果你和爸爸猜謎時猜錯了一道題，最可能之後會做什麼？
(A) 再想一想，試著從謎語的字義去找答案
(B) 馬上去吃元宵，不再猜任何謎語了
(C) 生氣怪爸爸沒有好好幫忙猜謎
(D) 認為猜謎毫無意義，決定放棄活動
答案：(A)""",
        "notes": "教學策略「如果吹泡泡」要求假設情境推論，但遷移題是普通事實題。已修補為「如果猜錯了...會怎樣？」的假設推論題型，符合策略目標並考察推論訊息。"
    }
]

print("載入 Excel...")
all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
df = all_sheets['學習遷移題目']

for r in results:
    idx = r["row"] - 2
    df.at[idx, '審查狀態'] = r["status"]
    df.at[idx, '審查備註'] = r["notes"]
    if r["status"] == "修補" and "repaired" in r:
        df.at[idx, '修補後題目'] = r["repaired"]
    print(f"✅ Row {r['row']} → {r['status']}")

all_sheets['學習遷移題目'] = df

with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    for name, sheet_df in all_sheets.items():
        sheet_df.to_excel(writer, sheet_name=name, index=False)

print("\n✅ Row 12-16 審查結果已寫回 Excel！")
print("累計：15/585 題完成")
