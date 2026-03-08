"""
批次儲存審查結果：Row 17-21
"""
import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

results = [
    {
        "row": 17,
        "status": "通過",
        "notes": "詮釋整合一致，選最佳標題題型正確，(A)「歡樂猜燈謎」能概括全文主旨（玩謎樂趣+過節歡樂+謎語智慧）。"
    },
    {
        "row": 18,
        "status": "修補",
        "repaired": """如果你家也有倒貼「福」字的習俗，根據文章，這樣做最可能有什麼含意？
(A) 表示「福到了」，是討吉利的習俗
(B) 表示字貼反面才能看清楚
(C) 表示家裡的福氣要分給別人
(D) 表示要把不好的事情全倒掉
答案：(A)""",
        "notes": "教學策略「如果吹泡泡」要求假設情境推論，但遷移題是普通提取事實題（答案直接在文中）。已修補為「如果你家也有倒貼習俗...根據文章最可能有什麼含意？」，讓學生結合假設情境和文章知識進行推論。"
    },
    {
        "row": 19,
        "status": "通過",
        "notes": "詮釋整合一致，選最佳標題題型正確，(A)「熱鬧有趣的年俗」能概括全文（年糕、春聯、倒福等各種年俗）。"
    },
    {
        "row": 20,
        "status": "通過",
        "notes": "「玩拼圖找因果」策略體現在因果結構的問法（為什麼...原因是），符合策略目標。答案(A)在文中有明確線索「躲避不及」，屬邊界推論。整體通過。"
    },
    {
        "row": 21,
        "status": "通過",
        "notes": "「學會語詞好閱讀」策略體現在詞義理解（「一頭霧水」），需從語境推論詞義，符合推論訊息。選項(A)「聽不太懂」正確。"
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

print("\n✅ Row 17-21 審查結果已寫回 Excel！")
print("累計：20/585 題完成")
