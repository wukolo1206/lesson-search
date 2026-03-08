"""
批次儲存審查結果：Row 22-26
"""
import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

results = [
    {
        "row": 22,
        "status": "通過",
        "notes": "認知歷程（詮釋整合）一致，問主旨合理。「小偵探找原因」策略稍有偏差（主旨題非因果題），但整體可接受。"
    },
    {
        "row": 23,
        "status": "修補",
        "repaired": """根據文章，晶晶為什麼說七星鱸的尾巴「有一點一點的黑色」？
(A) 七星鱸尾巴上有明顯的黑點是牠們的特徵
(B) 七星鱸因為生病才長出黑色斑點
(C) 爸爸把墨汁塗在魚尾巴上做記號
(D) 七星鱸吃了太多黑色食物
答案：(A)""",
        "notes": "原遷移題問「七星鱸稀少的原因」，但提供的文本節錄中完全未提及此資訊，答案無法從文本找到依據，屬脫離文本的嚴重問題。已修補為文本有明確依據的因果推論題（晶晶的觀察→特徵說明的因果）。"
    },
    {
        "row": 24,
        "status": "通過",
        "notes": "「學會語詞好閱讀」策略一致，「合不攏嘴」在語境中推論為「非常開心」，符合推論訊息。"
    },
    {
        "row": 25,
        "status": "修補",
        "repaired": """這篇文章主要在介紹什麼？
(A) 七星鱸的外形特徵和民間傳說
(B) 如何讓魚的尾巴長出黑色印記
(C) 土地公神如何在井裡養魚
(D) 晶晶學習科學觀察的過程
答案：(A)""",
        "notes": "原遷移題(A)答案「要好好愛護並認識臺灣原生種魚類」帶入了環保立場，但文章主旨是介紹七星鱸外形特徵和民間傳說，並無呼籲保護的意涵。已修補，讓主旨題更貼近文章實際內容。"
    },
    {
        "row": 26,
        "status": "通過",
        "notes": "「玩拼圖找因果」策略一致，問盲人弄不清真相的原因，文章有明確因果線索（每人只摸一部分），答案(A)正確。"
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

print("\n✅ Row 22-26 審查結果已寫回 Excel！")
print("累計：25/585 題完成")
