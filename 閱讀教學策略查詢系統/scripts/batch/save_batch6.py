"""
批次儲存審查結果：Row 27-31
"""
import pandas as pd

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

results = [
    {
        "row": 27,
        "status": "通過",
        "notes": "「學會語詞好閱讀」策略，問「你一言我一語」詞義，從文章爭論情境可推論（大家紛紛發表意見）。"
    },
    {
        "row": 28,
        "status": "通過",
        "notes": "詮釋整合一致，「看事情不能只看一部分」是盲人摸象故事的核心寓意，需整體閱讀理解。"
    },
    {
        "row": 29,
        "status": "通過",
        "notes": "提取訊息一致，「在船上放進石頭」是文中明確的步驟之一，可直接引用。誘答設計合理。"
    },
    {
        "row": 30,
        "status": "修補",
        "repaired": """文章的題目是「大象有多重」，關於這個篇名，下面哪個說法最適當？
(A) 篇名用疑問的方式，引起讀者對秤象方法的好奇心
(B) 篇名說明大象是世界上最重的動物
(C) 篇名告訴我們曹沖從小就喜歡大象
(D) 篇名說的是曹操向別人送了一頭大象
答案：(A)""",
        "notes": "教學策略「篇名有意思」要求理解篇名設計的意涵，但遷移題問的是曹沖的能力，與策略不符。已修補為分析篇名以問句方式引發好奇心的題型，符合「篇名有意思」策略目標。"
    },
    {
        "row": 31,
        "status": "通過",
        "notes": "假設推論題品質良好（「若忘記做記號→無法知道放多少石頭」），從文章步驟可推論後果。題型雖與原題略不同，但推論層次合理。"
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

print("\n✅ Row 27-31 審查結果已寫回 Excel！")
print("累計：30/585 題完成")
