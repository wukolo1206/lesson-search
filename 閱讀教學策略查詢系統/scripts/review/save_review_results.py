"""
寫回第一批審查結果 (Row 2~6)
"""
import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

decisions = [
    {
        "excel_row": 2,
        "status": "已修補",
        "備註": "原題為提取訊息，但遷移題考的是詮釋整合(大意)，認知歷程錯位",
        "修補後題目": "根據文章，安安洗碗時，第一步先做了什麼？\n(A) 先擠洗碗精加水攪成泡泡水\n(B) 先拿菜瓜布直接搓洗碗盤\n(C) 先把碗盤全部泡進水桶\n(D) 先打電話問媽媽碗在哪\n答案：(A)\n解析：文章明確寫「首先，安安在小盆子裡擠了一些洗碗精，然後加水攪拌變成泡泡水」，直接從文中找即可。(B)跳過泡水步驟，(C)(D)無中生有。"
    },
    {
        "excel_row": 3,
        "status": "已修補",
        "備註": "推論成分不足，原題要求小偵探找原因(推論)但遷移題答案可直接提取",
        "修補後題目": "根據文章，安安洗完碗之後，爸爸為什麼忍不住鼓掌稱讚她？\n(A) 爸爸很少看到安安幫家事，覺得她長大了\n(B) 爸爸認為她洗碗的速度非常快\n(C) 因為安安把碗盤全部打破了\n(D) 安安叫爸爸快點幫她洗\n答案：(A)\n解析：文章說爸爸說「我們的小寶貝長大囉！」，需推論出爸爸稱讚的原因是看到孩子獨立成長，而非速度快。(B)過度推論，(C)(D)無中生有。"
    },
    {
        "excel_row": 4,
        "status": "通過",
        "備註": "詮釋整合+人物放大鏡策略完美對齊，誘答邏輯多元合理，優秀題目",
        "修補後題目": ""
    },
    {
        "excel_row": 5,
        "status": "已修補",
        "備註": "原題為提取訊息，但遷移題考的是詮釋整合(大意)，認知歷程錯位",
        "修補後題目": "根據文章，作者在菜市場跟媽媽一起做了什麼事？\n(A) 試著自己挑選魚貨\n(B) 買了一件新衣服\n(C) 跟老朋友在攤位聊天\n(D) 幫媽媽推菜籃車回家\n答案：(A)\n解析：文章說「也讓我試著挑挑看」，可直接從文中找到。(B)(D)無中生有，(C)是媽媽和老闆聊，非作者。"
    },
    {
        "excel_row": 6,
        "status": "通過",
        "備註": "推論訊息+小偵探找原因策略精準對齊，誘答邏輯多元，優秀題目",
        "修補後題目": ""
    },
]

print("載入 Excel...")
df = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')

# Ensure tracking columns are string type to prevent dtype errors
for col in ['審查狀態', '審查備註', '修補後題目']:
    if col in df.columns:
        df[col] = df[col].astype(str).replace('nan', '')

updated = 0
for d in decisions:
    row_idx = d["excel_row"] - 2
    if 0 <= row_idx < len(df):
        df.at[row_idx, '審查狀態'] = d.get("status", "通過")
        df.at[row_idx, '審查備註'] = d.get("備註", "")
        if d.get("修補後題目"):
            df.at[row_idx, '修補後題目'] = d["修補後題目"]
            df.at[row_idx, '遷移題目'] = d["修補後題目"]  # 同時覆寫原欄位
        updated += 1

print(f"寫回 {updated} 筆審查結果...")
with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='學習遷移題目', index=False)

print("✅ 完成！")
remaining = len(df[df['審查狀態'] == '待審查'])
print(f"📊 尚有 {remaining} 題待審查（已審 {585 - remaining} 題）")
