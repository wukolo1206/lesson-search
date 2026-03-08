import pandas as pd
import json

# 1. Load original data
with open('temp_output.json', 'r', encoding='utf-8') as f:
    original_data = json.load(f)

# 2. Load Supplementary Texts from the provided Excel file (all sheets)
supp_file = r"d:\test ch\閱讀教學策略查詢系統\學習扶助補充文本(教育部)\學扶補充文章1-6年級.xlsx"
all_supp_dfs = pd.read_excel(supp_file, sheet_name=None)

# Build a dictionary grading texts: {"三年級": [{"title":..., "content":...}], ...}
texts_by_grade = {}

for sheet_name, df in all_supp_dfs.items():
    for index, row in df.iterrows():
        grade = str(row.get("年級", "")).strip()
        if grade not in texts_by_grade:
            texts_by_grade[grade] = []
            
        texts_by_grade[grade].append({
            "title": str(row.get("篇名", "")).strip(),
            "content": str(row.get("文章全部文字內容", "")).strip()
        })

transfer_items = []
# Keep track of index per grade to cycle through texts
grade_indices = {g: 0 for g in texts_by_grade.keys()}

# 3. Generate Transfer Data (Article-to-Article Mapping)
for grade, questions in original_data.items():
    if grade == "參考資料": continue
    
    # Group original questions by article title
    articles = {}
    for q in questions:
        title = q.get("文本標題")
        if title not in articles:
            articles[title] = []
        articles[title].append(q)

    # Available supplementary texts for this grade
    supp_texts_for_grade = texts_by_grade.get(grade, [])
    
    for orig_title, orig_article_qs in articles.items():
        if len(supp_texts_for_grade) == 0:
            # If no supplementary texts, just create a placeholder
            for q in orig_article_qs:
                transfer_items.append({
                    "年度": q.get("年度"),
                    "年級": grade,
                    "原始案例標題": orig_title,
                    "原始題目": q.get("完整題目"),
                    "遷移文本內容": f"無可用【{grade}】補充文本",
                    "遷移題目": "請補充題目",
                    "教學策略": q.get("教學策略名稱"),
                    "認知歷程": q.get("認知歷程")
                })
            continue
            
        # Pick up to 3 different supplementary texts for this original article
        num_supps_to_pick = min(3, len(supp_texts_for_grade))
        picked_supps = []
        for _ in range(num_supps_to_pick):
            if grade not in grade_indices: grade_indices[grade] = 0
            
            supp = supp_texts_for_grade[grade_indices[grade] % len(supp_texts_for_grade)]
            picked_supps.append(supp)
            grade_indices[grade] += 1
            
        # For each picked supplementary text, duplicate all the original article's questions
        for supp in picked_supps:
            transfer_text = f"【{supp['title']}】\n{supp['content']}"
            
            for q in orig_article_qs:
                strategy = q.get("教學策略名稱", "")
                process = q.get("認知歷程", "")
                orig_q = q.get("完整題目", "")
                
                # Mock generation
                if strategy == "找出重點句" and process == "提取訊息":
                    transfer_q = f"(針對 {supp['title']}) 根據文章內容，下面哪個敘述正確？\n(1) ＯＯＯ\n(2) ＸＸＸ\n(3) ＯＸＯ\n(4) ＸＯＸ"
                elif process == "推論訊息":
                    transfer_q = f"(針對 {supp['title']}) 從文章中可以推論出什麼？\n(1) ＯＯＯ\n(2) ＸＸＸ\n(3) ＯＸＯ\n(4) ＸＯＸ"
                elif process == "詮釋整合":
                    transfer_q = f"(針對 {supp['title']}) 這篇文章主要想告訴我們什麼？\n(1) ＯＯＯ\n(2) ＸＸＸ\n(3) ＯＸＯ\n(4) ＸＯＸ"
                elif process == "比較評估":
                    transfer_q = f"(針對 {supp['title']}) 比較文章中的兩個觀點，哪一個最合理？\n(1) ＯＯＯ\n(2) ＸＸＸ\n(3) ＯＸＯ\n(4) ＸＯＸ"
                else:
                    transfer_q = f"(針對 {supp['title']}) 這題測試「{process}」能力：針對文本內容請選出適當答案。"

                transfer_items.append({
                    "年度": q.get("年度"),
                    "年級": grade,
                    "原始案例標題": orig_title,
                    "原始題目": orig_q,
                    "遷移文本內容": transfer_text,
                    "遷移題目": transfer_q,
                    "教學策略": strategy,
                    "認知歷程": process
                })

# 4. Write back to the main Excel file
output_file = '學習扶助閱讀測驗試題分析.xlsx'
df_transfer = pd.DataFrame(transfer_items)

with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_transfer.to_excel(writer, sheet_name='學習遷移題目', index=False)

print(f"Successfully generated {len(transfer_items)} similar questions using provided supplementary texts.")
print(f"Data saved to '學習遷移題目' sheet in {output_file}.")
