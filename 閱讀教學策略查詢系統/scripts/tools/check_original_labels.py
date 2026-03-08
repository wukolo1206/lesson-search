"""
原始題目「認知歷程」標籤健檢腳本
=====================================
掃描三~六年級原始題目，用題幹關鍵字推測「應該」的認知歷程，
並與 Excel 標記的認知歷程比對，輸出可能有誤的清單供人工確認。

判斷規則（依題幹關鍵字）：
  提取訊息：題幹問「誰/什麼/哪裡/何時/哪一個」，答案可直接從文中找到
  推論訊息：題幹問「為什麼/因為/所以/如何/代表什麼/言下之意」，需跨句推論
  詮釋整合：題幹問「主要/主旨/大意/全文/整篇/作者想說/整體而言」
  比較評估：題幹問「你認為/你贊成/哪個說法正確/如果你是」
"""

import pandas as pd
import re

EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
GRADE_SHEETS = ['三年級', '四年級', '五年級', '六年級']
OUTPUT_PATH = r'd:\test ch\閱讀教學策略查詢系統\docs\原題認知歷程健檢報告.md'

# ===== 關鍵字對應規則 =====
KEYWORD_RULES = [
    {
        "expected": "比較評估",
        "keywords": ["你認為", "你贊成", "你同意", "你覺得", "如果你是", "哪個說法", "下列哪個說法正確", "最適合"],
        "priority": 1  # 優先判斷（最高層次）
    },
    {
        "expected": "詮釋整合",
        "keywords": ["主要", "主旨", "大意", "整篇", "全文", "作者認為", "作者想說", "這篇文章告訴", "主要在說", "主要想告訴"],
        "priority": 2
    },
    {
        "expected": "推論訊息",
        "keywords": ["為什麼", "原因是", "因為", "所以", "才會", "才能", "怎麼知道", "代表什麼", "言下之意", "可以推論", "可以看出"],
        "priority": 3
    },
    {
        "expected": "提取訊息",
        "keywords": ["誰", "什麼人", "什麼地方", "在哪", "何時", "幾點", "哪一天", "哪一個", "第幾", "找出", "文章中提到"],
        "priority": 4  # 最低優先（最基本層次）
    },
]

def infer_cognitive_level(question_text: str) -> tuple[str, str]:
    """根據題幹關鍵字推測認知歷程，回傳 (推測結果, 命中的關鍵字)"""
    q = str(question_text)
    
    for rule in sorted(KEYWORD_RULES, key=lambda x: x["priority"]):
        for kw in rule["keywords"]:
            if kw in q:
                return rule["expected"], kw
    
    return "不確定", ""

# ===== 主程式 =====
print("載入原始題目資料庫...")
all_suspects = []  # 可能標籤有誤的題目
all_confirmed = 0
total = 0

for sheet in GRADE_SHEETS:
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
    except Exception as e:
        print(f"⚠️ 無法讀取分頁 {sheet}: {e}")
        continue
    
    for idx, row in df.iterrows():
        q_text = str(row.get('完整題目', ''))
        labeled_cog = str(row.get('認知歷程', '')).strip()
        strategy = str(row.get('教學策略名稱', '')).strip()
        title = str(row.get('文本標題', '')).strip()
        q_num = str(row.get('題號', idx))
        
        if not q_text or q_text == 'nan':
            continue
        
        total += 1
        inferred, matched_kw = infer_cognitive_level(q_text)
        
        # 判斷是否有衝突
        is_suspect = False
        reason = ""
        
        if inferred == "不確定":
            # 程式無法判斷，加入候選清單供人工確認
            is_suspect = True
            reason = "題幹無法辨識，建議人工確認"
        elif inferred != labeled_cog:
            # 程式判斷的認知歷程與標記不符
            is_suspect = True
            reason = f"題幹含「{matched_kw}」，推測應為【{inferred}】，但標記為【{labeled_cog}】"
        else:
            all_confirmed += 1
        
        if is_suspect:
            all_suspects.append({
                "年級": sheet,
                "題號": q_num,
                "文本標題": title,
                "標記的認知歷程": labeled_cog,
                "程式推測": inferred,
                "命中關鍵字": matched_kw,
                "原因": reason,
                "完整題目(前60字)": q_text[:60]
            })

# ===== 輸出報告 =====
print(f"\n掃描完成！共 {total} 題。")
print(f"✅ 確認無異議：{all_confirmed} 題")
print(f"⚠️ 可能有問題：{len(all_suspects)} 題")

# 寫成 Markdown 報告
lines = [
    "# 原始題目「認知歷程」標籤健檢報告",
    "",
    f"> **掃描日期**：2026-02-27  ",
    f"> **掃描範圍**：三~六年級，共 {total} 題  ",
    f"> ✅ 確認無異議：{all_confirmed} 題  ",
    f"> ⚠️ 可能有問題（建議人工確認）：{len(all_suspects)} 題  ",
    "",
    "---",
    "",
    "## ⚠️ 可能標籤有誤的題目清單",
    "",
    "> **注意**：程式判斷僅供參考，最終以您的教學專業為準。",
    "> 「不確定」表示程式無法從關鍵字辨識，請優先人工確認。",
    "",
    "| 年級 | 題號 | 文本標題 | 標記的認知歷程 | 程式推測 | 命中關鍵字 | 原因 | 題目節錄 |",
    "|---|---|---|---|---|---|---|---|",
]

for s in all_suspects:
    lines.append(
        f"| {s['年級']} | {s['題號']} | {s['文本標題']} | "
        f"{s['標記的認知歷程']} | {s['程式推測']} | "
        f"`{s['命中關鍵字']}` | {s['原因']} | {s['完整題目(前60字)']} |"
    )

with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f"\n📄 報告已儲存至：{OUTPUT_PATH}")
print("請打開報告，人工確認標記「⚠️」的題目是否真的有問題。")
