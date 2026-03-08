"""
學習遷移題目：AI 邊審查邊修補腳本
===========================================
此程式會讀取 Excel 中的 585 道遷移題目，逐題進行品質審查。
若發現題目有問題（無選項、誘答力不足或格式錯誤），
則自動呼叫 Gemini API 依照審題標準重新修補，並覆寫回 Excel。

執行前請確認：
1. 已安裝 google-generativeai：pip install google-generativeai
2. 已在下方填入您的 Gemini API Key
"""

import pandas as pd
import json
import re
import time
import os
import google.generativeai as genai

# ===========================================================
# ⚙️  設定區 (請在此填入您的 API Key)
# ===========================================================
API_KEY = "請在這裡填入您的 Gemini API Key"
EXCEL_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
OUTPUT_PATH = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析_審查後.xlsx'
# 每次 API 呼叫中間等待的秒數 (避免超過 Rate Limit)
RATE_LIMIT_SLEEP = 4
# 設為 True 只跑前幾題測試，設為 False 才跑全部
DRY_RUN = True
DRY_RUN_LIMIT = 3  # 測試時只跑前 N 題

# ===========================================================
# 審題標準 (依照 docs/審題AI_Prompt.md 制訂)
# ===========================================================
REVIEW_CRITERIA = """
你是一位閱讀理解領域的專業出題與評量專家，專門為「學習扶助」學生出題。
你的任務是同時「審查」和「修補」一道學習遷移題目（類似題）。

你需要依照以下標準進行嚴格把關：

【第一層次：基礎正確性】
- 題幹清晰，提問無歧義
- 答案唯一且有文本依據
- 用詞符合該年級程度

【第二層次：選項誘答品質】 ← 這是最重要的！
- 必須提供完整的 A, B, C, D 四個選項
- 每個錯誤選項（誘答選項）必須對應一種學生常見的「迷思概念」：
  * 「望文生義」：表面有提到，但斷章取義
  * 「過度推論」：超出文章線索的過度猜測
  * 「局部事實」：只說對了一半，不夠全面
- 四個選項的文字長度應大致一致
- 正確答案不可永遠都是同一個選項

【第三層次：認知歷程與策略對齊】
- 「提取訊息」題：答案需白紙黑字在文中找到，不需推論
- 「直接推論」題：需連結兩個以上段落或句子才能回答
- 「詮釋整合」題：需理解整篇文章的主旨、作者觀點或情感

【第四層次：學習遷移品質】
- 解題路徑與難度必須和【原始題目】完全相同
- 錯誤選項的「陷阱邏輯」必須模仿原始題目的錯誤選項
- 表面情境必須和原文章不同（不能只是換個名字）
"""

def has_valid_options(q_text: str) -> bool:
    """檢查這題有沒有完整的選項"""
    if pd.isna(q_text) or len(str(q_text)) < 10:
        return False
    q = str(q_text)
    # 有效格式範例：(A) / A) / (1) / 1.
    has_abcd = bool(re.search(r'[\(\[]?[ABCDabcd1234][\)\]\.）]', q))
    has_answer = bool(re.search(r'答案|解答|正解', q))
    return has_abcd and has_answer


def review_and_repair(model, orig_data: dict, trans_row: pd.Series) -> dict | None:
    """
    呼叫 Gemini API 進行審查。
    如果題目有問題，AI 會直接修補並回傳完整的新題目。
    若題目通過審查，AI 會回傳 status: "pass"。
    """
    q_text = str(trans_row['遷移題目'])
    transfer_text = str(trans_row['遷移文本內容'])

    prompt = f"""
{REVIEW_CRITERIA}

=== 以下是你需要審查的資料 ===

【原始文章與題目（作為遷移品質的對照基準）】
- 原始文章標題：{orig_data.get('原始案例標題', '未知')}
- 原始題目：{orig_data.get('原始題目', '未知')}
- 認知歷程：{orig_data.get('認知歷程', '未知')}
- 教學策略：{orig_data.get('教學策略', '未知')}

【遷移文章（新的類似題應基於此文章出題）】
{transfer_text[:500]}
（如文章過長已截斷，請根據上述片段進行出題）

【AI 已生成的遷移題目（需要你審查）】
{q_text}

=== 你的任務 ===
1. 先根據上面所有的【審題標準】，判斷這道題目是否合格。
2. 如果完全合格（有四個選項、選項有誘答力、格式正確），請只回傳：
   {{"status": "pass", "reason": "...簡要說明合格的理由..."}}
3. 如果有任何不合格（缺少選項、誘答力不足、格式錯誤），請直接修補並重新出一道完美的題目，回傳：
   {{
     "status": "repaired",
     "repair_reason": "...說明哪裡有問題...",
     "遷移題目": "...(修補後的完整題幹)",
     "遷移選項A": "...",
     "遷移選項B": "...",
     "遷移選項C": "...",
     "遷移選項D": "...",
     "遷移正確答案": "A 或 B 或 C 或 D",
     "遷移解析": "...(請解釋為何正確，並說明每個錯誤選項是哪種迷思概念)"
   }}

重要提示：請只回傳 JSON，不要有任何其他文字說明。
"""

    try:
        response = model.generate_content(prompt)
        # 清理可能的 markdown code block
        raw = response.text.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'^```\s*', '', raw)
        raw = re.sub(r'\s*```$', '', raw)
        return json.loads(raw)
    except Exception as e:
        print(f"  ⚠️ API 呼叫或解析失敗: {e}")
        return None


def main():
    # 初始化 API
    if API_KEY == "請在這裡填入您的 Gemini API Key":
        print("❌ 錯誤：請先在腳本頂部的設定區填入您的 Gemini API Key！")
        return
    
    genai.configure(api_key=API_KEY)
    model = genai.GenerativeModel('gemini-2.0-flash')

    print("📂 載入 Excel 資料中...")
    df_trans = pd.read_excel(EXCEL_PATH, sheet_name='學習遷移題目')
    total = len(df_trans)
    print(f"✅ 共載入 {total} 題學習遷移題目。")

    if DRY_RUN:
        print(f"🧪 [測試模式] 只執行前 {DRY_RUN_LIMIT} 題，不實際修改檔案。")
    
    # 統計
    stats = {'pass': 0, 'repaired': 0, 'error': 0, 'skipped': 0}
    repaired_rows = []

    process_limit = DRY_RUN_LIMIT if DRY_RUN else total

    for idx, row in df_trans.iterrows():
        if idx >= process_limit:
            break

        print(f"\n{'='*50}")
        print(f"[{idx+1}/{process_limit}] 審查題目：{row.get('原始案例標題', '???')} → {str(row.get('遷移題目', ''))[:30]}...")
        
        # 組合原始題目資料 (直接從遷移表中取得)
        orig_data = {
            '原始案例標題': row.get('原始案例標題', ''),
            '原始題目': row.get('原始題目', ''),
            '認知歷程': row.get('認知歷程', ''),
            '教學策略': row.get('教學策略', '')
        }

        # 快速前置檢查：如果遷移文本內容都是空的，跳過
        if pd.isna(row.get('遷移文本內容')) or len(str(row.get('遷移文本內容', ''))) < 10:
            print(f"  ⏭️ 跳過（遷移文本內容為空）")
            stats['skipped'] += 1
            continue

        # 呼叫 AI 進行審查+修補
        result = review_and_repair(model, orig_data, row)
        
        if result is None:
            print(f"  ❌ AI 呼叫失敗，跳過此題。")
            stats['error'] += 1
        elif result.get('status') == 'pass':
            print(f"  ✅ 通過審查：{result.get('reason', '')[:50]}...")
            stats['pass'] += 1
        elif result.get('status') == 'repaired':
            print(f"  🔧 已修補：{result.get('repair_reason', '')[:80]}...")
            stats['repaired'] += 1
            repaired_rows.append({
                'idx': idx,
                'data': result
            })

        # Rate limit 休息
        print(f"  💤 休息 {RATE_LIMIT_SLEEP} 秒...")
        time.sleep(RATE_LIMIT_SLEEP)

    # 只有非測試模式才存回 Excel
    if not DRY_RUN and repaired_rows:
        print(f"\n💾 正在將 {len(repaired_rows)} 道修補好的題目存回 Excel...")
        for item in repaired_rows:
            idx = item['idx']
            d = item['data']
            df_trans.at[idx, '遷移題目'] = d.get('遷移題目', df_trans.at[idx, '遷移題目'])
            # 如果遷移表有獨立的選項欄位才需要以下幾行
            # df_trans.at[idx, '遷移選項A'] = d.get('遷移選項A', '')
            # df_trans.at[idx, '遷移選項B'] = d.get('遷移選項B', '')
            # df_trans.at[idx, '遷移選項C'] = d.get('遷移選項C', '')
            # df_trans.at[idx, '遷移選項D'] = d.get('遷移選項D', '')
            # df_trans.at[idx, '遷移正確答案'] = d.get('遷移正確答案', '')
            # df_trans.at[idx, '遷移解析'] = d.get('遷移解析', '')
        
        # 儲存到新檔案（確保原始資料安全）
        with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
            df_trans.to_excel(writer, sheet_name='學習遷移題目', index=False)
        print(f"✅ 已儲存至：{OUTPUT_PATH}")
    elif DRY_RUN:
        print(f"\n=== [測試模式已完成] ===")
        print(f"如果結果正確，請把腳本頂部的 DRY_RUN 改成 False 再重新執行。")
        print(f"\n修補建議預覽：")
        for item in repaired_rows:
            print(json.dumps(item['data'], ensure_ascii=False, indent=2))
    
    # 輸出統計
    print(f"\n{'='*50}")
    print(f"📊 審查完畢！統計結果：")
    print(f"  ✅ 通過：{stats['pass']} 題")
    print(f"  🔧 已修補：{stats['repaired']} 題")
    print(f"  ❌ 失敗：{stats['error']} 題")
    print(f"  ⏭️ 跳過：{stats['skipped']} 題")


if __name__ == "__main__":
    main()
