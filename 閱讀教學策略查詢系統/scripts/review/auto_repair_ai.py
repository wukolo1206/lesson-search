import pandas as pd
import json
import re
import google.generativeai as genai
import time
import os

# --- 設定金鑰 (請替換成您的 API Key) ---
# 注意：為了安全起見，建議使用環境變數或確保腳本不會外流
# os.environ["GEMINI_API_KEY"] = "您的_API_KEY" 
# genai.configure(api_key=os.environ["GEMINI_API_KEY"])

file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
print("開始載入資料庫...")
df_trans = pd.read_excel(file_path, sheet_name='學習遷移題目')
df_orig = pd.read_excel(file_path, sheet_name='三年級')  # 暫時用三年級舉例

# 讀取審題標準 Prompt
with open(r'd:\test ch\閱讀教學策略查詢系統\docs\審題AI_Prompt.md', 'r', encoding='utf-8') as f:
    eval_prompt_template = f.read()

# 找出需要修復的題目 (這裡以找不到選項的題目為例)
def needs_repair(q_text):
    if pd.isna(q_text): return True
    return not any(x in str(q_text) for x in ['A', 'B', '1', '2', '('])

repair_count = 0
for idx, row in df_trans.iterrows():
    q_str = str(row['遷移題目'])
    
    if needs_repair(q_str):
        print(f"發現需要修復的題目 (Row {idx+2}): {row['原始案例標題']}")
        
        # 1. 抓取原始基準題目資料
        orig_title = str(row['原始案例標題'])
        matches = df_orig[df_orig['文本標題'] == orig_title]
        # (實務上這裡可能需要更精準的對齊邏輯，例如比對題目 ID)
        
        if len(matches) > 0:
            orig_row = matches.iloc[0]
            
            # 2. 構建給 AI 的修復 Prompt
            repair_prompt = f"""
你是一位專業的教育測驗出題專家。請根據我們制定的【AI 審題標準】(如下)，幫我「重寫/修復」一道具備學習遷移價值的題目。
這題原本生成的格式壞掉了（沒有提供四個選項），請你參考原題的邏輯，為這篇新的遷移文章重新出一個完美的題目，並包含 A, B, C, D 四個選項。

### 審題標準
{eval_prompt_template}

### 輸入資料
[原始文本 excerpts]：{str(orig_row['文本內容'])[:200]}...
[原始題目]：{orig_row['完整題目']}
[認知歷程]：{orig_row['認知歷程']}
[教學策略]：{orig_row['教學策略名稱']}

[遷移文本 excerpts]：{str(row['遷移文本內容'])[:200]}...
[壞掉的遷移題目]：{q_str}

### 你的任務：
請直接輸出修復後的「遷移題目」、「遷移選項A~D」、「正確答案」與「解析」。
請務必確保：
1. 選項具有「誘答一致性」，能反映原題的迷思。
2. 題目完全吻合該「認知歷程」與「教學策略」。

請以 JSON 格式回應：
{{
  "遷移題目": "...",
  "遷移選項A": "...",
  "遷移選項B": "...",
  "遷移選項C": "...",
  "遷移選項D": "...",
  "遷移正確答案": "A",
  "遷移解析": "..."
}}
"""
            # 3. 呼叫 API (這裡先註解掉，避免未設定金鑰時報錯)
            print("== 準備發送給 AI 的修復請求 ==")
            print(repair_prompt[:500] + "...\n(為節省空間已截斷)\n")
            
            # response = model.generate_content(repair_prompt)
            # new_data = json.loads(response.text)
            # df_trans.at[idx, '遷移題目'] = new_data['遷移題目']
            # df_trans.at[idx, '遷移選項A'] = new_data['遷移選項A'] ...等
            
            repair_count += 1
            if repair_count >= 1: # 測試時只印出一題
                break

# 實務上修復完畢後會把它存回 Excel
# df_trans.to_excel("修復版_學習扶助閱讀測驗試題分析.xlsx", index=False)
print("腳本執行完畢。")
