import pandas as pd
import json
import re

file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'

print("開始載入資料庫...")
df_trans = pd.read_excel(file_path, sheet_name='學習遷移題目')
print(f"共載入 {len(df_trans)} 題學習遷移題目。")

# 準備產出報告
report = {
    "total_questions": len(df_trans),
    "cognitive_mismatch": [],      # 認知歷程不符
    "strategy_mismatch": [],       # 教學策略不符
    "answer_distribution": {},     # 答案分佈異常
    "question_format_anomaly": [], # 題目格式異常 (找不到四個選項或答案)
    "missing_data": []             # 缺漏資料
}

# 統計答案
ans_counts = {'A': 0, 'B': 0, 'C': 0, 'D': 0, 'Other': 0}

for idx, row in df_trans.iterrows():
    q_id = f"Row {idx+2}: {str(row.get('原始案例標題', '未知'))[:10]}"
    
    # 1. 檢查缺漏資料
    if pd.isna(row.get('遷移題目')) or pd.isna(row.get('遷移文本內容')):
        report["missing_data"].append(q_id)
        continue
        
    # 2. 檢查認知歷程與策略的配對合理性
    cog = str(row.get('認知歷程', '')).strip()
    strat = str(row.get('教學策略', '')).strip()
    
    # 建立一個簡單的容錯對應表 (因為 AI 生成的標籤可能有微小差異)
    if "找一找" in strat and not any(x in cog for x in ["提取", "推論"]):
         report["cognitive_mismatch"].append(f"{q_id} (策略:{strat} vs 認知:{cog})")
    elif "推一推" in strat and not any(x in cog for x in ["推論", "詮釋", "整合"]):
         report["cognitive_mismatch"].append(f"{q_id} (策略:{strat} vs 認知:{cog})")
    elif "想一想" in strat and not any(x in cog for x in ["比較", "評估", "詮釋", "整合"]):
         report["cognitive_mismatch"].append(f"{q_id} (策略:{strat} vs 認知:{cog})")
         
    # 3. 解析題幹中的選項與解答
    # 因為所有的東西都擠在 "遷移題目" 這一格裡面
    q_text = str(row.get('遷移題目', ''))
    
    # 尋找答案：通常是 "答案：(A)" 或 "答案: A"
    ans_match = re.search(r'答案[：:]\s*\(?([A-D1-4])\)?', q_text)
    if ans_match:
        ans = ans_match.group(1).upper()
        # 簡單轉換 1234 為 ABCD
        ans_map = {'1':'A', '2':'B', '3':'C', '4':'D'}
        ans = ans_map.get(ans, ans)
        if ans in ans_counts:
            ans_counts[ans] += 1
        else:
            ans_counts['Other'] += 1
    else:
        report["question_format_anomaly"].append(f"{q_id} (找不到明確的正解標示)")
        ans_counts['Other'] += 1

    # 檢查是否有四個選項的特徵 (例如 A) B) C) D) 或 (A) (B) (C) (D) 或 (1) (2))
    has_options = bool(re.search(r'\([A-D1-4]\)|[A-D1-4]\)', q_text))
    if not has_options:
         report["question_format_anomaly"].append(f"{q_id} (題目內似乎沒有選項 A, B, C, D)")


# 輸出報告
print("="*50)
print("📊 學習遷移題目 - 自動化健檢報告 (Rule-based)")
print("="*50)
print(f"總題數: {report['total_questions']}")

print("\n[格式異常/找不到答案] (須人工回去補齊格式):")
if not report["question_format_anomaly"]:
    print("  ✅ 全數格式正確！")
for item in report["question_format_anomaly"][:10]:
    print(f"  - {item}")
if len(report["question_format_anomaly"]) > 10: print(f"  ...等共 {len(report['question_format_anomaly'])} 題")

print("\n[認知歷程/策略 標記可能衝突] (AI 可能生成錯認知層次):")
if not report["cognitive_mismatch"]:
    print("  ✅ 全數配對合理！")
for item in report["cognitive_mismatch"][:10]:
    print(f"  - {item}")
if len(report["cognitive_mismatch"]) > 10: print(f"  ...等共 {len(report['cognitive_mismatch'])} 題")

print("\n[整體答案分佈] (檢查是否有 AI 懶惰偏好特定選項):")
for k, v in report['answer_distribution'].items():
    if k != 'Other' or v > 0:
        print(f"  選項 {k}: {v} 題 ({v/max(1, report['total_questions'])*100:.1f}%)")

print("\n[缺漏資料 (完全空白)]:")
print(f"  共 {len(report['missing_data'])} 題")
