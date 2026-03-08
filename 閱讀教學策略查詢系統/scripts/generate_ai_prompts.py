"""
generate_ai_prompts.py (v2)
Reads the transfer question Excel sheet and generates AI prompts WITH original question reference,
so the AI can calibrate difficulty and style to match the original exactly.
"""
import pandas as pd

excel_file = '學習扶助閱讀測驗試題分析.xlsx'
df = pd.read_excel(excel_file, sheet_name='學習遷移題目', dtype=str).fillna('')

print(f"[OK] Read {len(df)} rows from the transfer sheet. Generating prompts v2...")


def get_strategy_instruction(strategy: str, process: str) -> str:
    strategy = strategy.strip()
    process = process.strip()

    vocab_keywords = ['拆字', '解詞', '詞義', '上下文', '推詞', '字義']
    if any(k in strategy for k in vocab_keywords):
        return (
            "【策略：詞彙理解】\n"
            "找出補充文本中一個有不同字義或較抽象的詞彙，設計「詞義理解」的單選題。\n"
            "選項陷阱設計：\n"
            "  (A) 正確答案：透過上下文推理的正確解釋。\n"
            "  (B) 陷阱(字面混淆)：字面相似但脫離情境的意思。\n"
            "  (C) 陷阱(常理干擾)：日常常見用法，但本文不適用。\n"
            "  (D) 另一個合理但錯誤的選項。\n"
            "答案請標示在最後一行：答案：(X)"
        )
    cause_keywords = ['因果', '原因', '結果', '觀點', '理由', '找原因', '找結果']
    if any(k in strategy for k in cause_keywords):
        return (
            "【策略：尋找關係（因果/觀點）】\n"
            "設計一道「因果關係」或「觀點立場」的單選題。\n"
            "選項陷阱設計：\n"
            "  (A) 正確答案：文章中明確或暗示的因果/觀點。\n"
            "  (B) 陷阱(張冠李戴)：有提到但不是該結果的原因。\n"
            "  (C) 陷阱(因果倒置)：把結果當成了原因。\n"
            "  (D) 另一個合理但錯誤的選項。\n"
            "答案請標示在最後一行：答案：(X)"
        )
    summary_keywords = ['重點句', '大意', '歸納', '刪除細節', '統整', '主旨', '段落整理']
    if any(k in strategy for k in summary_keywords) or process == '詮釋整合':
        return (
            "【策略：統整大意】\n"
            "設計一道「段落大意」或「全文主旨」的單選題。\n"
            "選項陷阱設計：\n"
            "  (A) 正確答案：涵蓋最核心觀念的精準敘述。\n"
            "  (B) 陷阱(以偏概全)：只提某個小細節，並非全篇中心。\n"
            "  (C) 陷阱(過度擴充)：合理但文章未提到的道理。\n"
            "  (D) 另一個合理但錯誤的選項。\n"
            "答案請標示在最後一行：答案：(X)"
        )
    predict_keywords = ['預測', '提問', '自我', '推測', '想像']
    if any(k in strategy for k in predict_keywords) or process == '比較評估':
        return (
            "【策略：自我提問/預測】\n"
            "設計一道「預測發展」或「評估合理性」的單選題。\n"
            "選項陷阱設計：\n"
            "  (A) 正確答案：根據文本線索的合理推斷。\n"
            "  (B) 陷阱(憑空想像)：無任何文本依據的猜測。\n"
            "  (C) 陷阱：部分合理但與文章方向相反。\n"
            "  (D) 另一個看似合理但不符線索的選項。\n"
            "答案請標示在最後一行：答案：(X)"
        )
    if process == '提取訊息':
        return (
            "【認知歷程：提取訊息】\n"
            "設計一道「找出正確敘述」的單選題。\n"
            "選項陷阱設計：\n"
            "  (A) 正確答案：與文章完全吻合的陳述。\n"
            "  (B) 陷阱：相近但某關鍵細節錯誤。\n"
            "  (C) 陷阱：人物、時間或地點有誤。\n"
            "  (D) 另一個讓人誤解的錯誤選項。\n"
            "答案請標示在最後一行：答案：(X)"
        )
    return (
        f"【認知歷程：{process}】\n"
        "設計一道符合此認知歷程的單選題，包含四個選項，選項間需有明顯教學鑑別力。\n"
        "答案請標示在最後一行：答案：(X)"
    )


output_lines = [
    "=" * 70,
    "學習遷移題庫：AI 出題提示詞草稿 (v2 含原題難度對齊)",
    "使用方法：將每一段貼入 Gemini 或 ChatGPT 取得題目草稿",
    "=" * 70,
    ""
]

# Set limit to False to generate all 585
LIMIT = 30
generated = 0

for idx, row in df.iterrows():
    if LIMIT and idx >= LIMIT:
        break

    transfer_text = str(row.get('遷移文本內容', ''))
    if not transfer_text or transfer_text.startswith('無可用'):
        continue

    orig_title   = str(row.get('原始案例標題', ''))
    orig_q       = str(row.get('原始題目', ''))       # <-- KEY: original question for calibration
    strategy     = str(row.get('教學策略', ''))
    process      = str(row.get('認知歷程', ''))
    grade        = str(row.get('年級', ''))

    instruction = get_strategy_instruction(strategy, process)
    text_preview = transfer_text[:800] + ("..." if len(transfer_text) > 800 else "")

    # Strip placeholder question if it's still the template
    orig_q_display = orig_q if (orig_q and '(針對' not in orig_q and 'ＯＯＯ' not in orig_q) else "(原始題目未填入，請依補充文本風格設計)"

    prompt = (
        f"\n{'=' * 60}\n"
        f"【第 {idx + 1} 筆 - {grade} | 策略：{strategy} | 歷程：{process}】\n"
        f"【原始案例】: {orig_title}\n"
        f"{'=' * 60}\n\n"
        f"你是台灣國小{grade}的國語科命題老師。\n"
        f"任務：根據「補充文本」出一道學習遷移的閱讀理解選擇題。\n\n"
        f"【重要：請嚴格對齊以下原始題目的難度、提問風格與選項長度】\n"
        f"--- 原始題目 ---\n"
        f"{orig_q_display}\n"
        f"--- 原始題目結束 ---\n\n"
        f"【補充文本（請根據此文出題）】\n"
        f"{text_preview}\n\n"
        f"【出題策略指示】\n"
        f"{instruction}\n\n"
        f"請注意：\n"
        f"1. 題幹風格要與原始題目相似（例如原題問「為什麼」，你也要問因果；原題問「主旨」，你也要問主旨）。\n"
        f"2. 選項長度和複雜度要與原題相當。\n"
        f"3. 陷阱選項的設計要符合「{strategy}」策略的誘答邏輯。\n\n"
        f"請直接輸出（不需要解釋）：\n"
        f"1. 題目\n"
        f"2. (A) (B) (C) (D) 四個選項\n"
        f"3. 最後一行：答案：(X)\n"
    )

    output_lines.append(prompt)
    output_lines.append("")
    generated += 1

output_path = 'ai_prompts_output_v2.txt'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print(f"[OK] Generated {generated} prompts (v2) -> {output_path}")
print("The prompts now include the original question as a difficulty reference.")
print("After reviewing, set LIMIT = False to generate all 585.")
