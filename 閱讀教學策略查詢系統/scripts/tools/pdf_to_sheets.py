"""
pdf_to_sheets.py
================
從國語測驗 PDF 自動擷取文章正文、偵測段落，
產生可直接在 GAS 編輯器執行的 updateArticleText() 函式。

使用方式：
    python scripts/tools/pdf_to_sheets.py <PDF路徑> <試算表文本標題> [工作表名稱]

範例：
    python scripts/tools/pdf_to_sheets.py "114年篩選測驗國語科3年級試卷.pdf" "小光與雜貨店" "三年級"

版面規則（以標準化評量試卷為基準）：
    - 右欄文章：x0 >= ARTICLE_X_MIN（預設 420）
    - 題目引導語（請閱讀以下）→ 文章起點
    - 題號行（數字+.）→ 文章終點
    - 段落首行縮排：x0 >= PARA_INDENT_MIN（預設 468），代表空兩格
    - 右欄左緣（x0 ≈ 433）→ 段落延續行
"""

import pdfplumber
import sys
import os
import re

# === 版面參數 ===
ARTICLE_X_MIN = 420        # 右欄起點
PARA_INDENT_MIN = 468      # 段落首行縮排門檻（x0 >= 此值 = 新段落）
LINE_MERGE_TOLERANCE = 8   # 同一行的 top 容差（px）
SHEET_ID = '1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI'
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), '..', '..', 'output')

# 文章開始標記（包含此字串的行之後才是文章）
ARTICLE_START_PATTERNS = [r'請閱讀以下', r'請閱讀下面']
# 文章結束標記（出現這些字串表示進入題目區）
ARTICLE_END_PATTERNS = [r'^\d{1,2}[\.\．]', r'^自$']
# 忽略的雜訊行
NOISE_PATTERNS = [r'^第\d+頁', r'^自$']


def words_to_lines(words, x_min):
    """將 pdfplumber word list 轉為行列表 [(top, x0_min, text), ...]"""
    right_words = [w for w in words if w['x0'] >= x_min]
    if not right_words:
        return []

    # 依 top 座標分組（容差 LINE_MERGE_TOLERANCE px）
    lines = {}
    for w in right_words:
        key = round(w['top'] / LINE_MERGE_TOLERANCE) * LINE_MERGE_TOLERANCE
        if key not in lines:
            lines[key] = {'top': w['top'], 'x0': w['x0'], 'words': []}
        lines[key]['words'].append(w)
        lines[key]['x0'] = min(lines[key]['x0'], w['x0'])

    result = []
    for key in sorted(lines.keys()):
        entry = lines[key]
        ws = sorted(entry['words'], key=lambda x: x['x0'])
        text = ''.join(w['text'] for w in ws).strip()
        if text:
            result.append((entry['top'], entry['x0'], text))
    return result


def extract_article(pdf_path):
    """擷取 PDF 中的文章段落，回傳 [(paragraph_text, is_letter_line), ...]"""
    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            all_lines.extend(words_to_lines(words, ARTICLE_X_MIN))

    # --- 找文章起點 ---
    article_start_idx = None
    for i, (top, x0, text) in enumerate(all_lines):
        for pat in ARTICLE_START_PATTERNS:
            if re.search(pat, text):
                article_start_idx = i + 1  # 從下一行開始
                break
        if article_start_idx is not None:
            break

    if article_start_idx is None:
        return None, "找不到文章起點（「請閱讀以下文章」）"

    # --- 找文章終點（題號出現） ---
    article_end_idx = len(all_lines)
    for i in range(article_start_idx, len(all_lines)):
        top, x0, text = all_lines[i]
        for pat in ARTICLE_END_PATTERNS:
            if re.match(pat, text):
                article_end_idx = i
                break
        if article_end_idx != len(all_lines):
            break

    article_lines = all_lines[article_start_idx:article_end_idx]

    # --- 過濾雜訊 ---
    filtered = []
    for top, x0, text in article_lines:
        skip = False
        for pat in NOISE_PATTERNS:
            if re.search(pat, text):
                skip = True
                break
        if not skip:
            filtered.append((top, x0, text))

    if not filtered:
        return None, "過濾後沒有文章內容"

    # --- 依首行縮排偵測段落 ---
    paragraphs = []
    current_para_lines = []

    for top, x0, text in filtered:
        is_new_para = (x0 >= PARA_INDENT_MIN)
        if is_new_para and current_para_lines:
            paragraphs.append(''.join(current_para_lines))
            current_para_lines = [text]
        else:
            current_para_lines.append(text)

    if current_para_lines:
        paragraphs.append(''.join(current_para_lines))

    # 清理段落：移除空白段落
    paragraphs = [p.strip() for p in paragraphs if p.strip()]
    return '\n'.join(paragraphs), None


def generate_gas_function(title, sheet_name, article_text):
    escaped = article_text.replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n')
    return f"""// 一次性文章分段更新 — {title}（執行後可刪除）
function updateArticleText() {{
  var ss = SpreadsheetApp.openById('{SHEET_ID}');
  var updates = [{{
    sheet: '{sheet_name}',
    titleCol: '文本標題',
    contentCol: '文章全文',
    title: '{title}',
    text: '{escaped}'
  }}];
  var results = [];
  updates.forEach(function(u) {{
    var sheet = ss.getSheetByName(u.sheet);
    if (!sheet) {{ results.push(u.title + ': 找不到工作表'); return; }}
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) {{ return String(h).trim(); }});
    var ti = headers.indexOf(u.titleCol), ci = headers.indexOf(u.contentCol);
    if (ti < 0 || ci < 0) {{ results.push(u.title + ': 找不到欄位'); return; }}
    var updated = 0;
    for (var i = 1; i < data.length; i++) {{
      if (String(data[i][ti]).trim() === u.title) {{
        sheet.getRange(i + 1, ci + 1).setValue(u.text);
        updated++;
      }}
    }}
    results.push(u.title + ': 更新 ' + updated + ' 列');
  }});
  Logger.log(results.join('\\n'));
  return results.join('\\n');
}}
"""


def main():
    if len(sys.argv) < 3:
        print('用法：python pdf_to_sheets.py <PDF路徑> <文本標題> [工作表名稱]')
        sys.exit(1)

    pdf_path = sys.argv[1]
    title = sys.argv[2]
    sheet_name = sys.argv[3] if len(sys.argv) > 3 else '三年級'

    if not os.path.exists(pdf_path):
        base = os.path.join(os.path.dirname(__file__), '..', '..')
        pdf_path = os.path.join(base, pdf_path)

    article, err = extract_article(pdf_path)
    if err:
        print(f'錯誤：{err}')
        sys.exit(1)

    paras = article.split('\n')
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 存預覽
    preview_path = os.path.join(OUTPUT_DIR, f'{title}_分段預覽.txt')
    with open(preview_path, 'w', encoding='utf-8') as f:
        for i, p in enumerate(paras, 1):
            f.write(f'【第{i}段】\n{p}\n\n')

    # 存 GAS 函式
    gas_code = generate_gas_function(title, sheet_name, article)
    gas_path = os.path.join(OUTPUT_DIR, f'{title}_updateGAS.js')
    with open(gas_path, 'w', encoding='utf-8') as f:
        f.write(gas_code)

    # 輸出摘要（用 json 避免編碼問題）
    import json
    summary = {
        'title': title,
        'para_count': len(paras),
        'paragraphs': [p[:50] + ('...' if len(p) > 50 else '') for p in paras],
        'preview_file': preview_path,
        'gas_file': gas_path
    }
    with open(os.path.join(OUTPUT_DIR, 'last_run_summary.json'), 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print('done')


if __name__ == '__main__':
    main()
