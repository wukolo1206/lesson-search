"""
batch_pdf_scan.py
=================
批次掃描資料夾內所有 PDF，嘗試擷取文章正文並存成 JSON 供審查。

使用方式：
    python scripts/tools/batch_pdf_scan.py <資料夾路徑> [輸出JSON路徑]

範例：
    python scripts/tools/batch_pdf_scan.py "學習扶助考古題" output/batch_scan_result.json

輸出格式：
    [{
      "file": "114年篩選測驗國語科3年級試卷答案卷.pdf",
      "folder": "三年級國語答案",
      "grade": "三年級",
      "sheet": "三年級",
      "articles": [
        {"index": 1, "preview": "小光每天在雜貨店...", "text": "...full text..."}
      ],
      "error": null
    }]
"""

import pdfplumber
import sys
import os
import re
import json

# === 版面參數 ===
LINE_MERGE_TOLERANCE = 8

# 右欄（第1篇文章）：x0 >= 420
RIGHT_X_MIN = 420
RIGHT_PARA_INDENT = 468

# 左欄（第2/3篇文章）：36 <= x0 < 420
LEFT_X_MIN = 36
LEFT_X_MAX = 419
LEFT_PARA_INDENT = 55

# 只擷取「文章」，排除「圖文」（infographic）
ARTICLE_START_PATTERNS = [r'請閱讀以下文章', r'請閱讀下面文章', r'請閱讀下面的文章']
# 終點偵測：
#   \d{1,3}\.  → 題號 20. 120. 等
#   [(（][1-4][)）] → 答案選項 (1)(2)(3)(4) 在行首
#   請閱讀以下圖文 → 下一個圖文section開始
ARTICLE_END_PATTERNS = [
    r'^\d{1,3}[\.\．]',
    r'^[\(（][1-4][\)）]',
    r'請閱讀以下圖文',
    r'請閱讀下面圖文',
    r'^根據文章',          # 常見題目開頭（無題號時）
    r'^根據以上',
    r'^自$',
]
# 雜訊過濾（不應出現在文章中的行）
NOISE_PATTERNS = [
    r'^第\d+頁',
    r'^\d+頁，共\d+頁',
    r'\d+年基本學習內容標準化評量測驗',
    r'^自$',
]

GRADE_MAP = {
    '三年級': '三年級',
    '四年級': '四年級',
    '五年級': '五年級',
    '六年級': '六年級',
}


def words_to_lines(words, x_min, x_max=None):
    """將 word list 過濾至指定 x 範圍並組成行列表 [(top, x0_min, text), ...]"""
    filtered = [w for w in words if w['x0'] >= x_min and (x_max is None or w['x0'] <= x_max)]
    if not filtered:
        return []
    lines = {}
    for w in filtered:
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


def extract_articles_from_lines(all_lines, para_indent_min):
    """從已過濾的行列表中擷取所有文章，回傳 list of str"""
    # 找所有文章起點
    start_indices = []
    for i, (top, x0, text) in enumerate(all_lines):
        for pat in ARTICLE_START_PATTERNS:
            if re.search(pat, text):
                start_indices.append(i + 1)
                break

    articles = []
    for si in start_indices:
        end_idx = len(all_lines)
        for i in range(si, len(all_lines)):
            _, _, text = all_lines[i]
            for pat in ARTICLE_END_PATTERNS:
                if re.match(pat, text):
                    end_idx = i
                    break
            if end_idx != len(all_lines):
                break

        article_lines = all_lines[si:end_idx]

        # 過濾雜訊
        filtered = [(t, x, txt) for t, x, txt in article_lines
                    if not any(re.search(p, txt) for p in NOISE_PATTERNS)]

        if not filtered:
            continue

        # 段落偵測（首行縮排）
        paragraphs = []
        current = []
        for top, x0, text in filtered:
            if x0 >= para_indent_min and current:
                paragraphs.append(''.join(current))
                current = [text]
            else:
                current.append(text)
        if current:
            paragraphs.append(''.join(current))

        paragraphs = [p.strip() for p in paragraphs if p.strip()]
        if paragraphs:
            articles.append('\n'.join(paragraphs))

    return articles


def extract_all_articles(pdf_path):
    """擷取 PDF 中所有文章段落（右欄 + 左欄），回傳 list of str"""
    right_lines = []
    left_lines = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                right_lines.extend(words_to_lines(words, RIGHT_X_MIN))
                left_lines.extend(words_to_lines(words, LEFT_X_MIN, LEFT_X_MAX))
    except Exception as e:
        return [], str(e)

    if not right_lines and not left_lines:
        return [], '無法擷取文字（可能是掃描圖片 PDF）'

    right_articles = extract_articles_from_lines(right_lines, RIGHT_PARA_INDENT)
    left_articles  = extract_articles_from_lines(left_lines,  LEFT_PARA_INDENT)

    articles = right_articles + left_articles

    if not articles:
        return [], '找不到文章起點（「請閱讀以下文章」）'

    return articles, None


def detect_grade(folder_name):
    for key in GRADE_MAP:
        if key in folder_name:
            return key, GRADE_MAP[key]
    return None, None


def scan_folder(folder_path):
    results = []
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        print(f'找不到資料夾：{folder_path}')
        return results

    # 走訪所有子資料夾和 PDF
    for root, dirs, files in os.walk(folder_path):
        dirs.sort()
        rel_folder = os.path.relpath(root, folder_path)
        grade_label, sheet_name = detect_grade(rel_folder)
        # 若上層資料夾無年級，嘗試用 root 本身
        if not grade_label:
            grade_label, sheet_name = detect_grade(root)

        for fname in sorted(files):
            if not fname.lower().endswith('.pdf'):
                continue
            pdf_path = os.path.join(root, fname)
            articles, err = extract_all_articles(pdf_path)

            entry = {
                'file': fname,
                'folder': rel_folder,
                'grade': grade_label or '?',
                'sheet': sheet_name or '?',
                'articles': [],
                'error': err
            }

            for idx, art in enumerate(articles, 1):
                lines = art.split('\n')
                preview = lines[0][:60] + ('...' if len(lines[0]) > 60 else '')
                entry['articles'].append({
                    'index': idx,
                    'para_count': len(lines),
                    'preview': preview,
                    'text': art
                })

            results.append(entry)
            status = f"OK {len(articles)} 篇" if articles else f"NG {err}"
            print(f"  [{grade_label or '?'}] {fname}: {status}".encode('utf-8', errors='replace').decode('utf-8', errors='replace'), flush=True)

    return results


def main():
    if len(sys.argv) < 2:
        print('用法：python batch_pdf_scan.py <資料夾路徑> [輸出JSON路徑]')
        sys.exit(1)

    folder = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(__file__), '..', '..', 'output', 'batch_scan_result.json'
    )
    out_path = os.path.abspath(out_path)

    # 若路徑是相對的，從專案根目錄解析
    if not os.path.isabs(folder):
        base = os.path.join(os.path.dirname(__file__), '..', '..')
        folder = os.path.join(base, folder)

    # Force UTF-8 output on Windows
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    print(f'掃描資料夾：{folder}')
    results = scan_folder(folder)

    success = sum(1 for r in results if r['articles'])
    fail = len(results) - success
    total_articles = sum(len(r['articles']) for r in results)
    print(f'\n結果：{len(results)} 個 PDF，成功 {success} 個，失敗 {fail} 個，共 {total_articles} 篇文章')

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'已存至：{out_path}')


if __name__ == '__main__':
    main()
