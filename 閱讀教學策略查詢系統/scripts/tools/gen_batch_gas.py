"""
gen_batch_gas.py
================
從 batch_scan_result.json 生成一個批次更新 GAS 函式。
執行：python scripts/tools/gen_batch_gas.py
"""
import json, re, os, sys

BASE = os.path.join(os.path.dirname(__file__), '..', '..')
SCAN_FILE = os.path.join(BASE, 'output', 'batch_scan_result_final.json')
OUT_FILE  = os.path.join(BASE, 'output', 'batchUpdateArticles.js')
SHEET_ID  = '1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI'

# (filename, article_index, sheet_name, spreadsheet_title)
MAPPING = [
    # ── 三年級（12/14，柿子/整理文具為圖表無法擷取）──
    ('10805國語篩選測驗3年級紙本測驗-答案.pdf',      1, '三年級', '小老鼠'),
    ('10805國語篩選測驗3年級紙本測驗-答案.pdf',      2, '三年級', '水鴨'),
    ('109.05篩選_國語三年級紙筆-答案卷.pdf',         1, '三年級', '馬露西'),
    ('109.05篩選_國語三年級紙筆-答案卷.pdf',         2, '三年級', '獅子蚊子'),
    ('110年05月篩選測驗-國語文3年級答案卷.pdf',       1, '三年級', '雨'),
    ('110年05月篩選測驗-國語文3年級答案卷.pdf',       2, '三年級', '登革熱'),
    ('111年05月篩選測驗-國語3年級答案版.pdf',         1, '三年級', '蝌蚪成長'),
    ('111年05月篩選測驗-國語3年級答案版.pdf',         2, '三年級', '玉米'),
    ('112年05月篩選測驗-國語3年級-答案版.pdf',        1, '三年級', '兄妹互動'),
    ('112年05月篩選測驗-國語3年級-答案版.pdf',        2, '三年級', '鴕鳥'),
    ('113年篩選測驗國語科3年級試卷答案卷.pdf',        1, '三年級', '馬與驢子'),
    ('114年篩選測驗國語科3年級試卷答案卷.pdf',        1, '三年級', '小光與雜貨店'),
    # ── 四年級（14/14 完整）──
    ('10805國語篩選測驗4年級紙本測驗-答案.pdf',       1, '四年級', '天上下的禮物'),
    ('10805國語篩選測驗4年級紙本測驗-答案.pdf',       2, '四年級', '奇怪的城堡'),
    ('109.05篩選_國語四年級紙筆-答案卷.pdf',          1, '四年級', '居家害蟲'),
    ('109.05篩選_國語四年級紙筆-答案卷.pdf',          2, '四年級', '哥倫布立蛋'),
    ('110年05月篩選測驗-國語文4年級答案卷.pdf',        1, '四年級', '紅姬緣椿象'),
    ('110年05月篩選測驗-國語文4年級答案卷.pdf',        2, '四年級', '特別的禮物'),
    ('111年05月篩選測驗-國語4年級答案版.pdf',          1, '四年級', '小黑蚊'),
    ('111年05月篩選測驗-國語4年級答案版.pdf',          2, '四年級', '冰島水怪傳說'),
    ('112年05月篩選測驗-國語4年級-答案版.pdf',         1, '四年級', '金銀盾牌'),
    ('112年05月篩選測驗-國語4年級-答案版.pdf',         2, '四年級', '灌肚猴'),
    ('113年篩選測驗國語科4年級試卷答案卷.pdf',         1, '四年級', '寄居蟹與海葵'),
    ('113年篩選測驗國語科4年級試卷答案卷.pdf',         2, '四年級', '動物放煙火'),
    ('114年篩選測驗國語科4年級試卷答案卷.pdf',         1, '四年級', '穿山甲'),
    ('114年篩選測驗國語科4年級試卷答案卷.pdf',         2, '四年級', '黑熊的壞心情'),
    # ── 五年級（14/21，114年#1有cid:1亂碼跳過）──
    ('10805國語篩選測驗5年級紙本測驗-答案.pdf',        1, '五年級', '海上船隻的領航者'),
    ('10805國語篩選測驗5年級紙本測驗-答案.pdf',        2, '五年級', '希利克鳥'),
    ('10805國語篩選測驗5年級紙本測驗-答案.pdf',        3, '五年級', '尋找金錶'),
    ('109.05篩選_國語五年級紙筆-答案卷.pdf',           1, '五年級', '聖誕老公公'),
    ('109.05篩選_國語五年級紙筆-答案卷.pdf',           2, '五年級', '金盞花的故事'),
    ('110年05月篩選測驗-國語文5年級答案卷.pdf',         1, '五年級', '擦鞋老闆的智慧'),
    ('110年05月篩選測驗-國語文5年級答案卷.pdf',         2, '五年級', '天寶・葛蘭汀'),
    ('110年05月篩選測驗-國語文5年級答案卷.pdf',         3, '五年級', '養蚵'),
    ('111年05月篩選測驗-國語5年級答案版.pdf',           1, '五年級', '川金絲猴'),
    ('111年05月篩選測驗-國語5年級答案版.pdf',           2, '五年級', '林昀儒的故事'),
    ('112年05月篩選測驗-國語5年級-答案版.pdf',          1, '五年級', '火星日落'),
    ('112年05月篩選測驗-國語5年級-答案版.pdf',          2, '五年級', '餐桌上的魚'),
    ('113年篩選測驗國語科5年級試卷答案卷.pdf',          1, '五年級', '蜜蜂的智慧'),
    ('113年篩選測驗國語科5年級試卷答案卷.pdf',          2, '五年級', '鴿子與果子'),
    # 114年5年級 #1 SKIP — 圖表Q&A格式含(cid:1)亂碼
    ('114年篩選測驗國語科5年級試卷答案卷.pdf',          2, '五年級', '礦工受困的故事'),
    # ── 六年級（15/21，109年#3為票價表跳過）──
    ('10805國語篩選測驗6年級紙本測驗-答案.pdf',         1, '六年級', '金錢草的啟示'),
    ('10805國語篩選測驗6年級紙本測驗-答案.pdf',         2, '六年級', '水都威尼斯'),
    ('10805國語篩選測驗6年級紙本測驗-答案.pdf',         3, '六年級', '陳澄波'),
    ('109.05篩選_國語六年級紙筆-答案卷.pdf',            1, '六年級', '臺灣俠醫林杰樑'),
    ('109.05篩選_國語六年級紙筆-答案卷.pdf',            2, '六年級', '母親與阿拉伯數字'),
    # 109年6年級 #3 SKIP — 只是「開心遊樂園票價表」問題說明文，非文章正文
    ('110年05月篩選測驗-國語文6年級答案卷.pdf',          1, '六年級', '太空船上的生活'),
    ('110年05月篩選測驗-國語文6年級答案卷.pdf',          2, '六年級', '林良的91歲生活'),
    ('110年05月篩選測驗-國語文6年級答案卷.pdf',          3, '六年級', '自助旅行建議'),
    ('111年05月篩選測驗-國語6年級答案版.pdf',            1, '六年級', '胡椒的祕密'),
    ('111年05月篩選測驗-國語6年級答案版.pdf',            2, '六年級', '尋找幸福的青鳥'),
    ('111年05月篩選測驗-國語6年級答案版.pdf',            3, '六年級', '家庭型態統計'),
    ('112年05月篩選測驗-國語6年級-答案版.pdf',           1, '六年級', '奧運場上的動人時刻'),
    ('112年05月篩選測驗-國語6年級-答案版.pdf',           2, '六年級', '孔蛛'),
    ('113年篩選測驗國語科6年級試卷答案卷.pdf',           1, '六年級', '抹香鯨'),
    ('113年篩選測驗國語科6年級試卷答案卷.pdf',           2, '六年級', '魯班與泰山'),
    ('114年篩選測驗國語科6年級試卷答案卷.pdf',           1, '六年級', '強森與朱莉的相遇'),
]


def fix_article(text):
    """移除文章末尾混入的答案選項（舊格式：答案數字+題號合併如 '120.'）"""
    m = re.search(r'\d{3}[\.\．]', text)
    if not m:
        return text
    before = text[:m.start()]
    # 從末端找最後一個句子結尾
    for end_char in ['」', '。', '！', '？']:
        pos = before.rfind(end_char)
        if pos > 0:
            return before[:pos + 1]
    return before.rstrip()


def esc_for_js(s):
    """把 Python 字串轉成 JS 單引號字串安全格式"""
    s = s.replace('\\', '\\\\')
    s = s.replace("'", "\\'")
    s = s.replace('\n', '\\n')
    return s


def main():
    with open(SCAN_FILE, encoding='utf-8') as f:
        data = json.load(f)

    # 建立查找表
    lookup = {}
    for r in data:
        for a in r['articles']:
            lookup[(r['file'], a['index'])] = a['text']

    updates = []
    for fname, idx, sheet, title in MAPPING:
        key = (fname, idx)
        if key not in lookup:
            print(f'找不到: {fname} #{idx}', flush=True)
            continue
        text = fix_article(lookup[key])
        updates.append({
            'sheet': sheet,
            'title': title,
            'text': text,
        })

    print(f'共 {len(updates)} 筆更新', flush=True)

    # 生成 GAS
    lines = [
        f'// 批次更新文章全文 — 共 {len(updates)} 篇（執行後可刪除）',
        'function batchUpdateArticles() {',
        f"  var ss = SpreadsheetApp.openById('{SHEET_ID}');",
        '  var updates = [',
    ]
    for u in updates:
        lines.append('    {')
        lines.append(f"      sheet: '{u['sheet']}',")
        lines.append(f"      title: '{u['title']}',")
        lines.append(f"      text: '{esc_for_js(u['text'])}'")
        lines.append('    },')
    lines.append('  ];')
    lines.append("""  var results = [];
  updates.forEach(function(u) {
    var sheet = ss.getSheetByName(u.sheet);
    if (!sheet) { results.push(u.title + ': 找不到工作表'); return; }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var ti = headers.indexOf('文本標題');
    var ci = headers.indexOf('文章全文');
    if (ti < 0 || ci < 0) { results.push(u.title + ': 找不到欄位'); return; }
    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ti]).trim() === u.title) {
        sheet.getRange(i + 1, ci + 1).setValue(u.text);
        updated++;
      }
    }
    results.push(u.title + ': 更新 ' + updated + ' 列');
  });
  Logger.log(results.join('\\n'));
  return results.join('\\n');
}
""")

    gas_code = '\n'.join(lines)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write(gas_code)
    print(f'GAS 已存至: {OUT_FILE}', flush=True)


if __name__ == '__main__':
    main()
