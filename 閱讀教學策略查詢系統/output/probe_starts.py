
import pdfplumber, re, json, os

ARTICLE_X_MIN = 420
LINE_MERGE_TOLERANCE = 8

def words_to_lines(words, x_min):
    right_words = [w for w in words if w["x0"] >= x_min]
    if not right_words:
        return []
    lines = {}
    for w in right_words:
        key = round(w["top"] / LINE_MERGE_TOLERANCE) * LINE_MERGE_TOLERANCE
        if key not in lines:
            lines[key] = {"top": w["top"], "x0": w["x0"], "words": []}
        lines[key]["words"].append(w)
        lines[key]["x0"] = min(lines[key]["x0"], w["x0"])
    result = []
    for key in sorted(lines.keys()):
        entry = lines[key]
        ws = sorted(entry["words"], key=lambda x: x["x0"])
        text = "".join(w["text"] for w in ws).strip()
        if text:
            result.append((entry["top"], entry["x0"], text))
    return result

# Check one 5th grade and one 3rd grade PDF
tests = [
    r"D:	est ch\閱讀教學策略查詢系統\學習扶助考古題\五年級國語解答K年篩選測驗國語科5年級試卷答案卷.pdf",
    r"D:	est ch\閱讀教學策略查詢系統\學習扶助考古題\三年級國語答案K年篩選測驗國語科3年級試卷答案卷.pdf",
]
result = {}
for pdf_path in tests:
    fname = os.path.basename(pdf_path)
    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            all_lines.extend(words_to_lines(words, ARTICLE_X_MIN))
    starts = [(i, text) for i, (top, x0, text) in enumerate(all_lines) if re.search(r"請閱讀", text)]
    result[fname] = starts

with open("output/probe_starts.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print("done")
