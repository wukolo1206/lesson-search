import os
import shutil
import re

src_root = r'd:\test ch\閱讀教學策略查詢系統\學習扶助考古題'
dst_dir = r'd:\test ch\字音形系統\pdf'

if not os.path.exists(dst_dir):
    os.makedirs(dst_dir)

count = 0
grade_map = {'三': '3', '四': '4', '五': '5', '六': '6'}
for root, dirs, files in os.walk(src_root):
    for f in files:
        if f.endswith('.pdf') and '109' in f:
            src_path = os.path.join(root, f)
            print(f"Checking {f}")
            # 例如: 109.05篩選_國語三年級紙筆-答案卷.pdf -> 109_G3.pdf
            m = re.search(r'(109).*?([三四五六])年級', f)
            if m:
                year = m.group(1)
                grade_chinese = m.group(2)
                grade = 'G' + grade_map[grade_chinese]
                new_name = f"{year}_{grade}.pdf"
                dst_path = os.path.join(dst_dir, new_name)
                print(f'Copying {f} -> {new_name}')
                shutil.copy2(src_path, dst_path)
                count += 1

print(f'Copied {count} files')
