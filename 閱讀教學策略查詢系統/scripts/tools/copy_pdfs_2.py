import os
import shutil
import re

src_root = r'd:\test ch\閱讀教學策略查詢系統\學習扶助考古題'
dst_dir = r'd:\test ch\字音形系統\pdf'

if not os.path.exists(dst_dir):
    os.makedirs(dst_dir)

count = 0
for root, dirs, files in os.walk(src_root):
    for f in files:
        if f.endswith('.pdf'):
            src_path = os.path.join(root, f)
            # 例如: 10805國語篩選測驗3年級紙本測驗-答案卷.pdf -> 108_G3.pdf
            m = re.search(r'(\d{3})\d{2}.*?([3456])年級', f)
            if m:
                year = m.group(1)
                grade = 'G' + m.group(2)
                new_name = f"{year}_{grade}.pdf"
                dst_path = os.path.join(dst_dir, new_name)
                print(f'Copying {f} -> {new_name}')
                shutil.copy2(src_path, dst_path)
                count += 1

print(f'Copied {count} files')
