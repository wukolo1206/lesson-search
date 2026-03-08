import os
from pathlib import Path

def build_student_practice():
    """
    將 student_app 目錄下的 HTML, CSS, JS 打包成單一的
    student_practice.html 供 google apps script 使用。
    """
    base_dir = Path(r"d:\test ch\閱讀教學策略查詢系統")
    student_app_dir = base_dir / "student_app"
    
    html_path = student_app_dir / "index.html"
    css_path = student_app_dir / "style.css"
    js_path = student_app_dir / "app.js"
    output_path = base_dir / "student_practice.html"
    
    if not html_path.exists():
        print(f"Error: 找不到 {html_path}")
        return
        
    print("Reading HTML...")
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
        
    # Inject CSS
    if css_path.exists():
        print("Injecting CSS...")
        with open(css_path, 'r', encoding='utf-8') as f:
            css_content = f.read()
            html_content = html_content.replace(
                '<link rel="stylesheet" href="style.css">',
                f'<style>\n{css_content}\n</style>'
            )
            
    # Inject JS
    if js_path.exists():
        print("Injecting JS...")
        with open(js_path, 'r', encoding='utf-8') as f:
            js_content = f.read()
            html_content = html_content.replace(
                '<script src="app.js"></script>',
                f'<script>\n{js_content}\n</script>'
            )
            
    print(f"Writing to {output_path}...")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print("Build successful!")

if __name__ == "__main__":
    build_student_practice()
