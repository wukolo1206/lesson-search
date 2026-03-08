import re

path = r'd:\test ch\閱讀教學策略查詢系統\index.html'
with open(path, encoding='utf-8') as f:
    content = f.read()

new_block = """      paragraphs.forEach(function (para, idx) {
        var p = document.createElement('p');
        p.className = (idx === 0 && titleLine === para) ? 'mb-4 font-bold text-lg' : 'mb-6 indent-8 min-h-[1.5em]';

        var gdriveMatch = para.match(/\\[GDrive:\\s*([a-zA-Z0-9_\\-]+)\\]/i);
        if (gdriveMatch) {
          var fileId = gdriveMatch[1];
          var tagPos = para.indexOf(gdriveMatch[0]);
          var beforeText = para.slice(0, tagPos).trim();
          var afterText = para.slice(tagPos + gdriveMatch[0].length).trim();
          p.style.textIndent = '0';
          if (beforeText) { var sp1 = document.createElement('span'); sp1.textContent = beforeText; p.appendChild(sp1); }
          var imgEl = document.createElement('img');
          imgEl.src = 'https://drive.google.com/uc?id=' + fileId;
          imgEl.className = 'max-w-full h-auto rounded shadow-sm mx-auto my-4 block border border-slate-600';
          imgEl.style.maxHeight = '400px';
          p.appendChild(imgEl);
          if (afterText) { var sp2 = document.createElement('span'); sp2.textContent = afterText; p.appendChild(sp2); }
        } else {
          p.textContent = para;
        }
        el.appendChild(p);
      });"""

pattern = r'      paragraphs\.forEach\(function \(para, idx\) \{.*?\}\);'
match = re.search(pattern, content, re.DOTALL)
if match:
    print('Found block. Replacing...')
    new_content = content[:match.start()] + new_block + content[match.end():]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('Done.')
else:
    print('Block NOT found.')
    idx = content.find('paragraphs.forEach')
    print('paragraphs.forEach at index:', idx)
