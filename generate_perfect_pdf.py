import re
import subprocess
from pathlib import Path

def md_to_html_exam(md_path, html_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        text = f.read()
        
    # Split out the teacher part
    parts = text.split('## 教師版標準解答')
    student_text = parts[0]
    
    # Process specific sections
    # 1. 替換 【字】（　） 為 ruby
    # 我們將 【字】（　） 視為「要寫注音」
    # 為了美觀，加上特定的 CSS 類別
    student_text = re.sub(r'【(.*?)】（[　 ]*）', r'<ruby>\1<rt><span class="zhuyin-blank"></span></rt></ruby>', student_text)
    
    # 2. 改錯字的表格 1.（　） -> 1.<span class="err-grid"></span>
    student_text = re.sub(r'(\d+)\.（[　 ]*）', r'\1.<span class="err-grid"></span>', student_text)
    
    # 3. 選擇題 (　) 1. -> <span class="choice-box">(   )</span> 1.
    student_text = re.sub(r'\(.*?\)', r'<span class="choice-box">(   )</span>', student_text)
    
    # 4. 照樣寫短語 （ 陽光 ） -> <span class="fill-blank">陽光</span>
    # 答：（　　　） -> <span class="long-blank"></span>
    student_text = re.sub(r'答：（[　 ]*）', r'答：<span class="long-blank"></span>', student_text)
    
    # Build HTML
    html_template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="utf-8">
<title>期中評量</title>
<style>
@page {{
    size: B4 landscape;
    margin: 15mm;
}}
body {{
    writing-mode: vertical-rl;
    font-family: 'BiauKai', 'DFKai-SB', 'TW-Kai', serif;
    font-size: 16pt;
    line-height: 2.2;
    column-count: 2;
    column-gap: 20mm;
    margin: 0;
    height: 90vh; /* force column fill */
}}
h1, h2, h3 {{
    text-align: center;
    margin-right: 10px;
    margin-left: 10px;
}}
ruby {{
    ruby-position: over; /* Right side in vertical */
    ruby-align: center;
}}
rt {{
    font-size: 10pt;
}}
.zhuyin-blank {{
    display: inline-block;
    width: 12pt;
    height: 35pt;
    border: 1px solid black;
    vertical-align: text-top;
    margin-left: 2pt;
}}
.err-grid {{
    display: inline-block;
    width: 28pt;
    height: 28pt;
    border: 1px solid black;
    margin-top: 5pt;
    margin-bottom: 5pt;
}}
.choice-box {{
    display: inline-block;
    width: 30pt;
    font-family: sans-serif;
}}
.long-blank {{
    display: inline-block;
    width: 20pt;
    height: 150pt;
    border-right: 1px solid black;
    vertical-align: middle;
}}
.section-title {{
    font-weight: bold;
    font-size: 18pt;
    margin-right: 20px;
}}
.page-header {{
    font-size: 14pt;
    font-weight: bold;
}}
p {{
    margin: 5px 0;
    text-indent: 2em;
}}
.no-indent {{
    text-indent: 0;
}}
</style>
</head>
<body>
"""
    # 處理行
    lines = student_text.split('\\n')
    for line in lines:
        if line.strip() == '':
            continue
        if line.startswith('臺中市') or line.startswith('命題者') or line.startswith('校長') or line.startswith('四年') or line.startswith('第一面'):
            html_template += f'<div class="page-header">{line}</div>\\n'
        elif line.startswith('一、') or line.startswith('二、') or line.startswith('三、') or line.startswith('四、') or line.startswith('五、') or line.startswith('六、'):
            html_template += f'<div class="section-title">{line}</div>\\n'
        elif line.startswith('---'):
            html_template += '<div style="break-after: column;"></div>\\n'
        elif line.startswith('※'):
            html_template += f'<div class="no-indent">{line}</div>\\n'
        else:
            # 判斷是否為題號開頭
            if re.match(r'^[<span]|\d+\.', line.strip()) or '答：' in line:
                html_template += f'<div class="no-indent">{line}</div>\\n'
            else:
                html_template += f'<p>{line}</p>\\n'
                
    html_template += "</body></html>"
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_template)
        
    print(f"Generated {{html_path}}")

if __name__ == '__main__':
    md_to_html_exam('期中評量初稿.md', 'exam_layout.html')
    # Run md-to-pdf
    subprocess.run(['npx', '--yes', 'md-to-pdf', 'exam_layout.html'])
