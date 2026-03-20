import markdown

def create_html(md_file, html_file, title):
    md_content = open(md_file, 'r', encoding='utf-8').read()
    
    html_content = '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>''' + title + '''</title>
<style>
@page {
    size: A4;
    margin: 1cm;
}
body {
    font-family: SimSun, "Microsoft YaHei", sans-serif;
    font-size: 8pt;
    line-height: 1.3;
    max-width: 210mm;
    margin: 0 auto;
    padding: 3mm;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 4px 0;
    font-size: 7pt;
}
th, td {
    border: 1px solid #666;
    padding: 2px 4px;
    text-align: left;
}
th {
    background-color: #e0e0e0;
}
h1 {
    color: #333;
    font-size: 12pt;
    text-align: center;
    border-bottom: 2px solid #333;
    padding-bottom: 3px;
    margin-bottom: 3px;
}
h2 {
    color: #444;
    font-size: 9pt;
    border-bottom: 1px solid #999;
    padding-bottom: 2px;
    margin-top: 6px;
    margin-bottom: 3px;
}
h3 {
    color: #555;
    font-size: 8pt;
    margin-top: 4px;
    margin-bottom: 2px;
}
hr {
    border: none;
    border-top: 1px solid #ccc;
    margin: 4px 0;
}
ul, ol {
    margin: 2px 0;
    padding-left: 15px;
}
li {
    margin: 1px 0;
}
p {
    margin: 2px 0;
}
blockquote {
    margin: 3px 0;
    padding: 3px 8px;
    border-left: 3px solid #667eea;
    background: #f0f0f0;
    font-style: italic;
}
code {
    font-family: Consolas, monospace;
    font-size: 7pt;
    background: #f5f5f5;
    padding: 1px 2px;
}
pre {
    font-family: Consolas, monospace;
    font-size: 7pt;
    background: #f5f5f5;
    padding: 3px;
    margin: 3px 0;
    white-space: pre-wrap;
}
strong {
    color: #333;
}
</style>
</head>
<body>
''' + markdown.markdown(md_content, extensions=['tables']) + '''
</body>
</html>'''
    
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML exported: {html_file}")

# 生成两个HTML文件
create_html(
    r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\P-2026-002_背街小巷精细化治理_A4介绍.md',
    r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\P-2026-002_背街小巷精细化治理_A4介绍.html',
    '背街小巷精细化治理项目'
)

create_html(
    r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\P-2026-001_河西道路项目_A4介绍.md',
    r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\P-2026-001_河西道路项目_A4介绍.html',
    '老旧小区周边道路改造项目'
)

print("")
print("HTML files created. Open in browser and print to PDF (Ctrl+P).")
