import markdown

md_content = open(r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.md', 'r', encoding='utf-8').read()

html_content = '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>精准谋划知识库 - T-2026-001</title>
<style>
@page {
    size: A4;
    margin: 1.5cm;
}
body {
    font-family: SimSun, "Microsoft YaHei", sans-serif;
    font-size: 10pt;
    line-height: 1.5;
    max-width: 210mm;
    margin: 0 auto;
    padding: 10mm;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 8px 0;
    font-size: 9pt;
}
th, td {
    border: 1px solid #666;
    padding: 4px 8px;
    text-align: left;
}
th {
    background-color: #e0e0e0;
}
h1 {
    color: #333;
    font-size: 16pt;
    text-align: center;
    border-bottom: 2px solid #333;
    padding-bottom: 8px;
    margin-bottom: 5px;
}
h2 {
    color: #444;
    font-size: 12pt;
    border-bottom: 1px solid #999;
    padding-bottom: 4px;
    margin-top: 12px;
    margin-bottom: 6px;
}
h3 {
    color: #555;
    font-size: 10pt;
    margin-top: 8px;
    margin-bottom: 4px;
}
hr {
    border: none;
    border-top: 1px solid #ccc;
    margin: 10px 0;
}
strong {
    color: #333;
}
ul {
    margin: 4px 0;
    padding-left: 20px;
}
li {
    margin: 2px 0;
}
p {
    margin: 4px 0;
}
</style>
</head>
<body>
''' + markdown.markdown(md_content, extensions=['tables']) + '''
</body>
</html>'''

with open(r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("HTML file exported: T-2026-001_精准谋划知识库_A4介绍.html")
print("")
print("To convert to PDF, you can:")
print("1. Open the HTML file in a browser and print to PDF")
print("2. Use a tool like wkhtmltopdf")
print("3. Use Microsoft Word to open and save as PDF")
