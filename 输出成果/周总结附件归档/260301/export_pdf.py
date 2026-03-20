import markdown

md_content = open(r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.md', 'r', encoding='utf-8').read()

html_content = '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
@page {
    size: A4;
    margin: 2cm;
}
body {
    font-family: SimSun, "Microsoft YaHei", sans-serif;
    font-size: 11pt;
    line-height: 1.6;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 10px 0;
}
th, td {
    border: 1px solid #666;
    padding: 6px 10px;
    text-align: left;
}
th {
    background-color: #e0e0e0;
}
h1 {
    color: #333;
    font-size: 18pt;
    text-align: center;
    border-bottom: 2px solid #333;
    padding-bottom: 10px;
}
h2 {
    color: #444;
    font-size: 14pt;
    border-bottom: 1px solid #999;
    padding-bottom: 5px;
}
h3 {
    color: #555;
    font-size: 12pt;
}
hr {
    border: none;
    border-top: 1px solid #ccc;
    margin: 15px 0;
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

try:
    from weasyprint import HTML
    HTML(string=html_content).write_pdf(r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.pdf')
    print("PDF exported successfully using WeasyPrint!")
except ImportError:
    print("WeasyPrint not available, trying pdfkit...")
    try:
        import pdfkit
        pdfkit.from_string(html_content, r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.pdf')
        print("PDF exported successfully using pdfkit!")
    except ImportError:
        print("Neither WeasyPrint nor pdfkit is available.")
        print("Please install one of them:")
        print("  pip install weasyprint")
        print("  or")
        print("  pip install pdfkit")
        with open(r'd:\LMAI\000打杂工具箱\输出成果\PPT分享文档\T-2026-001_精准谋划知识库_A4介绍.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        print("HTML file exported instead: T-2026-001_精准谋划知识库_A4介绍.html")
