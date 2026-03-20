import markdown
from weasyprint import HTML, CSS
from pathlib import Path

# Read the markdown file
md_file = Path(r'e:\李萌\输出成果\2026-03-10\P-2026-001_河西区老旧小区周边道路改造项目工作汇报_横板.md')
with open(md_file, 'r', encoding='utf-8') as f:
    md_content = f.read()

# Convert markdown to HTML
html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])

# Add CSS for landscape A4 and better styling
css_style = """
@page {
    size: A4 landscape;
    margin: 1.5cm;
}

body {
    font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.4;
    color: #333;
}

h1 {
    font-size: 18pt;
    color: #1a5490;
    border-bottom: 2px solid #1a5490;
    padding-bottom: 5px;
    margin-top: 20px;
    margin-bottom: 15px;
}

h2 {
    font-size: 14pt;
    color: #2c7bb6;
    border-bottom: 1px solid #2c7bb6;
    padding-bottom: 3px;
    margin-top: 15px;
    margin-bottom: 10px;
}

h3 {
    font-size: 12pt;
    color: #444;
    margin-top: 12px;
    margin-bottom: 8px;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 10pt;
}

table th, table td {
    border: 1px solid #ccc;
    padding: 6px 8px;
    text-align: left;
    vertical-align: top;
}

table th {
    background-color: #f0f4f8;
    font-weight: bold;
    color: #1a5490;
}

table tr:nth-child(even) {
    background-color: #f9f9f9;
}

ul, ol {
    margin: 8px 0;
    padding-left: 25px;
}

li {
    margin: 4px 0;
}

pre {
    background-color: #f5f5f5;
    border: 1px solid #ddd;
    padding: 10px;
    font-family: "Courier New", monospace;
    font-size: 9pt;
    overflow-x: auto;
}

code {
    font-family: "Courier New", monospace;
    background-color: #f0f0f0;
    padding: 2px 4px;
    border-radius: 3px;
}

blockquote {
    border-left: 4px solid #1a5490;
    padding-left: 15px;
    margin: 10px 0;
    color: #666;
    font-style: italic;
}

hr {
    border: none;
    border-top: 1px solid #ccc;
    margin: 15px 0;
}

.page-break {
    page-break-after: always;
}

.footnote {
    font-size: 9pt;
    color: #666;
    text-align: right;
    margin-top: 20px;
}
"""

# Combine HTML and CSS
full_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>{css_style}</style>
</head>
<body>
{html_content}
</body>
</html>
"""

# Generate PDF
output_pdf = Path(r'e:\李萌\输出成果\2026-03-10\P-2026-001_河西区老旧小区周边道路改造项目工作汇报_横板.pdf')
HTML(string=full_html).write_pdf(output_pdf)

print(f"PDF generated successfully: {output_pdf}")
