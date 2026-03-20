from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from pathlib import Path

# Register Chinese fonts (try common system fonts)
try:
    pdfmetrics.registerFont(TTFont('SimSun', 'C:\\Windows\\Fonts\\simsun.ttc', subfontIndex=0))
    pdfmetrics.registerFont(TTFont('SimHei', 'C:\\Windows\\Fonts\\simhei.ttf'))
    font_name = 'SimHei'
except:
    font_name = 'Helvetica'

# Create PDF with landscape A4
output_pdf = Path(r'e:\李萌\输出成果\2026-03-10\P-2026-001_河西区老旧小区周边道路改造项目_文件清单_横板.pdf')
doc = SimpleDocTemplate(str(output_pdf), pagesize=(A4[1], A4[0]),  # Landscape A4
                     rightMargin=1.5*cm, leftMargin=1.5*cm,
                     topMargin=1.5*cm, bottomMargin=1.5*cm)

# Styles
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleCN', fontName=font_name, fontSize=18, textColor=colors.HexColor('#1a5490'),
                       spaceAfter=12, leading=20))
styles.add(ParagraphStyle(name='Heading2CN', fontName=font_name, fontSize=14, textColor=colors.HexColor('#2c7bb6'),
                       spaceAfter=8, spaceBefore=10, leading=16))
styles.add(ParagraphStyle(name='Heading3CN', fontName=font_name, fontSize=12, textColor=colors.HexColor('#444444'),
                       spaceAfter=6, spaceBefore=8, leading=14))
styles.add(ParagraphStyle(name='BodyCN', fontName=font_name, fontSize=10, leading=14, spaceAfter=6))
styles.add(ParagraphStyle(name='SmallCN', fontName=font_name, fontSize=9, leading=12, spaceAfter=4))

# Build story
story = []

# Header info
header_data = [
    ['汇报单位：天津市河西区城市管理委员会', '汇报日期：2026年3月10日', '项目编号：P-2026-001']
]
header_table = Table(header_data, colWidths=[8*cm, 6*cm, 5*cm])
header_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
]))
story.append(header_table)
story.append(Spacer(1, 0.3*cm))

# Title
story.append(Paragraph('文件清单', styles['TitleCN']))
story.append(Spacer(1, 0.2*cm))

# 申报表
story.append(Paragraph('申报表', styles['Heading2CN']))
declare_data = [
    ['文件名称'],
    ['附件1：（新）城市更新2026年中央预算内投资项目申报表.xlsx']
]
declare_table = Table(declare_data, colWidths=[19*cm])
declare_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
story.append(declare_table)
story.append(Spacer(1, 0.2*cm))

# 绩效目标
story.append(Paragraph('绩效目标', styles['Heading2CN']))
perf_data = [
    ['文件名称'],
    ['附件2-绩效目标表.docx']
]
perf_table = Table(perf_data, colWidths=[19*cm])
perf_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
story.append(perf_table)
story.append(Spacer(1, 0.2*cm))

# 前期资料（8份）
story.append(Paragraph('前期资料（8份）', styles['Heading2CN']))
early_data = [
    ['序号', '文件名称'],
    ['1', '列入住建部改造计划证明文件.pdf'],
    ['2', '立项文件.png'],
    ['3', '实施方案批复.pdf'],
    ['4', '老旧小区年代说明.pdf'],
    ['5', '项目未获得过中央预算内投资支持承诺书.pdf'],
    ['6', '用地、环评、能评、规划、施工许可说明.pdf']
]
early_table = Table(early_data, colWidths=[2*cm, 17*cm])
early_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
]))
story.append(early_table)
story.append(Spacer(1, 0.2*cm))

# 资金材料
story.append(Paragraph('资金材料', styles['Heading2CN']))
fund_data = [
    ['文件名称'],
    ['财政承诺文件.pdf']
]
fund_table = Table(fund_data, colWidths=[19*cm])
fund_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
story.append(fund_table)
story.append(Spacer(1, 0.2*cm))

# 技术文件（6份）
story.append(Paragraph('技术文件（6份）', styles['Heading2CN']))
tech_data = [
    ['序号', '文件名称'],
    ['1', '资金申请报告-河西区20条道路改造项目0225.docx/pdf'],
    ['2', '资金申请报告-河西区33条道路改造项目0225.docx/pdf'],
    ['3', '天津市河西区老旧小区周边20条衔接道路基础设施配套项目-实施方案.pdf'],
    ['4', '天津市河西区老旧小区周边33条衔接道路基础设施配套项目-实施方案.docx/pdf'],
    ['5', '天津市河西区老旧小区周边20条衔接道路基础设施配套项目-调综合估算表.docx'],
    ['6', '天津市河西区老旧小区周边33条衔接道路基础设施配套项目-调综合估算表.docx']
]
tech_table = Table(tech_data, colWidths=[2*cm, 17*cm])
tech_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
]))
story.append(tech_table)
story.append(Spacer(1, 0.2*cm))

# 其他材料（8份）
story.append(Paragraph('其他材料（8份）', styles['Heading2CN']))
other_data = [
    ['序号', '文件名称'],
    ['1', '1-1.纳入国家重大项目库.png'],
    ['2', '1.纳入国家重大项目库.jpg'],
    ['3', '2.未列入严重失信主体名单.png'],
    ['4', '3.天津市河西区公用事业服务中心-信用报告.pdf'],
    ['5', '4.项目资金申请报告内容和附件真实性说明.pdf']
]
other_table = Table(other_data, colWidths=[2*cm, 17*cm])
other_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
]))
story.append(other_table)
story.append(Spacer(1, 0.2*cm))

# 软建设（5份）
story.append(Paragraph('软建设（5份）', styles['Heading2CN']))
soft_data = [
    ['序号', '文件名称'],
    ['1', '1.项目软建设.docx'],
    ['2', '2.2026-2030年城市道桥设施更新行动方案批复.png'],
    ['3', '3.天津"津城"城市更新规划指引.pdf'],
    ['4', '4.天津市城市更新行动计划（2023—2027年）.pdf'],
    ['5', '5.天津城市更新条例.doc']
]
soft_table = Table(soft_data, colWidths=[2*cm, 17*cm])
soft_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
]))
story.append(soft_table)
story.append(Spacer(1, 0.3*cm))

# Summary table
story.append(Paragraph('文件汇总', styles['Heading2CN']))
summary_data = [
    ['文件类别', '文件数量', '说明'],
    ['申报表', '1份', '项目申报表'],
    ['绩效目标', '1份', '项目绩效目标'],
    ['前期资料', '6份', '立项、批复、说明等'],
    ['资金材料', '1份', '财政承诺文件'],
    ['技术文件', '6份', '资金申请报告、实施方案等'],
    ['其他材料', '5份', '纳入重大项目库、信用报告等'],
    ['软建设', '5份', '政策支撑文件'],
    ['合计', '25份', 'P-2026-001-A/B共42份文件']
]
summary_table = Table(summary_data, colWidths=[3*cm, 3*cm, 13*cm])
summary_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fff9e6')),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
story.append(summary_table)
story.append(Spacer(1, 0.5*cm))

# Footer
footer_data = [
    ['汇报日期: 2026-03-10', '', '汇报单位: 建研院']
]
footer_table = Table(footer_data, colWidths=[6*cm, 9*cm, 4*cm])
footer_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
    ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('TEXTCOLOR', (0, 0), (-1, -1), colors.grey),
]))
story.append(footer_table)

# Build PDF
doc.build(story)

print(f"PDF generated successfully: {output_pdf}")
