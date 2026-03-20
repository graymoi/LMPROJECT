from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
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
output_pdf = Path(r'e:\李萌\输出成果\2026-03-10\P-2026-001_河西区老旧小区周边道路改造项目工作汇报_横板.pdf')
doc = SimpleDocTemplate(str(output_pdf), pagesize=(A4[1], A4[0]),  # Landscape A4
                     rightMargin=1.5*cm, leftMargin=1.5*cm,
                     topMargin=1.5*cm, bottomMargin=1.5*cm)

# Styles
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleCN', fontName=font_name, fontSize=16, textColor=colors.HexColor('#1a5490'),
                       spaceAfter=12, leading=20))
styles.add(ParagraphStyle(name='Heading2CN', fontName=font_name, fontSize=12, textColor=colors.HexColor('#2c7bb6'),
                       spaceAfter=8, spaceBefore=10, leading=16))
styles.add(ParagraphStyle(name='Heading3CN', fontName=font_name, fontSize=11, textColor=colors.HexColor('#444444'),
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

# Section 1: 项目概况
story.append(Paragraph('一、项目概况', styles['TitleCN']))

overview_data = [
    ['项目名称', '河西区老旧小区周边衔接道路基础设施配套项目'],
    ['项目类型', '城市更新-基础设施改造'],
    ['申报状态', '已申报，等待资金批复'],
    ['道路数量', '53条（20条+33条）'],
    ['改造长度', '36.69 km'],
    ['改造面积', '64.45万m²'],
    ['总投资', '19151.59万元'],
    ['工程费用', '15382.06万元'],
    ['拟申请中央资金', '15321.27万元']
]

overview_table = Table(overview_data, colWidths=[4*cm, 15*cm])
overview_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f4f8')),
    ('FONTNAME', (0, 0), (0, -1), font_name),
    ('FONTSIZE', (0, 0), (0, -1), 10),
    ('FONTNAME', (1, 0), (1, -1), font_name),
    ('FONTSIZE', (1, 0), (1, -1), 10),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
]))
story.append(overview_table)
story.append(Spacer(1, 0.3*cm))

# Section 2: 工作周期
story.append(Paragraph('二、工作周期', styles['TitleCN']))
story.append(Paragraph('项目周期：23天（2026年2月5日-2月27日）', styles['BodyCN']))

timeline_data = [
    ['2月5日', '2月6日', '2月10日', '2月13日', '2月25日', '2月27日'],
    ['项目启动', '部门对接', '方案编制', '项目分拆', '要件办理', '申报提交'],
    ['明确方向', '数据收集', '方案设计', '20+33分拆', '前置要件', '完成申报'],
    ['', '', '2.10-2.25 为图纸、实施方案、资金申请报告主要周期', '', '', '']
]

timeline_table = Table(timeline_data, colWidths=[3*cm]*6)
timeline_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, 3), (2, 3), colors.HexColor('#fff9e6')),
    ('SPAN', (2, 3), (4, 3)),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
story.append(timeline_table)
story.append(Spacer(1, 0.3*cm))

# Page break for section 3
story.append(PageBreak())

# Section 3: 多方协调工作
story.append(Paragraph('三、多方协调工作', styles['TitleCN']))
story.append(Paragraph('协调时间：2026年2月4日-2月28日（25天）', styles['BodyCN']))
story.append(Paragraph('协调单位：区发改委、区住建委、区城管委、市发改委', styles['BodyCN']))
story.append(Spacer(1, 0.2*cm))

coord_data = [
    ['时间', '协调单位', '协调内容', '工作成果'],
    ['2月4日', '项目前期准备', '在市发改委召开座谈会', '明确项目方向'],
    ['2月5日', '区发改委、住建委、城管委', '河西区发改委对接会议', '探讨项目申报方向，协调数据共享机制'],
    ['2月6日上午', '区住建委', '住建委座谈', '收集数据资料，排查城更范围'],
    ['2月6日下午', '区城管委', '河西区城管委调取数据', '拆分区管道路中的支路，调取背街小巷统计数据'],
    ['2月9日上午', '区住建委', '从河西区住建委取得城更台账以及21-25年老旧小区改造清单', '取得城更既有项目以及谋划储备项目的台账，以及21-25年老旧小区改造清单'],
    ['2月9日下午', '市发改委', '向市发改委汇报', '汇报整体进度以及项目拟申报资金'],
    ['2月10日', '方案编制启动', '承担项目全程的设计及资金申报工作', '《天津市河西区老旧小区周边衔接道路基础设施配套项目》前期资料包'],
    ['2月11-24日', '方案编制', '正式编制实施方案及资金申请报告', '以资金申请报告、实施方案为主要技术成果'],
    ['2月25-26日', '区住建委、市发改委', '前置要件办理', '前往河西区住建委、市发改委，协助办理各种前置要件'],
    ['2月26日', '成果完成', '完成所有成果', '天津市河西区老旧小区周边衔接道路基础设施配套项目完整成果包'],
    ['2月27日', '申报提交', '完成申报提交', '']
]

coord_table = Table(coord_data, colWidths=[2.5*cm, 4*cm, 7*cm, 5.5*cm])
coord_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('FONTNAME', (0, 0), (-1, 0), font_name),
    ('FONTSIZE', (0, 0), (-1, 0), 9),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
]))
story.append(coord_table)
story.append(Spacer(1, 0.2*cm))

coord_features = [
    '✅ 跨部门协调：市区城管委、市发改委、区发改委、区住建委多部门联动',
    '✅ 高频次对接：25天内完成9次重要协调对接',
    '✅ 全流程覆盖：从项目启动到申报提交，全流程协调',
    '✅ 问题导向：针对每个关键节点进行针对性协调'
]

for feature in coord_features:
    story.append(Paragraph(feature, styles['SmallCN']))

story.append(Spacer(1, 0.3*cm))

# Page break for section 4
story.append(PageBreak())

# Section 4: 重要成果
story.append(Paragraph('四、重要成果', styles['TitleCN']))

# 4.1 技术成果
story.append(Paragraph('4.1 技术成果', styles['Heading2CN']))

tech_data = [
    ['成果类型', '成果内容', '数量'],
    ['实施方案', '项目实施方案（P-2026-001-A/B）', '2份'],
    ['资金申请报告', '资金申请报告（P-2026-001-A/B）', '2份'],
    ['道路清单', '53条道路详细清单', '1份'],
    ['综合估算表', '工程综合估算表', '2份'],
    ['绩效目标表', '项目绩效目标表', '2份']
]

tech_table = Table(tech_data, colWidths=[4*cm, 10*cm, 3*cm])
tech_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
]))
story.append(tech_table)
story.append(Spacer(1, 0.2*cm))

# 4.2 文件成果
story.append(Paragraph('4.2 文件成果', styles['Heading2CN']))

file_data = [
    ['文件类别', '文件数量', '说明'],
    ['申报表', '2份', '项目申报表'],
    ['绩效目标', '2份', '项目绩效目标'],
    ['前期资料', '8份', '立项、批复、说明等'],
    ['资金材料', '2份', '财政承诺文件'],
    ['技术文件', '6份', '资金申请报告、实施方案等'],
    ['其他材料', '8份', '纳入重大项目库、信用报告等'],
    ['软建设', '5份', '政策支撑文件'],
    ['合计', '33份', 'P-2026-001-A/B共42份文件']
]

file_table = Table(file_data, colWidths=[3*cm, 3*cm, 11*cm])
file_table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (-1, -1), font_name),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f4f8')),
    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#fff9e6')),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
]))
story.append(file_table)
story.append(Spacer(1, 0.2*cm))

# 4.3 管理成果
story.append(Paragraph('4.3 管理成果', styles['Heading2CN']))

mgmt_features = [
    '✅ 项目谋划方法论：精准卡位、快速响应、避免重复、系统思维',
    '✅ 攻防问答要点：9个常见评审质疑与应对策略',
    '✅ 硬支撑清单：地图+实景图、无障碍设施列表、叠加老小区范围图、道路体检机制'
]

for feature in mgmt_features:
    story.append(Paragraph(feature, styles['SmallCN']))

story.append(Spacer(1, 0.3*cm))

# Section 5: 下一步工作计划
story.append(Paragraph('五、下一步工作计划', styles['TitleCN']))

story.append(Paragraph('近期工作（1-2周）', styles['Heading3CN']))
recent_work = [
    '• 跟踪中央预算内资金批复进度',
    '• 准备合同签署材料',
    '• 施工单位招标准备'
]
for item in recent_work:
    story.append(Paragraph(item, styles['BodyCN']))

story.append(Spacer(1, 0.1*cm))

story.append(Paragraph('中期工作（1-2月）', styles['Heading3CN']))
mid_work = [
    '• 资金到位后立即启动施工',
    '• 按照项目段分批实施',
    '• 建立项目管理机制'
]
for item in mid_work:
    story.append(Paragraph(item, styles['BodyCN']))

story.append(Spacer(1, 0.1*cm))

story.append(Paragraph('长期工作（3-6月）', styles['Heading3CN']))
long_work = [
    '• 项目完成后组织验收',
    '• 总结项目经验',
    '• 完成项目归档'
]
for item in long_work:
    story.append(Paragraph(item, styles['BodyCN']))

story.append(Spacer(1, 0.3*cm))

# Section 6: 合同签署
story.append(Paragraph('六、合同签署', styles['TitleCN']))
story.append(Paragraph('建议签署时间：立即签署合同，积极申请中央资金', styles['BodyCN']))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph('合同主要内容：', styles['Heading3CN']))

contract_items = [
    '• 项目范围：53条道路改造',
    '• 项目投资：19151.59万元',
    '• 拟申请中央资金：15321.27万元',
    '• 项目周期：20个月（2026年5月-2027年12月）',
    '• 质量标准：符合国家相关标准'
]

for item in contract_items:
    story.append(Paragraph(item, styles['BodyCN']))

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
