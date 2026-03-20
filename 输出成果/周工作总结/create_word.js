const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, 
        VerticalAlign, HeadingLevel } = require('docx');
const fs = require('fs');

const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };

const doc = new Document({
  styles: {
    default: { document: { run: { font: "微软雅黑", size: 24 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal",
        run: { size: 36, bold: true, color: "000000", font: "微软雅黑" },
        paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: "000000", font: "微软雅黑" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "000000", font: "微软雅黑" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "000000", font: "微软雅黑" },
        paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 } },
      { id: "Quote", name: "Quote", basedOn: "Normal",
        run: { size: 22, color: "666666", font: "微软雅黑" },
        paragraph: { spacing: { before: 60, after: 60 }, indent: { left: 360 } } }
    ]
  },
  numbering: {
    config: [
      { reference: "bullet-list",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbered-list-1",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbered-list-2",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  sections: [{
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    children: [
      new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("周工作总结（2026年3月9日-3月15日）")] }),
      
      new Paragraph({ style: "Quote", children: [new TextRun({ text: "总结周期：2026年3月9日（周一）至2026年3月15日（周六）", bold: true })] }),
      new Paragraph({ style: "Quote", children: [new TextRun({ text: "总结日期：2026-03-15", bold: true })] }),
      new Paragraph({ style: "Quote", children: [new TextRun({ text: "工作主题：背街小巷实施方案编制与AI协作场景探索", bold: true })] }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("📊 本周工作概览")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("核心成果")] }),
      
      new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [
        new TextRun({ text: "背街小巷实施方案编制", bold: true }),
        new TextRun("：完成成本重新梳理，保持250万预算，准备下周与兴业沟通")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [
        new TextRun({ text: "河西道路项目推进", bold: true }),
        new TextRun("：完成报价及前置工作梳理，周三上午开会，AI辅助会前准备和会议记录（完整工作流体系）")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [
        new TextRun({ text: "诺维信项目设计变更", bold: true }),
        new TextRun("：处理施工单位问题汇总，下周集中解决设计变更")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [
        new TextRun({ text: "经开区老旧小区改造", bold: true }),
        new TextRun("：AI整理前置资料，刘总周三开会无障碍，无需提前交代")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [
        new TextRun({ text: "AI协作场景探索", bold: true }),
        new TextRun("：形成4个已应用的AI协作场景，建立工作总结新规则")
      ]}),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("工作统计")] }),
      
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun({ text: "工作天数", bold: true }),
        new TextRun("：6天（3月9日-3月14日）")
      ]}),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun({ text: "主要项目", bold: true }),
        new TextRun("：4个（背街小巷、河西道路、诺维信、经开区老旧小区改造）")
      ]}),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun({ text: "会议次数", bold: true }),
        new TextRun("：1次（河西城管委）")
      ]}),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun({ text: "AI协作场景", bold: true }),
        new TextRun("：4个（前置工作梳理、资金申请报告、会前准备和会议记录、经开区老旧小区改造前置资料整理）")
      ]}),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("一、本周工作总结")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1. 项目进展")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("1.1 背街小巷项目（P-2026-002）")] }),
      
      new Table({
        columnWidths: [2340, 4680, 2340],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "时间", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "工作内容", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "成果", bold: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周一（3.9）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("合并天大院成果+我们院成果，编制实施方案初稿")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("重点核对"履职"关键词及重复工作界面")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周三中午（3.11）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("拿到兴业审核意见，组织天大院修改")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("天大晚上提供修改版本")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周三下午-周五（3.11-3.13）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("重新梳理成本版本")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("我们保持250万，天大100万")] })] })
            ]
          })
        ]
      }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "成本谈判关键节点：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("原提交：总计460万（我们250万+天大210万）")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("兴业砍价：总计~170万（我们100多万+天大70万）")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("重新梳理：总计350万（我们250万+天大100万）")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun({ text: "下周重点", bold: true }),
        new TextRun("：与兴业正式沟通，争取合理预算")
      ]}),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [
        new TextRun({ text: "成本调整难点", bold: true }),
        new TextRun("：详见附件3_背街小巷实施方案成本调整难点分析.md")
      ]}),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "小调研工作：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("周二（3.10）：完成小调研/前期调研")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("原计划周三/周四：约城管委办公室陈丰开会")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("实际结果：时间安排错不上，未开会")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("下周计划：雷处希望周一先深入研究，再和陈丰开会")] }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("1.2 河西道路项目（P-2026-001）")] }),
      
      new Table({
        columnWidths: [2340, 4680, 2340],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "时间", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "工作内容", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "成果", bold: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周二（3.10）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("报价及前置工作梳理")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("为3.11汇报准备")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周二（3.10）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("AI配合：资金申请报告文件清单+简单模板")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("模板完成")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周三（3.11）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("上午去河西城管委开会")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("会议记录完成")] })] })
            ]
          })
        ]
      }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "已完成事项：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("✅ 和金科长对接前期工作")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("✅ 提交签署合同申请")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("✅ 内部流程（盖章）- 走得比较晚，退错后了两天")] }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "待办事项：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("❓ 查资金申请报告有没有官方模板")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("✅ 前期策划咨询谋划（已投入，报价120w）")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 后续设计工作（看机会参加）")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⚠️ 前期谋划取费如何取？一直都没有解决")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⚠️ 河西城管委内部流程怎么走？需对接资金科室、公用事业中心")] }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("1.3 诺维信项目（P-2023-001）")] }),
      
      new Table({
        columnWidths: [2340, 4680, 2340],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "时间", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "工作内容", bold: true })] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                shading: { fill: "D5E8F0", type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "成果", bold: true })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周一（3.9）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("给龙海（施工单位）解决施工疑问")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("问题记录")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周二（3.10）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("问题汇总（施工单位发来的联系单）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("消防救援窗移动问题")] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("周五（3.13）")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 4680, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("设计变更：楼梯屋面改为钢结构屋面")] })] }),
              new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun("待确认整个楼梯作法")] })] })
            ]
          })
        ]
      }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "下周待办：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 下周一和张金宝讨论设计变更")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 如果有改图的内容，预计下周内完成")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 给施工单位的回复要在下周一或周二")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 改楼梯屋面为钢结构屋面")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 确认整个楼梯的作法")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 处理龙海的联络单")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("⏳ 解决消防救援窗移动问题")] }),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2. 重点工作")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("2.1 AI协作场景探索")] }),
      
      new Paragraph({ children: [new TextRun({ text: "已应用的AI协作场景：", bold: true })] }),
      new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [
        new TextRun({ text: "前置工作梳理", bold: true }),
        new TextRun("：电脑上有项目全要素资料，AI直接输出报告")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [
        new TextRun({ text: "资金申请报告", bold: true }),
        new TextRun("：AI生成文件清单+简单模板")
      ]}),
      new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [
        new TextRun({ text: "会议纪要", bold: true }),
        new TextRun("：AI做记录，提取待办事项，同步更新到项目的全要素（完整工作流的一个环节）")
      ]}),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ children: [new TextRun({ text: "关键洞察：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun("💡 每次写工作总结的时候，要把一些成果的前置资料包做成文档的附件")
      ]}),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [
        new TextRun("⭐ AI规则规范要做成包，给团队使用")
      ]}),
      
      new Paragraph({ children: [] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("2.2 工作总结规则优化")] }),
      
      new Paragraph({ children: [new TextRun({ text: "新规则：", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("工作总结格式 = 文字简版 + 附件")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("附件可以是工作总结的展开")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("附件可以是这周的可复用文档")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("可复用内容要有编号记录，在本文件夹内固定存储")] })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("d:/LMAI/000打杂工具箱/输出成果/周工作总结/03.09-03.15周工作总结.docx", buffer);
  console.log("Word文档已生成：03.09-03.15周工作总结.docx");
});
