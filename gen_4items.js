const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };

function mkCell(text, opts = {}) {
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    columnSpan: opts.colSpan,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 340 },
      children: [new TextRun({
        text,
        bold: opts.bold || false,
        color: opts.color || "1A2332",
        font: { ascii: "Arial", eastAsia: "Microsoft YaHei" },
        size: opts.size || 21
      })]
    })]
  });
}

function heading(text, level = 1) {
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    children: [new TextRun({
      text, bold: true, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }
    })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 360 },
    children: [new TextRun({
      text, bold: opts.bold || false, color: opts.color || "1A2332",
      font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: opts.size || 21
    })]
  });
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: opts.ref || "bullets", level: opts.level || 0 },
    spacing: { after: 80, line: 360 },
    children: [new TextRun({
      text, bold: opts.bold || false,
      font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: opts.size || 21
    })]
  });
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 }
      }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: "1E3A5F", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, keepNext: false, keepLines: false } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "2C5282", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: false, keepLines: false } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "subBullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Q3四大工作事项（领导目标提炼版）", color: "94A3B8", size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "第 ", size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
          new TextRun({ text: " 页", size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })
        ]
      })] })
    },
    children: [

      // ===== 一、总体说明 =====
      heading("一、总体说明"),
      para("以下 4 大工作事项，是从领导给出的 Q3 目标方向中提炼、拔高、整合而成。你原来撰写的 6 项具体工作，全部作为子项纳入对应的大事项之下，确保方向对齐、颗粒度合适。"),

      // ===== 总览表 =====
      heading("二、四大工作事项总览"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [600, 3000, 5000, 2400],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("序号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 600, align: AlignmentType.CENTER }),
            mkCell("工作事项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 3000, align: AlignmentType.CENTER }),
            mkCell("业务价值（对应领导目标）", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 5000, align: AlignmentType.CENTER }),
            mkCell("纳入的原工作子项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2400, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("01", { width: 600, align: AlignmentType.CENTER, bold: true, color: "1E3A5F", fill: "FEF3C7" }),
            mkCell("数据架构与多维分析体系建设", { bold: true, width: 3000, fill: "FFFBEB" }),
            mkCell("经营/运营体系下的数据架构设计；以SBG为试点；产品、市场、项目三维财经管理能力建设", { width: 5000 }),
            mkCell("新产品成功看板优化", { width: 2400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("02", { width: 600, align: AlignmentType.CENTER, bold: true, color: "1E3A5F", fill: "DBEAFE" }),
            mkCell("规则中心与业财数据治理", { bold: true, width: 3000, fill: "EFF6FF" }),
            mkCell("规则中心的解法（把规则从代码中解放）；指标字典标准化建设；数据质量保障", { width: 5000 }),
            mkCell("阿米巴指标字典整理、阿米巴重点需求落地", { width: 2400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("03", { width: 600, align: AlignmentType.CENTER, bold: true, color: "1E3A5F", fill: "D1FAE5" }),
            mkCell("项目全闭环与预算管理", { bold: true, width: 3000, fill: "ECFDF5" }),
            mkCell("围绕项目维度的产品闭环管理（投产、预算执行、核算）；预算辅助支持；业财一体化", { width: 5000 }),
            mkCell("产研项目投产规划表、业财2.0-现金流子项", { width: 2400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("04", { width: 600, align: AlignmentType.CENTER, bold: true, color: "1E3A5F", fill: "FCE7F3" }),
            mkCell("AI赋能与业务价值深挖", { bold: true, width: 3000, fill: "FDF2F8" }),
            mkCell("AI作出有业务价值的场景（预算、指标字典）；深挖业务需求；日常需求落地与问题支持", { width: 5000 }),
            mkCell("销量预测模型MVP线上化", { width: 2400 }),
          ]}),
        ]
      }),

      // ===== 事项一 =====
      heading("三、事项一：数据架构与多维分析体系建设"),

      heading("3.1 业务价值", { level: 2 }),
      para("这是领导最强调的方向——从\"做单张报表\"升级到\"搭体系\"。不是为了做报表而做报表，而是搭一套能支持任意维度分析的数据架构，选 SBG 为试点跑通方法论。"),
      bulletItem("以 SBG 为试点，完成运营分析层面的数据需求深度调研", { ref: "bullets" }),
      bulletItem("设计 SBG 经营/运营体系下的数据架构（来源、加工、存储、应用全链路）", { ref: "bullets" }),
      bulletItem("建设产品、市场、项目三维财经管理能力——同一笔数据，三个维度都能切分看", { ref: "bullets" }),
      bulletItem("形成可复制的方法论，为推广到其他 BG 打基础", { ref: "bullets" }),

      heading("3.2 包含的子项（你原有的工作）", { level: 2 }),
      para("新产品成功看板优化（权重 10%）", { bold: true }),
      bulletItem("作为产品维度分析的核心载体，纳入三维财经管理体系中", { ref: "subBullets" }),
      bulletItem("看板 2.0 不只是功能升级，而是作为产品维度数据能力的展示窗口", { ref: "subBullets" }),

      heading("3.3 建议新增补充", { level: 2 }),
      bulletItem("SBG 运营分析数据需求调研报告（主动深挖需求的输出物）", { ref: "bullets" }),
      bulletItem("SBG 数据架构设计文档", { ref: "bullets" }),
      bulletItem("三维财经管理能力说明：产品维 / 市场维 / 项目维各自的分析主题和指标", { ref: "bullets" }),

      // ===== 事项二 =====
      heading("四、事项二：规则中心与业财数据治理"),

      heading("4.1 业务价值", { level: 2 }),
      para("把\"规则\"从代码里解放出来——现在阿米巴的分摊、口径都写死在代码里，改个规则要找 IT 开发。规则中心就是让业务人员能自己配置规则，不用改代码。"),
      bulletItem("完成阿米巴管报指标字典的标准化建设（指标字典 V1.0）", { ref: "bullets" }),
      bulletItem("梳理规则配置化的需求，为规则中心打基础（选 SBG 为试点梳理）", { ref: "bullets" }),
      bulletItem("影响阿米巴数据质量的需求全部落地", { ref: "bullets" }),
      bulletItem("确保数据质量——口径对齐、数据准确、问题可追溯", { ref: "bullets" }),

      heading("4.2 包含的子项（你原有的工作）", { level: 2 }),
      para("① 阿米巴指标字典整理（权重 40%）", { bold: true }),
      bulletItem("这是规则中心的基础——指标定义清楚了，规则才能配置化", { ref: "subBullets" }),
      bulletItem("完成与业务方、财务方的口径对齐", { ref: "subBullets" }),
      para("② 阿米巴重点需求落地（权重 20%）", { bold: true }),
      bulletItem("数据质量需求 100% 落地", { ref: "subBullets" }),
      bulletItem("体验类需求交付 ≥ 3 个", { ref: "subBullets" }),

      heading("4.3 建议新增补充", { level: 2 }),
      bulletItem("SBG 规则配置化需求梳理清单（为规则中心建设做输入）", { ref: "bullets" }),
      bulletItem("AI 辅助指标字典查询（对应领导提出的 AI 在指标字典上的应用）", { ref: "bullets" }),

      // ===== 事项三 =====
      heading("五、事项三：项目全闭环与预算管理"),

      heading("5.1 业务价值", { level: 2 }),
      para("打通项目从立项到核算的全生命周期数据——每个项目都算得清投入了多少、产出了多少、预算执行得怎么样。领导关注的是\"闭环\"，不是孤立的一张表。"),
      bulletItem("围绕项目维度，打通投产、预算执行、核算全链路", { ref: "bullets" }),
      bulletItem("配合领导推进预算相关工作（辅助支持角色）", { ref: "bullets" }),
      bulletItem("推动业财一体化——业务数据和财务数据不再是两张皮", { ref: "bullets" }),

      heading("5.2 包含的子项（你原有的工作）", { level: 2 }),
      para("① 产研项目投产规划表（权重 10%）", { bold: true }),
      bulletItem("作为项目闭环中的\"投产\"环节，纳入全链路管理", { ref: "subBullets" }),
      bulletItem("QA 知识沉淀 + 日常问题支持，保障项目投产顺利", { ref: "subBullets" }),
      para("② 业财 2.0 - 现金流子项（权重 5%）", { bold: true }),
      bulletItem("作为业财一体化的重要组成部分，与预算管理联动", { ref: "subBullets" }),

      heading("5.3 建议新增补充", { level: 2 }),
      bulletItem("项目全闭环数据链路说明（投产 → 预算执行 → 核算，每个环节的数据流）", { ref: "bullets" }),
      bulletItem("预算支持相关输出（配合领导预算工作的辅助材料）", { ref: "bullets" }),

      // ===== 事项四 =====
      heading("六、事项四：AI赋能与业务价值深挖"),

      heading("6.1 业务价值", { level: 2 }),
      para("AI 不是炫技，要解决实际业务问题。领导明确点了两个方向——预算和指标字典。同时要从\"被动接需求\"升级为\"主动深挖需求\"。"),
      bulletItem("AI 在预算上的应用场景：预算编制辅助、预算执行对比分析", { ref: "bullets" }),
      bulletItem("AI 在指标字典上的应用场景：智能口径查询、血缘追踪", { ref: "bullets" }),
      bulletItem("主动深挖业务需求：每月主动做 1-2 次业务侧调研", { ref: "bullets" }),
      bulletItem("日常需求落地与问题支持（基础运维保障）", { ref: "bullets" }),

      heading("6.2 包含的子项（你原有的工作）", { level: 2 }),
      para("销量预测模型 MVP 线上化（权重 15%）", { bold: true }),
      bulletItem("作为 AI 赋能的落地场景之一，与预算辅助联动", { ref: "subBullets" }),
      bulletItem("完成数据处理自动化、原始数据沉淀", { ref: "subBullets" }),
      bulletItem("建议扩展：从销量预测延伸到预算辅助（领导更关注的方向）", { ref: "subBullets" }),

      heading("6.3 建议新增补充", { level: 2 }),
      bulletItem("AI 辅助指标字典查询场景设计与原型", { ref: "bullets" }),
      bulletItem("AI 预算辅助场景设计（配合领导预算工作）", { ref: "bullets" }),
      bulletItem("业务需求深挖计划：月度调研清单", { ref: "bullets" }),

      // ===== 四、权重分配建议 =====
      heading("七、权重分配建议"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [600, 3500, 1200, 4700],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("序号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 600, align: AlignmentType.CENTER }),
            mkCell("工作事项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 3500, align: AlignmentType.CENTER }),
            mkCell("建议权重", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 1200, align: AlignmentType.CENTER }),
            mkCell("权重分配说明", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 4700, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("01", { width: 600, align: AlignmentType.CENTER, bold: true, fill: "FEF3C7" }),
            mkCell("数据架构与多维分析体系建设", { bold: true, width: 3500, fill: "FFFBEB" }),
            mkCell("30%", { width: 1200, align: AlignmentType.CENTER, bold: true, fill: "FEF3C7" }),
            mkCell("领导最关注的方向，含 SBG 试点+三维财经+看板优化", { width: 4700 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("02", { width: 600, align: AlignmentType.CENTER, bold: true, fill: "DBEAFE" }),
            mkCell("规则中心与业财数据治理", { bold: true, width: 3500, fill: "EFF6FF" }),
            mkCell("30%", { width: 1200, align: AlignmentType.CENTER, bold: true, fill: "DBEAFE" }),
            mkCell("含指标字典(40%)+需求落地(20%)，合计60%浓缩为30%", { width: 4700 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("03", { width: 600, align: AlignmentType.CENTER, bold: true, fill: "D1FAE5" }),
            mkCell("项目全闭环与预算管理", { bold: true, width: 3500, fill: "ECFDF5" }),
            mkCell("20%", { width: 1200, align: AlignmentType.CENTER, bold: true, fill: "D1FAE5" }),
            mkCell("含投产规划(10%)+现金流(5%)+新增预算支持", { width: 4700 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("04", { width: 600, align: AlignmentType.CENTER, bold: true, fill: "FCE7F3" }),
            mkCell("AI赋能与业务价值深挖", { bold: true, width: 3500, fill: "FDF2F8" }),
            mkCell("20%", { width: 1200, align: AlignmentType.CENTER, bold: true, fill: "FCE7F3" }),
            mkCell("含销量预测(15%)+新增AI场景(预算/指标字典)+深挖需求", { width: 4700 }),
          ]}),
        ]
      }),

      // ===== 五、一句话总结 =====
      heading("八、一句话总结"),
      para("以 SBG 为试点，围绕 \"数据架构 → 规则中心 → 项目闭环 → AI 赋能\" 这四条主线，把领导的 9 个方向和你原来的 6 项工作全部串起来，形成一个可落地、可衡量、可推广的 Q3 工作框架。"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\Q3四大工作事项（领导目标提炼版）.docx", buffer);
  console.log("Done");
});
