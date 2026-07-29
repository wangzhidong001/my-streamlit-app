const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const th = {
  borders, shading: { fill: "1E3A5F", type: ShadingType.CLEAR },
  margins: cellMargins,
  children: [new Paragraph({ children: [new TextRun({ text: "", bold: true, color: "FFFFFF", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 22 })] })]
};

function mkCell(text, opts = {}) {
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
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
      text,
      bold: true,
      font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }
    })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 360 },
    children: [new TextRun({
      text,
      bold: opts.bold || false,
      color: opts.color || "1A2332",
      font: { ascii: "Arial", eastAsia: "Microsoft YaHei" },
      size: opts.size || 21
    })]
  });
}

function bulletItem(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 360 },
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 })]
  });
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" },
          size: 21
        }
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
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Q3工作目标与领导目标对齐分析", color: "94A3B8", size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
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
      // ===== 一、总览 =====
      heading("一、目标总览"),
      para("本报告基于你撰写的 Q3 工作目标（6 项工作）与领导给出的目标方向进行对比分析，提炼业务价值，给出整合建议。"),

      // ===== 二、你写的Q3目标 =====
      heading("二、你撰写的 Q3 工作目标"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [500, 2600, 800, 2200, 2800],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("序号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 500, align: AlignmentType.CENTER }),
            mkCell("工作事项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2600, align: AlignmentType.CENTER }),
            mkCell("权重", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 800, align: AlignmentType.CENTER }),
            mkCell("目标", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2200, align: AlignmentType.CENTER }),
            mkCell("关键成果", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2800, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("1", { width: 500, align: AlignmentType.CENTER }),
            mkCell("阿米巴指标字典整理", { bold: true, width: 2600 }),
            mkCell("40%", { width: 800, align: AlignmentType.CENTER, fill: "FEF3C7" }),
            mkCell("完成阿米巴管报指标字典的标准化建设", { width: 2200 }),
            mkCell("指标字典V1.0、口径对齐、字典入库可查询", { width: 2800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("2", { width: 500, align: AlignmentType.CENTER }),
            mkCell("阿米巴重点需求落地", { bold: true, width: 2600 }),
            mkCell("20%", { width: 800, align: AlignmentType.CENTER, fill: "DBEAFE" }),
            mkCell("影响阿米巴数据质量需求全部落地", { width: 2200 }),
            mkCell("数据质量需求100%落地、体验类需求交付≥3个", { width: 2800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("3", { width: 500, align: AlignmentType.CENTER }),
            mkCell("销量预测模型MVP线上化", { bold: true, width: 2600 }),
            mkCell("15%", { width: 800, align: AlignmentType.CENTER, fill: "D1FAE5" }),
            mkCell("MVP场景上线", { width: 2200 }),
            mkCell("数据处理自动化、原始数据沉淀", { width: 2800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("4", { width: 500, align: AlignmentType.CENTER }),
            mkCell("产研项目投产规划表", { bold: true, width: 2600 }),
            mkCell("10%", { width: 800, align: AlignmentType.CENTER, fill: "FCE7F3" }),
            mkCell("QA知识沉淀及日常问题支持", { width: 2200 }),
            mkCell("输出投产规划表、完成其他日常需求", { width: 2800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("5", { width: 500, align: AlignmentType.CENTER }),
            mkCell("新产品成功看板优化", { bold: true, width: 2600 }),
            mkCell("10%", { width: 800, align: AlignmentType.CENTER, fill: "FCE7F3" }),
            mkCell("看板2.0上线", { width: 2200 }),
            mkCell("新产品成功看板V2.0上线", { width: 2800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("6", { width: 500, align: AlignmentType.CENTER }),
            mkCell("业财2.0-现金流子项", { bold: true, width: 2600 }),
            mkCell("5%", { width: 800, align: AlignmentType.CENTER, fill: "F3E8FF" }),
            mkCell("现金流子项上线", { width: 2200 }),
            mkCell("涉及系统完成改造、UAT测试、知识转移", { width: 2800 }),
          ]}),
        ]
      }),

      // ===== 三、领导目标拆解 =====
      heading("三、领导目标拆解与业务价值提炼"),
      para("领导给出的目标共 9 个方向，按业务价值拆解如下："),

      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [600, 3200, 5200],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("编号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 600, align: AlignmentType.CENTER }),
            mkCell("领导目标方向", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 3200, align: AlignmentType.CENTER }),
            mkCell("业务价值提炼", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 5200, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("1", { width: 600, align: AlignmentType.CENTER, fill: "FEF3C7" }),
            mkCell("经营、运营体系下数据架构设计；多维度分析能力建设，以SBG为例先做", { bold: true, width: 3200 }),
            mkCell("从\"做报表\"升级为\"建体系\"——不只做单张报表，而是搭数据架构，让数据能支持任意维度的分析（产品、市场、项目、组织）。选SBG先做是为了拿一个BG跑通，形成方法论，再推广。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("2", { width: 600, align: AlignmentType.CENTER, fill: "FEF3C7" }),
            mkCell("SBG运营分析层面的数据需求分析、数据架构设计", { bold: true, width: 3200 }),
            mkCell("具体落地动作：深入SBG业务，摸清他们要什么数据、怎么用，然后设计对应的数据架构。这是第1条的具体落地。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("3", { width: 600, align: AlignmentType.CENTER, fill: "DBEAFE" }),
            mkCell("规则中心的解法", { bold: true, width: 3200 }),
            mkCell("把\"规则\"从代码里解放出来——现在阿米巴的分摊、口径都写死在代码里，改个规则要找IT开发。规则中心就是让业务人员能自己配置规则，不用改代码。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("4", { width: 600, align: AlignmentType.CENTER, fill: "D1FAE5" }),
            mkCell("产品、市场、项目 三维财经管理", { bold: true, width: 3200 }),
            mkCell("数据要能同时从三个维度切分看同一个东西——产品维度（哪个产品赚钱）、市场维度（哪个市场/客户贡献大）、项目维度（哪个项目投入产出如何）。这是多维度分析能力的核心。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("5", { width: 600, align: AlignmentType.CENTER, fill: "FCE7F3" }),
            mkCell("日常需求落地、问题支持", { bold: true, width: 3200 }),
            mkCell("基础运维保障——确保日常业务用的系统不卡壳，问题及时响应，需求按时交付。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("6", { width: 600, align: AlignmentType.CENTER, fill: "FEF3C7" }),
            mkCell("围绕项目维度的产品闭环管理（投产、预算执行、核算）", { bold: true, width: 3200 }),
            mkCell("项目全生命周期打通——从立项预算→投产→执行→核算，全流程数据打通，让每个项目都算得清投入产出。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("7", { width: 600, align: AlignmentType.CENTER, fill: "F3E8FF" }),
            mkCell("预算：领导牵头，负责辅助", { bold: true, width: 3200 }),
            mkCell("预算体系支持——配合领导做预算相关的数据支持和系统支撑。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("8", { width: 600, align: AlignmentType.CENTER, fill: "D1FAE5" }),
            mkCell("AI作出有业务价值的场景（预算、指标字典）", { bold: true, width: 3200 }),
            mkCell("AI不只是炫技，要解决实际业务问题——预算编制辅助（AI帮做预算预测和对比）、指标字典（AI帮查指标口径、找血缘）。", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("9", { width: 600, align: AlignmentType.CENTER, fill: "DBEAFE" }),
            mkCell("深挖业务需求", { bold: true, width: 3200 }),
            mkCell("不是被动接需求，而是主动去业务那边问他们需要什么、痛在哪。把\"要我做\"变成\"我要做\"。", { width: 5200 }),
          ]}),
        ]
      }),

      // ===== 四、对齐分析 =====
      heading("四、你的目标 vs 领导目标：对齐与差距"),

      heading("4.1 已对齐的部分", { level: 2 }),
      bulletItem("指标字典（40%权重） ↔ 领导第3条\"规则中心的解法\" + 第8条\"AI在指标字典的应用：方向一致，你在做的指标字典就是规则中心的基础"),
      bulletItem("重点需求落地（20%权重） ↔ 领导第5条\"日常需求落地、问题支持\"：完全对齐"),
      bulletItem("产研项目投产规划表（10%权重） ↔ 领导第6条\"项目维度的产品闭环管理\"：方向一致，你做的投产表是项目闭环中的一环"),
      bulletItem("业财2.0现金流子项（5%权重） ↔ 领导第7条\"预算\"相关：部分对齐"),

      heading("4.2 存在差距的部分", { level: 2 }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [800, 3800, 4400],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("差距点", { bold: true, color: "FFFFFF", fill: "DC2626", width: 800, align: AlignmentType.CENTER }),
            mkCell("领导关注点", { bold: true, color: "FFFFFF", fill: "DC2626", width: 3800, align: AlignmentType.CENTER }),
            mkCell("你的目标中缺失/不足", { bold: true, color: "FFFFFF", fill: "DC2626", width: 4400, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("数据架构", { width: 800, align: AlignmentType.CENTER, fill: "FEF2F2", bold: true }),
            mkCell("经营/运营体系下的数据架构设计，以SBG为例", { width: 3800 }),
            mkCell("你的目标偏向具体功能，没有明确提到\"架构设计\"这个层面的工作", { width: 4400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("多维度分析", { width: 800, align: AlignmentType.CENTER, fill: "FEF2F2", bold: true }),
            mkCell("产品、市场、项目 三维财经管理", { width: 3800 }),
            mkCell("你做了产品维度（新产品成功看板）和项目维度（投产规划表），但没有明确提到\"三维打通\"这个整体思路", { width: 4400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("SBG先行", { width: 800, align: AlignmentType.CENTER, fill: "FEF2F2", bold: true }),
            mkCell("以SBG为例先做运营分析", { width: 3800 }),
            mkCell("你的目标中没有明确提到以哪个BG为重点推进", { width: 4400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("AI场景", { width: 800, align: AlignmentType.CENTER, fill: "FEF2F2", bold: true }),
            mkCell("AI在预算、指标字典上的业务价值场景", { width: 3800 }),
            mkCell("你做了销量预测模型（AI方向），但领导更关注AI在预算和指标字典上的应用", { width: 4400 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("深挖需求", { width: 800, align: AlignmentType.CENTER, fill: "FEF2F2", bold: true }),
            mkCell("主动深挖业务需求", { width: 3800 }),
            mkCell("你的目标偏执行层面，没有体现主动深挖业务需求的动作", { width: 4400 }),
          ]}),
        ]
      }),

      // ===== 五、整合建议 =====
      heading("五、整合后的 Q3 工作目标建议"),
      para("建议在保留你原有 6 项工作的基础上，补充以下内容使目标更贴合领导期望："),

      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [500, 2400, 900, 5200],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("序号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 500, align: AlignmentType.CENTER }),
            mkCell("工作事项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2400, align: AlignmentType.CENTER }),
            mkCell("建议权重", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 900, align: AlignmentType.CENTER }),
            mkCell("整合说明（贴合领导目标）", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 5200, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("1", { width: 500, align: AlignmentType.CENTER }),
            mkCell("阿米巴指标字典 + 规则中心基础", { bold: true, width: 2400 }),
            mkCell("35%", { width: 900, align: AlignmentType.CENTER, fill: "FEF3C7" }),
            mkCell("原有指标字典不变，补充：①以SBG为试点梳理规则配置化的需求，为规则中心打基础；②AI辅助指标字典查询（对应领导第3、8条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("2", { width: 500, align: AlignmentType.CENTER }),
            mkCell("SBG运营分析数据架构设计", { bold: true, width: 2400, fill: "F0FDF4" }),
            mkCell("20%", { width: 900, align: AlignmentType.CENTER, fill: "D1FAE5" }),
            mkCell("新增：以SBG为例做运营分析的数据需求调研 + 数据架构设计。这是领导最关注的方向之一（对应领导第1、2、4条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("3", { width: 500, align: AlignmentType.CENTER }),
            mkCell("阿米巴重点需求落地", { bold: true, width: 2400 }),
            mkCell("15%", { width: 900, align: AlignmentType.CENTER, fill: "DBEAFE" }),
            mkCell("原有不变（对应领导第5条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("4", { width: 500, align: AlignmentType.CENTER }),
            mkCell("项目维度产品闭环管理", { bold: true, width: 2400 }),
            mkCell("10%", { width: 900, align: AlignmentType.CENTER, fill: "FCE7F3" }),
            mkCell("把投产规划表升级为项目全闭环：投产+预算执行+核算打通（对应领导第6条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("5", { width: 500, align: AlignmentType.CENTER }),
            mkCell("AI业务场景落地", { bold: true, width: 2400 }),
            mkCell("10%", { width: 900, align: AlignmentType.CENTER, fill: "D1FAE5" }),
            mkCell("把销量预测升级为更贴合领导期望的AI场景：①预算辅助（领导牵头我辅助）②AI指标字典查询（对应领导第7、8条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("6", { width: 500, align: AlignmentType.CENTER }),
            mkCell("三维财经管理 + 深挖业务需求", { bold: true, width: 2400, fill: "F0FDF4" }),
            mkCell("5%", { width: 900, align: AlignmentType.CENTER, fill: "F3E8FF" }),
            mkCell("新产品成功看板（产品维）+ 投产规划表（项目维）+ 新增市场维分析，形成三维；每月主动做1-2次业务需求深挖（对应领导第4、9条）", { width: 5200 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("7", { width: 500, align: AlignmentType.CENTER }),
            mkCell("业财2.0现金流 + 预算支持", { bold: true, width: 2400 }),
            mkCell("5%", { width: 900, align: AlignmentType.CENTER, fill: "F3E8FF" }),
            mkCell("原有现金流子项 + 预算辅助支持（对应领导第7条）", { width: 5200 }),
          ]}),
        ]
      }),

      // ===== 六、总结 =====
      heading("六、总结"),
      para("核心思路：把你原本的 6 项工作和领导的 9 个方向做对齐，关键是以下 3 个升级："),
      bulletItem("从\"做具体功能\"升级到\"搭架构\"——不只做指标字典，而是建规则中心；不只做单张报表，而是做SBG的数据架构"),
      bulletItem("从\"单维度\"升级到\"三维打通\"——产品、市场、项目三个维度都要能切分看数据"),
      bulletItem("从\"被动接需求\"升级到\"主动挖价值\"——AI场景要贴近业务（预算、指标字典），主动去业务那边深挖需求"),
      para("建议在和领导对齐时，重点强调：以 SBG 为试点，先跑通数据架构+规则中心+三维财经管理这个组合拳，形成方法论后再推广到其他BG。"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\Q3工作目标对齐分析.docx", buffer);
  console.log("Done");
});
