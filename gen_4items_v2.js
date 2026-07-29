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
        text, bold: opts.bold || false, color: opts.color || "1A2332",
        font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: opts.size || 21
      })]
    })]
  });
}

function heading(text, level = 1) {
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    children: [new TextRun({ text, bold: true, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
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
    default: { document: { run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: "1E3A5F", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, keepNext: false, keepLines: false } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "2C5282", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: false, keepLines: false } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: "DC2626", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 2, keepNext: false, keepLines: false } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "valueBullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "◆", alignment: AlignmentType.LEFT,
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
        children: [new TextRun({ text: "Q3四大工作事项（领导目标提炼版·强化业务价值）", color: "94A3B8", size: 18, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
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
      para("以下 4 大工作事项，从领导给出的 Q3 目标方向中提炼、拔高、整合而成。你原来的 6 项具体工作，全部作为子项纳入对应大事项之下。"),
      para("核心变化：每个事项都明确回答\"做这件事到底能带来什么具体的、可衡量的业务价值\"——不是\"做了什么功能\"，而是\"因为做了这件事，业务会发生什么好的变化\"。"),

      // ===== 总览表 =====
      heading("二、四大工作事项总览（含核心价值点）"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [500, 2800, 4200, 1800],
        rows: [
          new TableRow({ cantSplit: true, children: [
            mkCell("序号", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 500, align: AlignmentType.CENTER }),
            mkCell("工作事项", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 2800, align: AlignmentType.CENTER }),
            mkCell("一句话说清业务价值", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 4200, align: AlignmentType.CENTER }),
            mkCell("纳入的原工作", { bold: true, color: "FFFFFF", fill: "1E3A5F", width: 1800, align: AlignmentType.CENTER }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("01", { width: 500, align: AlignmentType.CENTER, bold: true, fill: "FEF3C7" }),
            mkCell("数据架构与多维分析体系建设", { bold: true, width: 2800, fill: "FFFBEB" }),
            mkCell("让 SBG 从\"拿不到数据\"变成\"想怎么切就怎么切\"，决策周期从周级缩短到天级", { width: 4200 }),
            mkCell("新产品成功看板", { width: 1800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("02", { width: 500, align: AlignmentType.CENTER, bold: true, fill: "DBEAFE" }),
            mkCell("规则中心与业财数据治理", { bold: true, width: 2800, fill: "EFF6FF" }),
            mkCell("让改规则从\"找IT等2周\"变成\"业务自己配置当天生效\"，指标口径从\"三张嘴三个说法\"变成\"一个字典说了算\"", { width: 4200 }),
            mkCell("指标字典 + 需求落地", { width: 1800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("03", { width: 500, align: AlignmentType.CENTER, bold: true, fill: "D1FAE5" }),
            mkCell("项目全闭环与预算管理", { bold: true, width: 2800, fill: "ECFDF5" }),
            mkCell("让每个项目从\"投了多少钱不知道、赚不赚钱算不清\"变成\"投产-执行-核算一眼看透\"，预算从\"拍脑袋\"变成\"有数据支撑\"", { width: 4200 }),
            mkCell("投产规划表 + 现金流", { width: 1800 }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            mkCell("04", { width: 500, align: AlignmentType.CENTER, bold: true, fill: "FCE7F3" }),
            mkCell("AI赋能与业务价值深挖", { bold: true, width: 2800, fill: "FDF2F8" }),
            mkCell("让AI从\"实验室玩具\"变成\"业务提效工具\"——查指标不用翻文档、做预算不用拼表格、业务需求从\"等你做\"变成\"我知道你能做\"", { width: 4200 }),
            mkCell("销量预测模型", { width: 1800 }),
          ]}),
        ]
      }),

      // ===== 事项一 =====
      heading("三、事项一：数据架构与多维分析体系建设"),

      heading("3.1 现在的痛点", { level: 2 }),
      bulletItem("SBG 想做经营分析，但拿不到数据——要个数得找 IT，IT 要找数据开发，一周过去了还在排期", { ref: "bullets" }),
      bulletItem("数据只能按组织维度看，想切产品维度、市场维度、项目维度——切不出来，或者切出来对不上", { ref: "bullets" }),
      bulletItem("同样是\"SBG 收入\"，三张表三个数，没人知道信哪个", { ref: "bullets" }),
      bulletItem("其他 BG 看着也想要，但 SBG 都没跑通，没法推广", { ref: "bullets" }),

      heading("3.2 Q3 要做什么", { level: 2 }),
      bulletItem("以 SBG 为试点，深度调研运营分析层面到底需要什么数据、怎么用、多久要一次", { ref: "bullets" }),
      bulletItem("设计 SBG 经营/运营体系下的数据架构（从源系统到最终报表的全链路）", { ref: "bullets" }),
      bulletItem("建设\"产品、市场、项目\"三维财经管理能力——同一笔收入，换三个维度都能切得开、对得上", { ref: "bullets" }),
      bulletItem("形成可复制的方法论，为推广到其他 BG 打基础", { ref: "bullets" }),

      heading("3.3 做完之后的业务价值", { level: 2 }),

      para("◆ 决策效率提升：SBG 经营分析从\"等一周拿到静态报表\"变成\"当天自助获取多维度分析\"，决策周期缩短 70%+", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 数据可信度提升：三维数据打通后，\"SBG 收入\"不再有三个数，而是一个口径、一套数据，减少 80% 的数据核对时间", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 分析深度提升：从\"只看结果\"到\"能切产品/市场/项目找原因\"——发现哪个产品赚钱、哪个客户贡献大、哪个项目投入产出比低，直接指导经营动作", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 方法论沉淀：SBG 跑通后，复制到其他 BG 的成本降低 60%，不用从零开始搭", { ref: "valueBullets", bold: true, color: "DC2626" }),

      heading("3.4 包含的子项（你原有的工作）", { level: 2 }),
      para("新产品成功看板优化（权重 10%）", { bold: true }),
      bulletItem("作为\"产品维度\"分析的核心载体，纳入三维财经管理体系", { ref: "subBullets" }),
      bulletItem("看板 2.0 不只是功能升级，而是产品维度数据能力的展示窗口", { ref: "subBullets" }),

      heading("3.5 建议新增输出物", { level: 2 }),
      bulletItem("SBG 运营分析数据需求调研报告", { ref: "bullets" }),
      bulletItem("SBG 数据架构设计文档（含三维财经管理能力设计）", { ref: "bullets" }),

      // ===== 事项二 =====
      heading("四、事项二：规则中心与业财数据治理"),

      heading("4.1 现在的痛点", { level: 2 }),
      bulletItem("阿米巴的分摊规则、口径都写死在代码里——Q3 考核规则变了，得找 IT 改代码、排期、测试、上线，最快 2 周，慢则 1 个月", { ref: "bullets" }),
      bulletItem("同一个指标，阿米巴一个说法、财报一个说法、BG 经营又一个说法——开会的时候三张嘴三个数，没人说得清到底信哪个", { ref: "bullets" }),
      bulletItem("新人要理解一个指标，得翻历史文档、问老同事、自己对数据——搞清楚一个指标的口径要花 1-2 天", { ref: "bullets" }),
      bulletItem("数据质量问题反复出现，修了又出，出了又修，没有根因治理", { ref: "bullets" }),

      heading("4.2 Q3 要做什么", { level: 2 }),
      bulletItem("完成阿米巴管报指标字典 V1.0：每个指标的定义、公式、来源、口径差异，全部标准化", { ref: "bullets" }),
      bulletItem("梳理规则配置化需求：哪些规则可以从代码里抽出来、让业务自己配置（选 SBG 为试点梳理）", { ref: "bullets" }),
      bulletItem("完成与业务方、财务方的口径对齐——阿米巴、财报、BG 经营三套口径的差异全部明确写下来", { ref: "bullets" }),
      bulletItem("影响阿米巴数据质量的需求 100% 落地", { ref: "bullets" }),

      heading("4.3 做完之后的业务价值", { level: 2 }),

      para("◆ 规则变更效率：改规则从\"找 IT 排期 2 周\"变成\"业务自己配置当天生效\"，规则变更响应速度提升 90%+", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 口径一致性：指标字典上线后，\"一个指标三个说法\"的问题从常态变成个例，跨部门沟通效率提升 50%+，开会不再纠结\"数对不对\"", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 新人上手效率：新人查一个指标的口径，从\"翻文档+问人 1-2 天\"变成\"查字典 5 分钟搞定\"，新人上手周期缩短 70%", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 数据质量提升：数据质量问题从\"反复修\"变成\"根因治理\"，阿米巴数据质量问题发生率下降 60%+", { ref: "valueBullets", bold: true, color: "DC2626" }),

      heading("4.4 包含的子项（你原有的工作）", { level: 2 }),
      para("① 阿米巴指标字典整理（权重 40%）", { bold: true }),
      bulletItem("这是规则中心的基础——指标定义清楚了，规则才能配置化", { ref: "subBullets" }),
      para("② 阿米巴重点需求落地（权重 20%）", { bold: true }),
      bulletItem("数据质量需求 100% 落地，体验类需求交付 ≥ 3 个", { ref: "subBullets" }),

      heading("4.5 建议新增输出物", { level: 2 }),
      bulletItem("SBG 规则配置化需求梳理清单（为规则中心建设做输入）", { ref: "bullets" }),
      bulletItem("三套口径差异对照表（阿米巴 vs 财报 vs BG 经营）", { ref: "bullets" }),

      // ===== 事项三 =====
      heading("五、事项三：项目全闭环与预算管理"),

      heading("5.1 现在的痛点", { level: 2 }),
      bulletItem("一个项目投了多少钱、花在哪里、还剩多少预算——没人说得清，得从 3 个系统里拼数据", { ref: "bullets" }),
      bulletItem("项目做完了，到底赚不赚钱——算不清，因为收入在一个系统、成本在一个系统、工时在另一个系统", { ref: "bullets" }),
      bulletItem("预算靠拍脑袋——每年做预算的时候，各部门拍个数字报上去，没有数据支撑，老板也不知道合理不合理", { ref: "bullets" }),
      bulletItem("投产、预算执行、核算各干各的，数据不打通，项目全生命周期是断裂的", { ref: "bullets" }),

      heading("5.2 Q3 要做什么", { level: 2 }),
      bulletItem("围绕项目维度，打通\"投产 → 预算执行 → 核算\"全链路的数据", { ref: "bullets" }),
      bulletItem("配合领导推进预算相关工作（辅助支持角色）", { ref: "bullets" }),
      bulletItem("推动业财一体化——业务数据和财务数据不再是两张皮", { ref: "bullets" }),

      heading("5.3 做完之后的业务价值", { level: 2 }),

      para("◆ 项目投入产出可视化：从\"项目赚不赚钱算不清\"变成\"投产-执行-核算一眼看透\"，每个项目的 ROI 清晰可见，领导一眼就能看出该给哪些项目追加投入、该砍掉哪些项目", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 预算管理科学化：从\"拍脑袋做预算\"变成\"有历史数据、有项目维度、有执行跟踪\"的预算管理，预算编制效率提升 50%+，预算执行偏差率降低 30%", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 财务核算效率：项目核算从\"3 个系统拼数据花 3 天\"变成\"一键取数花 10 分钟\"，月结效率提升 80%", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 业财一体化落地：业务数据和财务数据打通后，业务人员不用再找财务要数，财务不用再帮业务拼数，双方各干各的核心工作", { ref: "valueBullets", bold: true, color: "DC2626" }),

      heading("5.4 包含的子项（你原有的工作）", { level: 2 }),
      para("① 产研项目投产规划表（权重 10%）", { bold: true }),
      bulletItem("作为项目闭环中的\"投产\"环节，纳入全链路管理", { ref: "subBullets" }),
      para("② 业财 2.0 - 现金流子项（权重 5%）", { bold: true }),
      bulletItem("作为业财一体化的重要组成部分，与预算管理联动", { ref: "subBullets" }),

      heading("5.5 建议新增输出物", { level: 2 }),
      bulletItem("项目全闭环数据链路说明（投产 → 预算执行 → 核算，每个环节的数据流）", { ref: "bullets" }),
      bulletItem("预算支持相关输出（配合领导预算工作的辅助材料）", { ref: "bullets" }),

      // ===== 事项四 =====
      heading("六、事项四：AI赋能与业务价值深挖"),

      heading("6.1 现在的痛点", { level: 2 }),
      bulletItem("AI 做了销量预测模型，但业务觉得和自己没关系——业务关心的是\"我做预算的时候能不能少拼点表格\"、\"我查指标口径的时候能不能少翻点文档\"", { ref: "bullets" }),
      bulletItem("业务需求都是被动接——业务不提，我们就不知道要做什么，很多有价值的场景没人挖", { ref: "bullets" }),
      bulletItem("日常问题支持耗了大量精力——业务方一个微信过来就开始排查，半天时间就没了", { ref: "bullets" }),

      heading("6.2 Q3 要做什么", { level: 2 }),
      bulletItem("AI 在预算上的应用场景：预算编制辅助、预算执行对比分析（配合领导）", { ref: "bullets" }),
      bulletItem("AI 在指标字典上的应用场景：智能口径查询、血缘追踪", { ref: "bullets" }),
      bulletItem("主动深挖业务需求：每月主动做 1-2 次业务侧调研", { ref: "bullets" }),
      bulletItem("日常需求落地与问题支持（基础运维保障）", { ref: "bullets" }),

      heading("6.3 做完之后的业务价值", { level: 2 }),

      para("◆ 预算编制提效：做预算从\"各部门拼 Excel 花 2 周\"变成\"AI 辅助生成初稿 1 天搞定\"，预算编制周期缩短 60%+，把财务和业务从繁琐的表格拼接中解放出来", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 指标查询提效：查一个指标的口径、公式、来源，从\"翻文档+问人 1-2 天\"变成\"问 AI 5 秒钟出答案\"，而且附带来源引用，答案可追溯", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 需求价值提升：从\"被动等业务提需求\"变成\"主动挖业务痛点\"，每个季度能多发现 3-5 个高价值需求，这些需求落地后带来的业务价值远超投入", { ref: "valueBullets", bold: true, color: "DC2626" }),
      para("◆ 日常支持降本：AI 辅助的指标口径查询 + 知识沉淀，让日常问题支持的响应时间缩短 50%+，不用每个问题都要人工排查", { ref: "valueBullets", bold: true, color: "DC2626" }),

      heading("6.4 包含的子项（你原有的工作）", { level: 2 }),
      para("销量预测模型 MVP 线上化（权重 15%）", { bold: true }),
      bulletItem("作为 AI 赋能的落地场景之一，与预算辅助联动", { ref: "subBullets" }),
      bulletItem("建议扩展：从销量预测延伸到预算辅助（领导更关注的方向）", { ref: "subBullets" }),

      heading("6.5 建议新增输出物", { level: 2 }),
      bulletItem("AI 辅助指标字典查询场景设计与原型", { ref: "bullets" }),
      bulletItem("AI 预算辅助场景设计（配合领导预算工作）", { ref: "bullets" }),
      bulletItem("业务需求深挖计划：月度调研清单", { ref: "bullets" }),

      // ===== 权重分配 =====
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

      // ===== 总结 =====
      heading("八、一句话总结"),
      para("以 SBG 为试点，围绕 \"数据架构 → 规则中心 → 项目闭环 → AI 赋能\" 这四条主线，每条线都瞄准一个具体的业务痛点、给出可衡量的价值预期——让领导一眼就能看出：做了这件事，业务会发生什么好的变化。"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\Q3四大工作事项（强化业务价值版）.docx", buffer);
  console.log("Done");
});
