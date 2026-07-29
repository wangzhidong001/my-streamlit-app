const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  ImageRun
} = require("docx");

// ===== 工具函数 =====
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const HEADER_FILL = "1F4E79";
const SUB_FILL = "D5E8F0";
const LIGHT_FILL = "F2F7FB";

function h(text, level) {
  return new Paragraph({
    heading: level,
    children: [new TextRun({ text, bold: true, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 360 },
    ...opts,
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 22, ...opts.run })]
  });
}

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function bullet(text, ref = "bullets", level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 80, line: 360 },
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 22 })]
  });
}

function makeCell(text, opts = {}) {
  const isHeader = opts.header;
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: isHeader
      ? { fill: HEADER_FILL, type: ShadingType.CLEAR }
      : opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 320 },
      children: [new TextRun({
        text,
        bold: isHeader,
        color: isHeader ? "FFFFFF" : "1F2937",
        font: { ascii: "Arial", eastAsia: "Microsoft YaHei" },
        size: isHeader ? 22 : 20
      })]
    })]
  });
}

function makeRow(cells, opts = {}) {
  return new TableRow({
    cantSplit: true,
    ...opts,
    children: cells
  });
}

// ===== 文档主体 =====
const children = [];

// 封面
children.push(new Paragraph({ spacing: { before: 3000 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 400 },
  children: [new TextRun({ text: "统一数据平台与指标中心建设方案", bold: true, size: 52, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, color: "1F4E79" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "支撑法定财报 · 阿米巴管报 · BG经营报表的一体化数据架构", size: 28, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, color: "4A5568" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 600 },
  children: [new TextRun({ text: "2026年7月", size: 24, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, color: "718096" })]
}));
children.push(pb());

// ===== 一、方案概述 =====
children.push(h("一、方案概述", HeadingLevel.HEADING_1));
children.push(p("本方案要解决的核心问题：公司同时跑三套报表（法定财报、阿米巴管报、BG经营报表），三套报表各取各的数、各算各的口径，导致数据重复采集、指标对不上、改个规则要等数周。"));
children.push(p("一句话方案：建一个统一的数据底座 + 一个指标中心，底层数据只取一次，上层各报表按自己的规则独立计算，互不干扰。"));

children.push(h("核心痛点", HeadingLevel.HEADING_2));
children.push(bullet("数据重复采集：三套报表各自从ERP/CRM/HR取数，相同数据被抽3次"));
children.push(bullet("口径对不齐：同一个\"收入\"，法定口径含税、阿米巴口径不含内部转移定价、BG口径又不一样"));
children.push(bullet("变更响应慢：阿米巴考核规则每季度调整，传统开发要等2-3周"));
children.push(bullet("问题找不到根：数据血缘断裂，报表出错了不知道是源系统的问题还是中间计算的问题"));

children.push(pb());

// ===== 二、整体架构 =====
children.push(h("二、整体架构（自下而上）", HeadingLevel.HEADING_1));
children.push(p("架构共分5层，最底层是数据源，最顶层是报表输出。注意：越往下越原始、越实时；越往上越加工、越综合。"));

// 架构图（倒金字塔：底层在下面，高层在上面）
children.push(h("2.1 五层架构图", HeadingLevel.HEADING_2));

// 用表格模拟架构图，倒序：最高层在上，最底层在下
children.push(new Paragraph({ spacing: { before: 200, after: 100 }, children: [
  new TextRun({ text: "（说明：以下从上到下 = 从高层到底层，最底层是原始数据）", italics: true, size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })
]}));

const archRows = [
  makeRow([makeCell("第5层：报表输出层（3套报表）", { header: true, width: 9000 })]),
  makeRow([makeCell("法定财报合并 | 阿米巴考核管报 | BG经营报表", { width: 9000, fill: "EBF5FF" })]),

  makeRow([makeCell("第4层：分析与应用层（AI + 指标服务）", { header: true, width: 9000 })]),
  makeRow([makeCell("指标口径助手 | 智能血缘分析 | 自然语言取数 | 经营洞察生成", { width: 9000, fill: "EBF5FF" })]),

  makeRow([makeCell("第3层：指标中心层（规则引擎）", { header: true, width: 9000 })]),
  makeRow([makeCell("原子指标池 | 映射中心 | 分摊引擎 | 版本管理 | 血缘追踪", { width: 9000, fill: "E6F4EA" })]),

  makeRow([makeCell("第2层：数据仓库层（统一建模）", { header: true, width: 9000 })]),
  makeRow([makeCell("ODS贴源层 | DWD明细层 | DWS汇总层 | ADS应用层（统一主数据编码）", { width: 9000, fill: "FFF7E6" })]),

  makeRow([makeCell("第1层：数据源层（最底层，原始数据）", { header: true, width: 9000 })]),
  makeRow([makeCell("ERP（订单/库存） | 总账系统 | CRM（客户/销售） | HR系统 | 项目管理系统 | 手工填报", { width: 9000, fill: "FCE4EC" })]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [9000],
  rows: archRows
}));

children.push(p(""));
children.push(h("2.2 各层职责一句话", HeadingLevel.HEADING_2));
children.push(bullet("第1层（数据源）：业务系统的原始数据，只取一次，不做加工"));
children.push(bullet("第2层（数仓）：清洗、建模、统一编码，形成大家共用的数据底"));
children.push(bullet("第3层（指标中心）：定义指标是什么、怎么算、归谁管，是整个方案的灵魂"));
children.push(bullet("第4层（分析应用）：给人用的工具——查指标、找问题、问数据"));
children.push(bullet("第5层（报表输出）：最终给不同人看的3套报表，各算各的"));

children.push(pb());

// ===== 三、指标中心（四层指标体系）=====
children.push(h("三、指标中心：四层指标体系", HeadingLevel.HEADING_1));
children.push(p("指标中心的核心思想：同一笔业务事实，在不同语境下是不同指标。我们不强行统一，而是把它们之间的关系说清楚。"));

children.push(h("3.1 四层指标架构图（L1在最下面，L4在最上面）", HeadingLevel.HEADING_2));

children.push(new Paragraph({ spacing: { before: 200, after: 100 }, children: [
  new TextRun({ text: "（从上到下 = L4 → L1，越往下越原始，越往上越分析）", italics: true, size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })
]}));

const indRows = [
  makeRow([
    makeCell("层级", { header: true, width: 1500 }),
    makeCell("定位", { header: true, width: 2000 }),
    makeCell("指标示例", { header: true, width: 3000 }),
    makeCell("核心特征", { header: true, width: 2500 }),
  ]),
  makeRow([
    makeCell("L4\n分析层\n（最上面）", { width: 1500, fill: "FCE4EC" }),
    makeCell("面向经营决策的自定义分析指标", { width: 2000, fill: "FCE4EC" }),
    makeCell("毛利率、人效、回款率、单客价值、库存周转天数", { width: 3000, fill: "FCE4EC" }),
    makeCell("由下层指标组合计算、口径灵活、可按场景调整", { width: 2500, fill: "FCE4EC" }),
  ]),
  makeRow([
    makeCell("L3\n报表层", { width: 1500, fill: "EBF5FF" }),
    makeCell("面向监管/管理/投资者的格式化输出", { width: 2000, fill: "EBF5FF" }),
    makeCell("资产负债表项目、利润表项目、管理报表项目", { width: 3000, fill: "EBF5FF" }),
    makeCell("按报表模板聚合、存在合并抵消、重分类调整", { width: 2500, fill: "EBF5FF" }),
  ]),
  makeRow([
    makeCell("L2\n财务核算层", { width: 1500, fill: "E6F4EA" }),
    makeCell("按会计准则确认的收入/成本/费用", { width: 2000, fill: "E6F4EA" }),
    makeCell("营业收入、营业成本、期间费用、应收账款", { width: 3000, fill: "E6F4EA" }),
    makeCell("口径严格、受准则约束、有明确的确认时点和科目归属", { width: 2500, fill: "E6F4EA" }),
  ]),
  makeRow([
    makeCell("L1\n业务运营层\n（最下面）", { width: 1500, fill: "FFF7E6" }),
    makeCell("记录业务发生的原始事实", { width: 2000, fill: "FFF7E6" }),
    makeCell("订单量、发货量、签约金额、客户拜访次数", { width: 3000, fill: "FFF7E6" }),
    makeCell("粒度最细、实时性高、未经财务规则加工", { width: 2500, fill: "FFF7E6" }),
  ]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [1500, 2000, 3000, 2500],
  rows: indRows
}));

children.push(p(""));
children.push(h("3.2 各层之间是什么关系？（3种模式）", HeadingLevel.HEADING_2));

children.push(p("模式1：一对一映射（Mapping）—— L1 → L2"));
children.push(bullet("适用场景：业务事实和财务确认之间有明确的转换规则"));
children.push(bullet("示例：订单金额（L1） → 【剔除退款 + 按履约进度确认 + 增值税价税分离】 → 营业收入（L2）"));
children.push(bullet("举个具体例子：1月签了100万订单（含13%增值税），客户付了30%首款，2月才发货。那么L1订单金额=100万，但L2营业收入在1月确认为0（未发货），2月确认=100÷1.13≈88.5万"));

children.push(p(""));
children.push(p("模式2：一对多拆解（Allocation）—— L1 → 多个L2"));
children.push(bullet("适用场景：一笔费用要按维度拆到不同科目"));
children.push(bullet("示例：市场部总费用100万（L1） → 【按项目工时比例分摊】 → A产品线营销费40万（L2）+ B产品线营销费35万（L2）+ 总部职能费用25万（L2）"));

children.push(p(""));
children.push(p("模式3：多对一聚合（Aggregation）—— 多个下层 → L4"));
children.push(bullet("适用场景：分析指标由多个下层指标组合计算"));
children.push(bullet("示例：单客贡献毛利（L4） = 营业收入（L2）÷ 活跃客户数（L1） - 营业成本（L2）÷ 活跃客户数（L1）"));

children.push(pb());

// ===== 四、核心模块 =====
children.push(h("四、核心模块设计", HeadingLevel.HEADING_1));

children.push(h("4.1 映射中心：解决\"按什么维度看数据\"", HeadingLevel.HEADING_2));
children.push(p("三套报表最大的差异是维度不同——法定按法人、阿米巴按虚拟巴、BG按事业群。映射中心就是维护多维关系："));
children.push(bullet("人 → 岗位 → 部门 → 法人公司 → 阿米巴单元 → BG事业群"));
children.push(bullet("每个映射关系都有版本（比如：张三2025年1-6月归东区销售巴，7月起归西区销售巴）"));
children.push(bullet("变更走审批，历史版本永久保留，支持\"回到当时口径\"复盘"));

children.push(p(""));
children.push(h("4.2 分摊引擎：解决\"费用怎么分\"", HeadingLevel.HEADING_2));

const allocRows = [
  makeRow([
    makeCell("分摊类型", { header: true, width: 2000 }),
    makeCell("适用场景", { header: true, width: 3500 }),
    makeCell("举个例子", { header: true, width: 3500 }),
  ]),
  makeRow([
    makeCell("固定比例分摊", { width: 2000 }),
    makeCell("总部职能费用按预先定好的比例分到各巴", { width: 3500 }),
    makeCell("HR部门费用按各巴人数比例分摊：A巴30人摊30%，B巴70人摊70%", { width: 3500 }),
  ]),
  makeRow([
    makeCell("动因分摊", { width: 2000, fill: LIGHT_FILL }),
    makeCell("IT费用按各巴实际使用量分摊", { width: 3500, fill: LIGHT_FILL }),
    makeCell("IT运维费：A巴用了1000次API调用摊40%，B巴用了1500次摊60%", { width: 3500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("阶梯分摊", { width: 2000 }),
    makeCell("收入达到一定规模后分摊比例递减", { width: 3500 }),
    makeCell("总部管理费：收入<1000万分摊5%，1000-5000万分摊3%，>5000万分摊1%", { width: 3500 }),
  ]),
  makeRow([
    makeCell("直接归属+剩余分摊", { width: 2000, fill: LIGHT_FILL }),
    makeCell("能直接归的直接归，剩下的再分摊", { width: 3500, fill: LIGHT_FILL }),
    makeCell("市场部费用：为A项目花的60万直接归A巴，剩余40万按收入比例分", { width: 3500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("工时/面积分摊", { width: 2000 }),
    makeCell("房租物业按面积、共享人力按工时", { width: 3500 }),
    makeCell("办公房租：A巴占300㎡摊30%，B巴占700㎡摊70%", { width: 3500 }),
  ]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [2000, 3500, 3500],
  rows: allocRows
}));

children.push(p(""));
children.push(h("4.3 数据血缘与一致性校验", HeadingLevel.HEADING_2));
children.push(p("每个指标都能追根溯源——点一下\"毛利率\"，就能看到它=L2营业收入-L2营业成本，再点营业收入就能看到它来自L1订单金额经过了哪些转换，最终追到ERP的哪张表哪个字段。"));

children.push(p(""));
children.push(bullet("纵向平衡：集团营业收入 = 各法人营业收入之和（合并抵消前）"));
children.push(bullet("横向平衡：营业收入 - 营业成本 = 毛利，每月校验"));
children.push(bullet("时序平衡：本月期末应收 = 上月期末 + 本月新增 - 本月回款"));
children.push(bullet("外部一致性：指标中心数据 vs ERP总账数据，差异控制在0.1%以内"));

children.push(pb());

// ===== 五、三套报表差异化支撑 =====
children.push(h("五、三套报表如何差异化支撑", HeadingLevel.HEADING_1));
children.push(p("统一平台不是\"一套报表打天下\"，而是在共享数据底座的基础上，让三套报表各算各的。"));

const reportRows = [
  makeRow([
    makeCell("对比维度", { header: true, width: 1500 }),
    makeCell("法定财报合并", { header: true, width: 2500 }),
    makeCell("阿米巴考核管报", { header: true, width: 2500 }),
    makeCell("BG经营报表", { header: true, width: 2500 }),
  ]),
  makeRow([
    makeCell("给谁看", { width: 1500 }),
    makeCell("CFO、审计、投资者", { width: 2500 }),
    makeCell("各巴负责人、HRBP、财务BP", { width: 2500 }),
    makeCell("BG总裁、运营总监", { width: 2500 }),
  ]),
  makeRow([
    makeCell("按什么组织", { width: 1500, fill: LIGHT_FILL }),
    makeCell("法人公司/子公司", { width: 2500, fill: LIGHT_FILL }),
    makeCell("虚拟阿米巴单元（可跨法人）", { width: 2500, fill: LIGHT_FILL }),
    makeCell("事业群（BG）/事业部", { width: 2500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("收入怎么算", { width: 1500 }),
    makeCell("会计准则（权责发生制）", { width: 2500 }),
    makeCell("内部交易价 + 外部收入拆分", { width: 2500 }),
    makeCell("BG可自定义确认节点", { width: 2500 }),
  ]),
  makeRow([
    makeCell("多久更新", { width: 1500, fill: LIGHT_FILL }),
    makeCell("月/季/年", { width: 2500, fill: LIGHT_FILL }),
    makeCell("周/月", { width: 2500, fill: LIGHT_FILL }),
    makeCell("月/季", { width: 2500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("规则改得多吗", { width: 1500 }),
    makeCell("很少（受准则约束）", { width: 2500 }),
    makeCell("频繁（每季度可能调考核规则）", { width: 2500 }),
    makeCell("中等（随组织架构调）", { width: 2500 }),
  ]),
  makeRow([
    makeCell("内部交易", { width: 1500, fill: LIGHT_FILL }),
    makeCell("合并抵消（要消掉）", { width: 2500, fill: LIGHT_FILL }),
    makeCell("模拟内部定价，各巴都要算损益", { width: 2500, fill: LIGHT_FILL }),
    makeCell("视BG管理需要处理", { width: 2500, fill: LIGHT_FILL }),
  ]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [1500, 2500, 2500, 2500],
  rows: reportRows
}));

children.push(pb());

// ===== 六、AI应用场景 =====
children.push(h("六、AI能做什么（6个场景，按优先级）", HeadingLevel.HEADING_1));

const aiRows = [
  makeRow([
    makeCell("场景", { header: true, width: 2200 }),
    makeCell("一句话说明", { header: true, width: 3800 }),
    makeCell("举个例子", { header: true, width: 3000 }),
  ]),
  makeRow([
    makeCell("P0 智能指标口径助手", { width: 2200 }),
    makeCell("用RAG（检索增强生成），输入指标名就能查到它的定义、公式、来源、责任部门", { width: 3800 }),
    makeCell("问：\"阿米巴的营业收入和法定的有什么不一样？\" → 自动对比两套口径的差异", { width: 3000 }),
  ]),
  makeRow([
    makeCell("P0 智能数据质量监控", { width: 2200, fill: LIGHT_FILL }),
    makeCell("自动检测数据异常（突增/突降/缺失），不用等人发现了才去查", { width: 3800, fill: LIGHT_FILL }),
    makeCell("某巴的交通费突然比上月涨了300% → 自动告警并定位：是源系统录错了一笔", { width: 3000, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("P1 智能血缘与根因分析", { width: 2200 }),
    makeCell("指标异常时，自动沿着血缘链往下追，找到最底层的原因", { width: 3800 }),
    makeCell("毛利率下降了 → 自动追溯：是哪个产品线？哪个客户？是收入降了还是成本涨了？", { width: 3000 }),
  ]),
  makeRow([
    makeCell("P1 规则变更影响评估", { width: 2200, fill: LIGHT_FILL }),
    makeCell("改一个分摊规则前，AI先模拟算一遍：哪些指标会变、变多少、影响哪些报表", { width: 3800, fill: LIGHT_FILL }),
    makeCell("把IT费用从\"按人数分摊\"改成\"按使用量分摊\" → 模拟：A巴费用增15万，B巴减15万", { width: 3000, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("P1 自然语言取数", { width: 2200 }),
    makeCell("用中文问问题，AI自动生成SQL查数据、出报表（高敏感数据不开）", { width: 3800 }),
    makeCell("问：\"上个月EBG的前5大客户收入排行\" → 自动出一张表", { width: 3000 }),
  ]),
  makeRow([
    makeCell("P2 经营分析自动生成", { width: 2200, fill: LIGHT_FILL }),
    makeCell("每月自动写经营分析初稿，人只需要改和补", { width: 3800, fill: LIGHT_FILL }),
    makeCell("自动生成：\"本月收入同比增15%，主要来自SBG产品线增长23%...\"", { width: 3000, fill: LIGHT_FILL }),
  ]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [2200, 3800, 3000],
  rows: aiRows
}));

children.push(p(""));
children.push(p("⚠️ 风控提醒：AI是辅助，不是替代。自然语言取数结果必须和关键指标交叉验证，高敏感数据不开放AI查询。", { run: { bold: true, color: "C53030" } }));

children.push(pb());

// ===== 七、实施路径 =====
children.push(h("七、分三步走（12个月）", HeadingLevel.HEADING_1));

children.push(h("第一期（0-4个月）：数据底座 + 法定财报先跑通", HeadingLevel.HEADING_2));
children.push(bullet("把ERP、总账、HR等核心系统的数据接进来"));
children.push(bullet("搭好数仓（ODS→DWD→DWS→ADS），统一主数据编码"));
children.push(bullet("把法定财报的逻辑迁到平台上，让第一套报表跑起来"));
children.push(bullet("同步建指标中台的基础框架（原子指标池、规则引擎雏形）"));
children.push(bullet("AI落地：先上智能数据质量监控试点"));

children.push(p(""));
children.push(h("第二期（4-8个月）：阿米巴管报迁移 + 映射中心", HeadingLevel.HEADING_2));
children.push(bullet("建好映射中心，把\"人-岗-巴-法人-BG\"的关系都维护起来"));
children.push(bullet("把阿米巴管报的核心逻辑（收入拆分、成本分摊）迁到平台"));
children.push(bullet("上线口径版本管理和模拟测算功能"));
children.push(bullet("AI落地：上线智能指标口径助手、规则变更影响评估"));

children.push(p(""));
children.push(h("第三期（8-12个月）：BG经营报表 + 全面智能化", HeadingLevel.HEADING_2));
children.push(bullet("把BG经营报表也迁过来，支持BG自己配置口径"));
children.push(bullet("建统一的报表门户和数据API服务层"));
children.push(bullet("试点开放自然语言取数（限部分用户）"));
children.push(bullet("全面推广AI辅助的数据质量监控和异常根因分析"));

children.push(pb());

// ===== 八、风险与关键成功因素 =====
children.push(h("八、风险与成功关键", HeadingLevel.HEADING_1));

const riskRows = [
  makeRow([
    makeCell("风险", { header: true, width: 2500 }),
    makeCell("影响", { header: true, width: 3000 }),
    makeCell("怎么应对", { header: true, width: 3500 }),
  ]),
  makeRow([
    makeCell("源系统数据质量差", { width: 2500 }),
    makeCell("报表结果不可信，平台没人用", { width: 3000 }),
    makeCell("接入时就做质量评估，给数据打\"可信度分\"，低分的不入库或加醒目标识", { width: 3500 }),
  ]),
  makeRow([
    makeCell("组织映射太复杂，天天变", { width: 2500, fill: LIGHT_FILL }),
    makeCell("映射维护成本高，口径追不上", { width: 3000, fill: LIGHT_FILL }),
    makeCell("映射中心做成配置化+版本化，变更走审批，历史版本永久留", { width: 3500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("三套报表的人意见不统一", { width: 2500 }),
    makeCell("平台被多方拉扯，进度拖慢", { width: 3000 }),
    makeCell("成立\"数据治理委员会\"，由CFO/COO牵头，定好优先级和边界", { width: 3500 }),
  ]),
  makeRow([
    makeCell("AI给错结果", { width: 2500, fill: LIGHT_FILL }),
    makeCell("决策失误", { width: 3000, fill: LIGHT_FILL }),
    makeCell("AI输出必须经过\"校验层\"（和关键指标交叉验证），敏感数据禁止AI直接访问", { width: 3500, fill: LIGHT_FILL }),
  ]),
  makeRow([
    makeCell("IT团队能力不够", { width: 2500 }),
    makeCell("平台做一半用不起来", { width: 3000 }),
    makeCell("第一期找厂商驻场+内部培养，同步练自己的ETL和AI工程师", { width: 3500 }),
  ]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [2500, 3000, 3500],
  rows: riskRows
}));

children.push(p(""));
children.push(h("5个成功关键", HeadingLevel.HEADING_2));
children.push(bullet("高层背书：这是\"一把手工程\"，需要CFO/COO持续关注和投入"));
children.push(bullet("先跑通一套：别贪多，先让法定财报稳定运行，证明了能力再迁阿米巴和BG"));
children.push(bullet("指标中台是灵魂：投入资源打磨规则引擎，这才是平台的核心价值"));
children.push(bullet("业务人员要参与：阿米巴口径是财务BP定的，不是IT定的，规则引擎要让业务人能看懂能配置"));
children.push(bullet("AI适度引入：AI是提效的，不能替代人做判断，高敏感场景留人工审核"));

// ===== 生成文档 =====
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 22 }
      }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: "1F4E79", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, keepNext: false, keepLines: false } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "2C5282", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1, keepNext: false, keepLines: false } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
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
        children: [new TextRun({ text: "统一数据平台与指标中心建设方案", size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "第 ", size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
          new TextRun({ text: " 页", size: 18, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
        ]
      })] })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\统一数据平台与指标中心建设方案.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Saved to: " + outPath);
});
