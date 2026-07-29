const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const HF = "1F4E79";
const LF = "F2F7FB";
const SF = "EBF5FF"; // stage fill

function h(text, level) {
  return new Paragraph({
    heading: level, spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 80, line: 340 }, ...opts,
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 21, ...opts.run })]
  });
}
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 50, line: 340 },
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 })]
  });
}
function subBullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 1 },
    spacing: { after: 40, line: 320 },
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 20, color: "4A5568" })]
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.header
      ? { fill: HF, type: ShadingType.CLEAR }
      : opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.colSpan,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT, spacing: { line: 300 },
      children: [new TextRun({
        text, bold: opts.header,
        color: opts.header ? "FFFFFF" : "1F2937",
        font: { ascii: "Arial", eastAsia: "Microsoft YaHei" },
        size: opts.header ? 21 : 19
      })]
    })]
  });
}
function row(cells, opts = {}) {
  return new TableRow({ cantSplit: true, ...opts, children: cells });
}

// ============ 文档内容 ============
const children = [];

// 封面
children.push(new Paragraph({ spacing: { before: 3200 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 300 },
  children: [new TextRun({ text: "AI在统一数据平台建设全生命周期的应用", bold: true, size: 44, color: HF, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 200 },
  children: [new TextRun({ text: "建设前 · 建设中 · 建设后 三阶段全景图", size: 26, color: "4A5568", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 500 },
  children: [new TextRun({ text: "2026年7月", size: 22, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(pb());

// 一、总览
children.push(h("一、AI应用全景图", HeadingLevel.HEADING_1));
children.push(p("AI不只是平台建成后的\"锦上添花\"，而是可以贯穿平台建设的全生命周期——从规划设计、开发实施到运维运营，每个阶段都能产生实质性价值。"));

// 三阶段汇总表
children.push(h("三阶段AI应用总览", HeadingLevel.HEADING_2));

const summaryRows = [
  row([
    cell("阶段", { header: true, width: 1500 }),
    cell("核心目标", { header: true, width: 2500 }),
    cell("AI应用场景", { header: true, width: 5000 }),
  ]),
  row([
    cell("建设前\n（规划设计）", { width: 1500, fill: "FFF7E6" }),
    cell("摸清家底、明确需求、做好设计", { width: 2500, fill: "FFF7E6" }),
    cell("指标盘点与去重 · 口径差异自动识别 · 现有数据血缘分析 · 需求文档自动生成 · 主数据匹配辅助", { width: 5000, fill: "FFF7E6" }),
  ]),
  row([
    cell("建设中\n（开发实施）", { width: 1500, fill: "E6F4EA" }),
    cell("提高效率、保障质量、加速交付", { width: 2500, fill: "E6F4EA" }),
    cell("ETL/SQL代码生成 · 数据质量自动检测 · 测试用例生成 · 文档自动生成 · 主数据标准化辅助", { width: 5000, fill: "E6F4EA" }),
  ]),
  row([
    cell("建设后\n（运维运营）", { width: 1500, fill: "EBF5FF" }),
    cell("降低门槛、提效决策、持续优化", { width: 2500, fill: "EBF5FF" }),
    cell("智能指标口径助手 · 智能数据质量监控 · 智能血缘与根因分析 · 规则变更影响评估 · 自然语言取数 · 经营分析自动生成", { width: 5000, fill: "EBF5FF" }),
  ]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1500, 2500, 5000], rows: summaryRows }));
children.push(pb());

// ============ 二、建设前 ============
children.push(h("二、建设前（规划设计阶段）", HeadingLevel.HEADING_1));
children.push(p("这一阶段的核心是\"摸清家底\"——搞清楚现在有什么数据、有多少指标、哪些重复了、口径差在哪。AI可以大幅缩短调研周期，提高需求梳理的完整性。"));

children.push(h("场景1：现有指标盘点与去重识别", HeadingLevel.HEADING_2));
children.push(bullet("做什么：自动扫描三套报表（法定/阿米巴/BG）的所有指标，识别重复定义和口径冲突"));
children.push(bullet("怎么做："));
children.push(subBullet("用大模型读取现有的报表模板、指标字典、ETL脚本，提取所有指标名称和定义"));
children.push(subBullet("通过语义相似度计算，自动识别\"同名不同义\"和\"同义不同名\"的指标"));
children.push(subBullet("生成指标盘点报告：总共有多少个指标、哪些重复、哪些有冲突、哪些口径不明确"));
children.push(bullet("示例：扫描后发现\"营业收入\"在法定报表、阿米巴报表、BG报表中各有一套定义，AI自动标注三者的差异点"));

children.push(p(""));
children.push(h("场景2：口径差异自动识别与对比", HeadingLevel.HEADING_2));
children.push(bullet("做什么：对同一指标在不同报表中的口径差异进行结构化对比"));
children.push(bullet("怎么做："));
children.push(subBullet("读取指标定义文档、财务制度、考核规则，提取口径要素（确认时点、计算范围、调整项、排除项）"));
children.push(subBullet("按要素维度自动对比，输出差异矩阵表"));
children.push(bullet("示例：自动对比\"收入\"在三套报表中的差异："));

const diffRows = [
  row([
    cell("口径要素", { header: true, width: 1800 }),
    cell("法定财报", { header: true, width: 2400 }),
    cell("阿米巴管报", { header: true, width: 2400 }),
    cell("BG经营报表", { header: true, width: 2400 }),
  ]),
  row([cell("确认时点", { width: 1800 }), cell("按权责发生制（发货后）", { width: 2400 }), cell("按内部交易确认时点", { width: 2400 }), cell("BG自定义确认节点", { width: 2400 })]),
  row([cell("是否含税", { width: 1800, fill: LF }), cell("不含增值税", { width: 2400, fill: LF }), cell("不含内部转移定价", { width: 2400, fill: LF }), cell("根据BG需要", { width: 2400, fill: LF })]),
  row([cell("内部交易", { width: 1800 }), cell("合并抵消", { width: 2400 }), cell("模拟内部定价计入", { width: 2400 }), cell("视需要处理", { width: 2400 })]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1800, 2400, 2400, 2400], rows: diffRows }));

children.push(p(""));
children.push(h("场景3：现有数据链路与血缘分析", HeadingLevel.HEADING_2));
children.push(bullet("做什么：分析现有三套报表的数据链路，搞清楚每个指标从哪来、经过了哪些加工"));
children.push(bullet("怎么做："));
children.push(subBullet("读取现有的ETL脚本、存储过程、报表SQL，解析表之间的依赖关系"));
children.push(subBullet("自动绘制数据血缘图，标注关键链路和薄弱环节"));
children.push(subBullet("识别\"断点\"——哪些指标没有血缘、哪些计算逻辑是黑盒"));
children.push(bullet("价值：避免重复建设，识别应该优先迁移的核心链路"));

children.push(p(""));
children.push(h("场景4：需求梳理与方案文档生成", HeadingLevel.HEADING_2));
children.push(bullet("做什么：辅助整理业务需求，自动生成需求文档和技术方案初稿"));
children.push(bullet("怎么做："));
children.push(subBullet("业务访谈录音→转文字→AI自动提取需求要点和指标定义"));
children.push(subBullet("基于行业最佳实践，自动生成数据模型设计建议、数仓分层方案初稿"));
children.push(subBullet("自动检查需求文档的完整性：有没有遗漏的指标、有没有定义模糊的口径"));

children.push(p(""));
children.push(h("场景5：主数据匹配辅助", HeadingLevel.HEADING_2));
children.push(bullet("做什么：在统一主数据编码前，先把各系统中同一实体的不同编码匹配起来"));
children.push(bullet("怎么做："));
children.push(subBullet("读取各业务系统的组织、人员、客户、产品编码表"));
children.push(subBullet("通过名称相似度、属性匹配（如客户行业、地区）自动识别同一实体"));
children.push(subBullet("输出匹配建议清单，人工审核确认"));
children.push(bullet("示例：ERP系统里叫\"北京XX科技有限公司\"，CRM里叫\"北京XX科技\"，AI自动识别为同一家客户"));

children.push(pb());

// ============ 三、建设中 ============
children.push(h("三、建设中（开发实施阶段）", HeadingLevel.HEADING_1));
children.push(p("这一阶段的核心是\"提效保质\"——让开发更快、质量更高、文档更全。AI可以作为开发工程师的\"助手\"，大幅提升产出效率。"));

children.push(h("场景6：ETL/SQL代码自动生成", HeadingLevel.HEADING_2));
children.push(bullet("做什么：用自然语言描述需求，AI自动生成ETL脚本和SQL代码"));
children.push(bullet("怎么做："));
children.push(subBullet("\"把ERP的订单表和CRM的客户表关联，按BG维度聚合收入\" → 自动生成SQL"));
children.push(subBullet("\"做一个阿米巴费用分摊的存储过程，按人头比例分摊\" → 自动生成完整的存储过程代码"));
children.push(subBullet("基于已有的表结构和字段注释，AI自动理解数据语义，生成的代码准确率更高"));
children.push(bullet("价值：开发效率提升30%-50%，减少重复编码工作"));
children.push(bullet("⚠️ 注意：AI生成的代码必须人工审核，尤其是涉及财务口径的逻辑"));

children.push(p(""));
children.push(h("场景7：数据接入阶段的质量自动检测", HeadingLevel.HEADING_2));
children.push(bullet("做什么：数据从源系统接入ODS层时，自动检测质量问题"));
children.push(bullet("检测内容："));
children.push(subBullet("空值检测：哪些字段空值率异常高"));
children.push(subBullet("异常值检测：金额为负、数量超大、日期格式不对"));
children.push(subBullet("一致性检测：同一订单在ERP和CRM中的金额对不上"));
children.push(subBullet("完整性检测：应该每天都有数据，某天突然断了"));
children.push(bullet("示例：接入总账数据时，AI自动发现\"2025年3月有120条凭证的部门编码为空\"，自动生成问题清单"));

children.push(p(""));
children.push(h("场景8：测试用例自动生成", HeadingLevel.HEADING_2));
children.push(bullet("做什么：为数据加工逻辑自动生成测试用例"));
children.push(bullet("怎么做："));
children.push(subBullet("分析ETL逻辑的输入输出，自动设计边界测试用例（空值、极值、异常值）"));
children.push(subBullet("生成测试数据和预期结果，开发人员只需要跑测试验证"));
children.push(subBullet("自动回归测试：代码改了以后，AI自动把之前的测试用例再跑一遍"));
children.push(bullet("价值：测试覆盖率提升，避免改出bug"));

children.push(p(""));
children.push(h("场景9：数据字典与技术文档自动生成", HeadingLevel.HEADING_2));
children.push(bullet("做什么：自动生成数据字典、表结构说明、接口文档"));
children.push(bullet("怎么做："));
children.push(subBullet("读取数据库的表结构、字段注释、SQL逻辑"));
children.push(subBullet("AI自动补全字段的业务含义说明（很多表的字段注释不全）"));
children.push(subBullet("生成标准化的数据字典文档：表名、字段名、类型、业务含义、示例值、来源表"));
children.push(bullet("价值：解决\"文档跟不上代码\"的老问题，文档和代码同步更新"));

children.push(p(""));
children.push(h("场景10：主数据标准化辅助", HeadingLevel.HEADING_2));
children.push(bullet("做什么：在建设过程中辅助主数据的清洗和标准化"));
children.push(bullet("怎么做："));
children.push(subBullet("人员主数据：自动识别重复的员工记录（同名同岗位、离职后又入职）"));
children.push(subBullet("客户主数据：自动匹配同一客户的不同名称变体，生成统一编码建议"));
children.push(subBullet("产品主数据：根据产品名称和属性，自动归类到产品线和细分市场"));
children.push(bullet("价值：主数据清洗工作量巨大，AI可以处理80%的明确匹配，剩下20%疑难的交给人"));

children.push(pb());

// ============ 四、建设后 ============
children.push(h("四、建设后（运维运营阶段）", HeadingLevel.HEADING_1));
children.push(p("这一阶段的核心是\"降低门槛、提效决策\"——让业务人员不用找IT就能查数据，让数据问题自动暴露，让经营分析自动产出。"));

children.push(h("场景11：智能指标口径助手（RAG）", HeadingLevel.HEADING_2));
children.push(bullet("做什么：任何时候想知道某个指标的定义，直接问AI"));
children.push(bullet("怎么做："));
children.push(subBullet("把指标中心的所有元数据（定义、公式、来源、Owner）喂给RAG系统"));
children.push(subBullet("用户用自然语言提问，AI检索后给出准确回答"));
children.push(bullet("示例："));
children.push(subBullet("问：\"阿米巴的营业收入和法定的有什么不一样？\""));
children.push(subBullet("答：\"两者在三个方面不同：1.确认时点... 2.内部交易处理... 3.是否含税... 具体差异见下表...\""));
children.push(bullet("价值：减少\"这个数怎么算的\"这类问题，让财务BP从解释口径中解放出来"));

children.push(p(""));
children.push(h("场景12：智能数据质量监控", HeadingLevel.HEADING_2));
children.push(bullet("做什么：24小时自动监控数据质量，异常自动告警"));
children.push(bullet("监控内容："));
children.push(subBullet("波动异常：交通费比上月涨300%、某客户收入突然为0"));
children.push(subBullet("趋势偏离：实际值和历史趋势线偏差超过2个标准差"));
children.push(subBullet("跨指标一致性：收入涨了但毛利没涨、订单量涨了但库存没降"));
children.push(subBullet("时效性：某张表该T+1更新但没更"));
children.push(bullet("示例：AI自动发现\"研发部的低值易耗费本月突然是上月的5倍\"，自动推送给财务BP和IT"));

children.push(p(""));
children.push(h("场景13：智能血缘与异常根因分析", HeadingLevel.HEADING_2));
children.push(bullet("做什么：指标异常时，自动沿着血缘链找到根因"));
children.push(bullet("怎么做："));
children.push(subBullet("维护完整的指标血缘图（L1→L2→L3→L4）"));
children.push(subBullet("某个L4指标异常时，AI自动下钻到L3、L2、L1，逐层排查"));
children.push(bullet("示例："));
children.push(subBullet("告警：\"EBG的毛利率下降了5个百分点\""));
children.push(subBullet("AI自动追溯：毛利率=收入-成本 → 收入正常 → 成本上涨 → 成本中A产品的采购成本涨了 → 找到原因：A产品的供应商涨价"));
children.push(subBullet("输出：\"毛利率下降主要由A产品采购成本上涨导致，影响约3.5个百分点\""));

children.push(p(""));
children.push(h("场景14：规则变更影响评估", HeadingLevel.HEADING_2));
children.push(bullet("做什么：改一个分摊规则或考核口径前，先让AI模拟算一遍影响"));
children.push(bullet("怎么做："));
children.push(subBullet("输入变更规则：\"把IT费用从按人数分摊改成按系统使用量分摊\""));
children.push(subBullet("AI用历史数据跑两套规则，输出每个阿米巴单元的费用变化"));
children.push(subBullet("自动识别受影响最大的TOP10单元，生成对比报告"));
children.push(bullet("示例："));
children.push(subBullet("模拟结果：\"A巴费用增加15万（+12%），B巴费用减少15万（-8%），C巴基本不变。建议：A巴负责人需要提前沟通\""));

children.push(p(""));
children.push(h("场景15：自然语言取数（Text-to-SQL）", HeadingLevel.HEADING_2));
children.push(bullet("做什么：业务人员用中文提问，AI自动生成SQL、查出数据、出图表"));
children.push(bullet("示例："));
children.push(subBullet("问：\"上个月EBG的前5大客户收入排行\""));
children.push(subBullet("AI自动：生成SQL → 执行查询 → 输出一张表 + 柱状图"));
children.push(subBullet("问：\"今年各BG的人均开票毛利，和去年对比\""));
children.push(subBullet("AI自动：计算两年的数据 → 输出对比表 + 同比变化"));
children.push(bullet("⚠️ 风控："));
children.push(subBullet("高敏感数据（如薪酬）不开放自然语言查询"));
children.push(subBullet("AI生成的SQL必须经过\"结果校验层\"，和预置关键指标交叉验证"));
children.push(subBullet("只返回聚合数据，不返回明细行（避免数据泄露）"));

children.push(p(""));
children.push(h("场景16：经营分析报告自动生成", HeadingLevel.HEADING_2));
children.push(bullet("做什么：每月/每季度自动生成经营分析报告初稿，人只需要改和补"));
children.push(bullet("怎么做："));
children.push(subBullet("读取当月所有关键指标数据，计算同比、环比、预算达成率"));
children.push(subBullet("自动识别亮点和问题：哪些指标超预期、哪些拖后腿"));
children.push(subBullet("生成结构化的分析报告：核心结论 → 各BG表现 → 重点问题 → 建议动作"));
children.push(bullet("示例输出："));
children.push(subBullet("\"本月整体收入同比增长15%，主要由SBG产品线贡献（+23%）。EBG毛利率下降5个百分点，主要原因是A产品采购成本上涨（见根因分析）。建议：采购部评估A产品替代供应商。\""));
children.push(bullet("价值：写报告的时间从3天缩短到2小时，财务BP可以把精力放在分析而不是写报告上"));

children.push(pb());

// ============ 五、优先级矩阵 ============
children.push(h("五、实施优先级矩阵", HeadingLevel.HEADING_1));

const prioRows = [
  row([
    cell("场景", { header: true, width: 2800 }),
    cell("阶段", { header: true, width: 1200 }),
    cell("业务价值", { header: true, width: 1000 }),
    cell("技术成熟度", { header: true, width: 1200 }),
    cell("实施难度", { header: true, width: 1000 }),
    cell("推荐优先级", { header: true, width: 1800 }),
  ]),
  row([cell("指标盘点与去重识别", { width: 2800 }), cell("建设前", { width: 1200 }), cell("高", { width: 1000, align: "center" }), cell("高", { width: 1200, align: "center" }), cell("低", { width: 1000, align: "center" }), cell("P0", { width: 1800, align: "center", fill: "FCE4EC" })]),
  row([cell("口径差异自动识别", { width: 2800, fill: LF }), cell("建设前", { width: 1200, fill: LF }), cell("高", { width: 1000, align: "center", fill: LF }), cell("高", { width: 1200, align: "center", fill: LF }), cell("低", { width: 1000, align: "center", fill: LF }), cell("P0", { width: 1800, align: "center", fill: "FCE4EC" })]),
  row([cell("ETL/SQL代码生成", { width: 2800 }), cell("建设中", { width: 1200 }), cell("高", { width: 1000, align: "center" }), cell("高", { width: 1200, align: "center" }), cell("低", { width: 1000, align: "center" }), cell("P0", { width: 1800, align: "center", fill: "FCE4EC" })]),
  row([cell("数据质量自动检测", { width: 2800, fill: LF }), cell("建设中+后", { width: 1200, fill: LF }), cell("高", { width: 1000, align: "center", fill: LF }), cell("高", { width: 1200, align: "center", fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("P0", { width: 1800, align: "center", fill: "FCE4EC" })]),
  row([cell("智能指标口径助手", { width: 2800 }), cell("建设后", { width: 1200 }), cell("高", { width: 1000, align: "center" }), cell("高(RAG成熟)", { width: 1200, align: "center" }), cell("低", { width: 1000, align: "center" }), cell("P0", { width: 1800, align: "center", fill: "FCE4EC" })]),
  row([cell("智能血缘与根因分析", { width: 2800, fill: LF }), cell("建设后", { width: 1200, fill: LF }), cell("高", { width: 1000, align: "center", fill: LF }), cell("中", { width: 1200, align: "center", fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("P1", { width: 1800, align: "center", fill: "FFF7E6" })]),
  row([cell("规则变更影响评估", { width: 2800 }), cell("建设后", { width: 1200 }), cell("中", { width: 1000, align: "center" }), cell("中", { width: 1200, align: "center" }), cell("中", { width: 1000, align: "center" }), cell("P1", { width: 1800, align: "center", fill: "FFF7E6" })]),
  row([cell("自然语言取数", { width: 2800, fill: LF }), cell("建设后", { width: 1200, fill: LF }), cell("高", { width: 1000, align: "center", fill: LF }), cell("中", { width: 1200, align: "center", fill: LF }), cell("高(风控严)", { width: 1000, align: "center", fill: LF }), cell("P1", { width: 1800, align: "center", fill: "FFF7E6" })]),
  row([cell("测试用例自动生成", { width: 2800 }), cell("建设中", { width: 1200 }), cell("中", { width: 1000, align: "center" }), cell("高", { width: 1200, align: "center" }), cell("中", { width: 1000, align: "center" }), cell("P1", { width: 1800, align: "center", fill: "FFF7E6" })]),
  row([cell("主数据匹配/标准化", { width: 2800, fill: LF }), cell("建设前+中", { width: 1200, fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("中", { width: 1200, align: "center", fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("P1", { width: 1800, align: "center", fill: "FFF7E6" })]),
  row([cell("经营分析报告自动生成", { width: 2800 }), cell("建设后", { width: 1200 }), cell("中", { width: 1000, align: "center" }), cell("低", { width: 1200, align: "center" }), cell("高", { width: 1000, align: "center" }), cell("P2", { width: 1800, align: "center", fill: "E6F4EA" })]),
  row([cell("需求文档自动生成", { width: 2800, fill: LF }), cell("建设前", { width: 1200, fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("中", { width: 1200, align: "center", fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("P2", { width: 1800, align: "center", fill: "E6F4EA" })]),
  row([cell("数据血缘分析(建设期)", { width: 2800 }), cell("建设前", { width: 1200 }), cell("中", { width: 1000, align: "center" }), cell("中", { width: 1200, align: "center" }), cell("中", { width: 1000, align: "center" }), cell("P2", { width: 1800, align: "center", fill: "E6F4EA" })]),
  row([cell("文档自动生成", { width: 2800, fill: LF }), cell("建设中", { width: 1200, fill: LF }), cell("中", { width: 1000, align: "center", fill: LF }), cell("高", { width: 1200, align: "center", fill: LF }), cell("低", { width: 1000, align: "center", fill: LF }), cell("P2", { width: 1800, align: "center", fill: "E6F4EA" })]),
];

children.push(new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: [2800, 1200, 1000, 1200, 1000, 1800],
  rows: prioRows
}));

children.push(p(""));
children.push(h("六、风控原则（所有AI场景通用）", HeadingLevel.HEADING_1));
children.push(bullet("AI是辅助，不是替代：所有AI输出的结论、代码、报告都必须经过人工审核确认"));
children.push(bullet("高敏感数据留口子：涉及薪酬、核心财务数据的场景，不开放AI直接访问原始数据"));
children.push(bullet("关键结果交叉验证：AI给出的数据结论，必须和预置的关键指标做一致性校验"));
children.push(bullet("出错可追溯：所有AI生成的内容，都要保留生成日志（用了什么数据、什么模型、什么时间）"));
children.push(bullet("灰度上线：新AI场景先小范围试点（P0用户），跑稳定了再推广"));

// ============ 生成 ============
const doc = new Document({
  styles: {
    default: { document: { run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: HF, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 0, keepNext: false, keepLines: false } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, color: "2C5282", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1, keepNext: false, keepLines: false } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1200, hanging: 300 } } } },
      ]},
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } }
    },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "AI在数据平台建设全生命周期的应用", size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "第 ", size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
        new TextRun({ children: [PageNumber.CURRENT], size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
        new TextRun({ text: " 页", size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } }),
      ]
    })] }) },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const out = "C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\AI在统一数据平台建设全生命周期的应用.docx";
  fs.writeFileSync(out, buffer);
  console.log("Saved: " + out);
});
