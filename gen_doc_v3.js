const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const HF = "1F4E79"; // header fill
const LF = "F2F7FB"; // light fill

function h(text, level) {
  return new Paragraph({
    heading: level,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 100, line: 340 },
    ...opts,
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 21, ...opts.run })]
  });
}

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 60, line: 340 },
    children: [new TextRun({ text, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" }, size: 21 })]
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
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 300 },
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

// ========== 文档内容 ==========
const children = [];

// 封面
children.push(new Paragraph({ spacing: { before: 3200 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 300 },
  children: [new TextRun({ text: "统一数据平台与指标中心建设方案", bold: true, size: 48, color: HF, font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 200 },
  children: [new TextRun({ text: "数据底座 · 指标中台 · 智能应用 三位一体", size: 26, color: "4A5568", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 500 },
  children: [new TextRun({ text: "2026年7月", size: 22, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
children.push(pb());

// 一、方案总览
children.push(h("一、方案总览", HeadingLevel.HEADING_1));
children.push(p("解决的核心问题：三套报表（法定财报、阿米巴管报、BG经营报表）各取各的数、各算各的口径，数据重复采集3次，指标对不齐，改个规则要等数周。"));
children.push(p("一句话方案：建一个统一数据平台（数据只取一次）+ 一个指标中心（口径统一管理），底层共享、上层各算各的。"));

children.push(h("4个核心痛点", HeadingLevel.HEADING_2));
children.push(bullet("数据重复采集：三套报表各自从ERP/CRM/HR取数"));
children.push(bullet("口径对不齐：同一指标在不同报表定义不同"));
children.push(bullet("变更响应慢：阿米巴规则调整要等2-3周"));
children.push(bullet("问题找不到根：数据血缘断裂，出了错难定位"));
children.push(pb());

// 二、统一数据平台架构
children.push(h("二、统一数据平台整体架构", HeadingLevel.HEADING_1));
children.push(p("平台共分7层，最底层是数据源，最顶层是应用输出。越往下越原始，越往上越加工。"));

children.push(h("2.1 七层架构图（最底层在下面，最高层在上面）", HeadingLevel.HEADING_2));
children.push(new Paragraph({ spacing: { before: 100, after: 80 }, children: [
  new TextRun({ text: "（从上到下 = 第7层 → 第1层，第1层是原始数据）", italics: true, size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })
]}));

const archRows = [
  row([cell("第7层：应用输出层（3套报表 + AI应用）", { header: true, width: 9000 })]),
  row([cell("法定财报合并  |  阿米巴考核管报  |  BG经营报表  |  智能指标助手  |  自然语言取数", { width: 9000, fill: "EBF5FF" })]),

  row([cell("第6层：数据服务层（API / 指标服务）", { header: true, width: 9000 })]),
  row([cell("统一数据API网关  |  指标服务接口  |  报表服务  |  权限控制", { width: 9000, fill: "EBF5FF" })]),

  row([cell("第5层：指标中心层（规则引擎）", { header: true, width: 9000 })]),
  row([cell("原子指标池  |  映射中心  |  分摊引擎  |  口径版本管理  |  血缘追踪", { width: 9000, fill: "E6F4EA" })]),

  row([cell("第4层：数据仓库层（ADS应用层）", { header: true, width: 9000 })]),
  row([cell("按报表主题预聚合的宽表：阿米巴主题宽表  |  财报主题宽表  |  BG经营主题宽表", { width: 9000, fill: "FFF7E6" })]),

  row([cell("第3层：数据仓库层（DWS汇总层 + DWD明细层）", { header: true, width: 9000 })]),
  row([cell("DWS：按业务域汇总（收入域、成本域、费用域、人力域）\nDWD：统一建模的事实表+维度表，清洗后的业务明细", { width: 9000, fill: "FFF7E6" })]),

  row([cell("第2层：贴源层（ODS）", { header: true, width: 9000 })]),
  row([cell("业务系统原始数据落地，保留原始语义，不做加工，只做格式统一\n→ ERP/总账/CRM/HR/项目系统/手工填报", { width: 9000, fill: "FCE4EC" })]),

  row([cell("第1层：数据源层（最底层）", { header: true, width: 9000 })]),
  row([cell("ERP（订单/库存）  |  总账系统  |  CRM（客户/销售）  |  HR系统  |  项目管理  |  工时系统  |  手工Excel", { width: 9000, fill: "FCE4EC" })]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [9000], rows: archRows }));

children.push(p(""));
children.push(h("2.2 统一数据平台各层一句话", HeadingLevel.HEADING_2));
children.push(bullet("第1-2层（数据源+ODS）：数据只取一次，原样落地，所有报表共用这一层"));
children.push(bullet("第3层（DWD+DWS）：统一建模、统一主数据编码，是平台的核心价值层"));
children.push(bullet("第4层（ADS）：按报表主题预加工，报表直接从这里取数，不用再重算"));
children.push(bullet("第5层（指标中心）：定义指标是什么、怎么算、归谁管（详见第三章）"));
children.push(bullet("第6层（数据服务）：统一API出口，外面的系统都从这里取数，不用直接连数仓"));
children.push(bullet("第7层（应用输出）：3套报表+AI应用，各算各的，互不干扰"));

children.push(h("2.3 主数据管理（必须做）", HeadingLevel.HEADING_2));
children.push(p("统一数据平台的前提是主数据统一。以下主数据必须先做标准化编码："));
children.push(bullet("组织主数据：法人公司、部门、阿米巴单元、BG的统一编码及映射关系"));
children.push(bullet("人员主数据：员工ID、岗位、部门归属的版本化记录"));
children.push(bullet("客户主数据：客户ID、行业分类、归属销售组织"));
children.push(bullet("产品主数据：产品编码、产品线、细分市场分类"));
children.push(bullet("科目主数据：会计科目、报表项目、阿米巴指标的统一编码"));
children.push(pb());

// 三、指标中心
children.push(h("三、指标中心（四层指标体系）", HeadingLevel.HEADING_1));
children.push(p("指标中心解决的问题：同一笔业务事实在不同报表里叫不同名字、算不同口径。我们不强行统一，而是把关系说清楚。"));

children.push(h("3.1 四层指标图（L1在最下面，L4在最上面）", HeadingLevel.HEADING_2));

const indRows = [
  row([
    cell("层级", { header: true, width: 1500 }),
    cell("定位", { header: true, width: 1800 }),
    cell("示例", { header: true, width: 3200 }),
    cell("特征", { header: true, width: 2500 }),
  ]),
  row([
    cell("L4 分析层（最上）", { width: 1500, fill: "FCE4EC" }),
    cell("经营决策用的分析指标", { width: 1800, fill: "FCE4EC" }),
    cell("毛利率、人效、回款率、单客价值、库存周转天数", { width: 3200, fill: "FCE4EC" }),
    cell("由下层组合计算、口径灵活", { width: 2500, fill: "FCE4EC" }),
  ]),
  row([
    cell("L3 报表层", { width: 1500, fill: "EBF5FF" }),
    cell("监管/管理报表的格式化项目", { width: 1800, fill: "EBF5FF" }),
    cell("利润表项目、资产负债表项目、管理报表项目", { width: 3200, fill: "EBF5FF" }),
    cell("按模板聚合、有合并抵消、重分类", { width: 2500, fill: "EBF5FF" }),
  ]),
  row([
    cell("L2 财务核算层", { width: 1500, fill: "E6F4EA" }),
    cell("按会计准则确认的收入/成本/费用", { width: 1800, fill: "E6F4EA" }),
    cell("营业收入、营业成本、期间费用、应收账款", { width: 3200, fill: "E6F4EA" }),
    cell("口径严格、受准则约束", { width: 2500, fill: "E6F4EA" }),
  ]),
  row([
    cell("L1 业务运营层（最下）", { width: 1500, fill: "FFF7E6" }),
    cell("业务发生的原始事实", { width: 1800, fill: "FFF7E6" }),
    cell("订单量、发货量、签约金额、客户拜访次数", { width: 3200, fill: "FFF7E6" }),
    cell("粒度最细、实时性高、不做财务加工", { width: 2500, fill: "FFF7E6" }),
  ]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1500, 1800, 3200, 2500], rows: indRows }));

children.push(p(""));
children.push(h("3.2 三层间关系（3种模式）", HeadingLevel.HEADING_2));

const relRows = [
  row([
    cell("模式", { header: true, width: 1500 }),
    cell("适用场景", { header: true, width: 2500 }),
    cell("示例", { header: true, width: 5000 }),
  ]),
  row([
    cell("一对一映射", { width: 1500 }),
    cell("L1→L2，有明确转换规则", { width: 2500 }),
    cell("订单金额（L1）→【剔除退款+按履约进度+价税分离】→营业收入（L2）\n例：100万订单（含13%税），1月签2月发货 → L2在2月确认=100÷1.13≈88.5万", { width: 5000 }),
  ]),
  row([
    cell("一对多拆解", { width: 1500, fill: LF }),
    cell("L1→多个L2，按维度拆分", { width: 2500, fill: LF }),
    cell("市场部总费用100万（L1）→【按项目工时分摊】→ A产品线40万+B产品线35万+总部25万", { width: 5000, fill: LF }),
  ]),
  row([
    cell("多对一聚合", { width: 1500 }),
    cell("多下层→L4，组合计算", { width: 2500 }),
    cell("单客贡献毛利（L4）= 营业收入（L2）÷活跃客户数（L1）- 营业成本（L2）÷活跃客户数（L1）", { width: 5000 }),
  ]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1500, 2500, 5000], rows: relRows }));

children.push(p(""));
children.push(h("3.3 指标元数据（每个指标都要登记）", HeadingLevel.HEADING_2));
children.push(bullet("基础信息：指标编码、名称、所属层级（L1-L4）、数据类型、计量单位"));
children.push(bullet("定义信息：业务定义、计算公式、数据来源（源系统+表+字段）"));
children.push(bullet("管理信息：更新频率、责任部门（Owner）、上级/下级指标、版本历史"));

children.push(h("3.4 血缘追踪与一致性校验", HeadingLevel.HEADING_2));
children.push(bullet("每个指标能追根溯源：点\"毛利率\"→看到=L2收入-L2成本→再追收入→追到ERP哪张表"));
children.push(bullet("纵向平衡：集团收入 = 各法人收入之和（合并前）"));
children.push(bullet("横向平衡：收入-成本=毛利，每月校验"));
children.push(bullet("外部一致：指标中心数据 vs ERP总账，差异<0.1%"));
children.push(pb());

// 四、核心能力模块
children.push(h("四、三大核心能力模块", HeadingLevel.HEADING_1));

children.push(h("4.1 映射中心：解决维度差异", HeadingLevel.HEADING_2));
children.push(p("三套报表维度不同——法定按法人、阿米巴按虚拟巴、BG按事业群。映射中心维护关系链："));
children.push(bullet("人 → 岗位 → 部门 → 法人公司 → 阿米巴单元 → BG事业群"));
children.push(bullet("每个关系都有版本：张三2025年1-6月归东区巴，7月起归西区巴（历史可追溯）"));
children.push(bullet("变更走审批，历史版本永久保留"));

children.push(p(""));
children.push(h("4.2 分摊引擎：解决费用分配", HeadingLevel.HEADING_2));

const allocRows = [
  row([cell("类型", { header: true, width: 1800 }), cell("适用场景", { header: true, width: 3000 }), cell("示例", { header: true, width: 4200 })]),
  row([cell("固定比例", { width: 1800 }), cell("总部职能费按预定比例", { width: 3000 }), cell("HR费用：A巴30人摊30%，B巴70人摊70%", { width: 4200 })]),
  row([cell("动因分摊", { width: 1800, fill: LF }), cell("IT费按实际使用量", { width: 3000, fill: LF }), cell("IT运维费：A巴1000次API调用摊40%，B巴1500次摊60%", { width: 4200, fill: LF })]),
  row([cell("阶梯分摊", { width: 1800 }), cell("收入越大分摊比例越低", { width: 3000 }), cell("总部管理费：<1000万分摊5%，1000-5000万3%，>5000万1%", { width: 4200 })]),
  row([cell("直接归属+剩余分摊", { width: 1800, fill: LF }), cell("能直接归的先归，剩下再分", { width: 3000, fill: LF }), cell("市场部费用：A项目的60万直接归A巴，剩余40万按收入比例分", { width: 4200, fill: LF })]),
  row([cell("工时/面积分摊", { width: 1800 }), cell("房租按面积、共享人力按工时", { width: 3000 }), cell("办公房租：A巴300㎡摊30%，B巴700㎡摊70%", { width: 4200 })]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1800, 3000, 4200], rows: allocRows }));

children.push(p(""));
children.push(h("4.3 三套报表差异化支撑", HeadingLevel.HEADING_2));

const rptRows = [
  row([
    cell("维度", { header: true, width: 1300 }),
    cell("法定财报", { header: true, width: 2500 }),
    cell("阿米巴管报", { header: true, width: 2600 }),
    cell("BG经营报表", { header: true, width: 2600 }),
  ]),
  row([cell("给谁看", { width: 1300 }), cell("CFO、审计、投资者", { width: 2500 }), cell("各巴负责人、财务BP", { width: 2600 }), cell("BG总裁、运营总监", { width: 2600 })]),
  row([cell("组织维度", { width: 1300, fill: LF }), cell("法人公司", { width: 2500, fill: LF }), cell("虚拟阿米巴单元（可跨法人）", { width: 2600, fill: LF }), cell("事业群/事业部", { width: 2600, fill: LF })]),
  row([cell("收入确认", { width: 1300 }), cell("会计准则（权责发生制）", { width: 2500 }), cell("内部交易价+外部收入拆分", { width: 2600 }), cell("BG自定义确认节点", { width: 2600 })]),
  row([cell("更新频率", { width: 1300, fill: LF }), cell("月/季/年", { width: 2500, fill: LF }), cell("周/月", { width: 2600, fill: LF }), cell("月/季", { width: 2600, fill: LF })]),
  row([cell("规则变更", { width: 1300 }), cell("极少（受准则约束）", { width: 2500 }), cell("频繁（每季度可能调）", { width: 2600 }), cell("中等（随组织架构）", { width: 2600 })]),
  row([cell("内部交易", { width: 1300, fill: LF }), cell("合并抵消（消掉）", { width: 2500, fill: LF }), cell("模拟内部定价，各巴都算损益", { width: 2600, fill: LF }), cell("视BG需要处理", { width: 2600, fill: LF })]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [1300, 2500, 2600, 2600], rows: rptRows }));
children.push(pb());

// 五、AI应用
children.push(h("五、AI应用场景（按优先级）", HeadingLevel.HEADING_1));

const aiRows = [
  row([
    cell("场景", { header: true, width: 2000 }),
    cell("说明", { header: true, width: 3800 }),
    cell("示例", { header: true, width: 3200 }),
  ]),
  row([cell("P0 智能指标口径助手", { width: 2000 }), cell("输入指标名，查到定义、公式、来源、责任部门", { width: 3800 }), cell("问：\"阿米巴收入和法定收入有什么不一样？\" → 自动对比", { width: 3200 })]),
  row([cell("P0 智能数据质量监控", { width: 2000, fill: LF }), cell("自动检测数据异常（突增/突降/缺失）", { width: 3800, fill: LF }), cell("交通费比上月涨300% → 自动告警+定位源系统录错", { width: 3200, fill: LF })]),
  row([cell("P1 智能血缘与根因分析", { width: 2000 }), cell("指标异常时自动沿血缘链追原因", { width: 3800 }), cell("毛利率降了 → 自动追溯：哪个产品线？收入降还是成本涨？", { width: 3200 })]),
  row([cell("P1 规则变更影响评估", { width: 2000, fill: LF }), cell("改规则前模拟：哪些指标会变、变多少", { width: 3800, fill: LF }), cell("IT费从\"按人数\"改\"按使用量\" → 模拟A巴增15万B巴减15万", { width: 3200, fill: LF })]),
  row([cell("P1 自然语言取数", { width: 2000 }), cell("用中文问问题，AI自动查数据出表", { width: 3800 }), cell("\"上个月EBG前5大客户收入排行\" → 自动出表", { width: 3200 })]),
  row([cell("P2 经营分析自动生成", { width: 2000, fill: LF }), cell("每月自动写经营分析初稿", { width: 3800, fill: LF }), cell("自动生成：\"本月收入同比增15%，主要来自SBG增长23%...\"", { width: 3200, fill: LF })]),
];

children.push(new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [2000, 3800, 3200], rows: aiRows }));
children.push(p("⚠️ 风控：AI是辅助不是替代。高敏感数据不开放AI查询，AI结果要和关键指标交叉验证。", { run: { bold: true, color: "C53030" } }));
children.push(pb());

// 六、实施路径
children.push(h("六、实施路径（三期，12个月）", HeadingLevel.HEADING_1));

children.push(h("第一期（0-4个月）：数据底座 + 法定财报跑通", HeadingLevel.HEADING_2));
children.push(bullet("接入ERP、总账、HR等核心系统数据"));
children.push(bullet("搭建ODS→DWD→DWS→ADS四层数仓，统一主数据编码"));
children.push(bullet("迁移法定财报到平台，证明平台能力"));
children.push(bullet("建指标中台基础框架（原子指标池、规则引擎雏形）"));
children.push(bullet("AI落地：智能数据质量监控试点"));

children.push(h("第二期（4-8个月）：阿米巴管报 + 映射中心", HeadingLevel.HEADING_2));
children.push(bullet("建好映射中心（人-岗-巴-法人-BG关系）"));
children.push(bullet("迁移阿米巴核心逻辑（收入拆分、成本分摊）"));
children.push(bullet("上线口径版本管理和模拟测算"));
children.push(bullet("AI落地：智能指标口径助手、规则变更影响评估"));

children.push(h("第三期（8-12个月）：BG经营报表 + 全面智能化", HeadingLevel.HEADING_2));
children.push(bullet("迁移BG经营报表，支持BG自定义口径"));
children.push(bullet("建统一报表门户和数据API服务层"));
children.push(bullet("试点开放自然语言取数"));
children.push(bullet("全面推广AI辅助的质量监控和根因分析"));

children.push(p(""));
children.push(h("5个关键成功因素", HeadingLevel.HEADING_2));
children.push(bullet("高层背书：CFO/COO牵头，这是\"一把手工程\""));
children.push(bullet("先跑通一套：先让法定财报稳定运行，再迁阿米巴和BG"));
children.push(bullet("指标中台是灵魂：投入资源打磨规则引擎"));
children.push(bullet("业务人员参与：口径是财务BP定的，不是IT定的"));
children.push(bullet("AI适度引入：AI提效，不替代人做判断"));

// ========== 生成 ==========
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
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } }
    },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "统一数据平台与指标中心建设方案", size: 17, color: "718096", font: { ascii: "Arial", eastAsia: "Microsoft YaHei" } })]
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
  const out = "C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\统一数据平台与指标中心建设方案（精简版）.docx";
  fs.writeFileSync(out, buffer);
  console.log("Saved: " + out);
});
