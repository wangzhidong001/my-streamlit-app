const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// ============================================================
// SLIDE DIMENSIONS
// ============================================================
pres.layout = "LAYOUT_16x9";
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.5;
const CONTENT_X = MARGIN;
const CONTENT_Y = MARGIN;
const CONTENT_W = SLIDE_W - 2 * MARGIN;
const CONTENT_H = SLIDE_H - 2 * MARGIN;

// ============================================================
// COLOR PALETTE - Midnight Navy
// ============================================================
const C = {
  navy:       "1F4E79",
  navyDark:   "163A5C",
  navyLight:  "2C5282",
  blue:       "3B82F6",
  teal:       "0D9488",
  green:      "16A34A",
  amber:      "D97706",
  red:        "DC2626",
  white:      "FFFFFF",
  offWhite:   "F8FAFC",
  grayLight:  "E2E8F0",
  grayMid:    "94A3B8",
  grayDark:   "475569",
  text:       "0F172A",
};

// Layer colors (matching the original diagram)
const LAYER_COLORS = {
  L7: { bg: "1E3A5F", text: C.white },
  L6: { bg: "2B4A6F", text: C.white },
  L5: { bg: "276749", text: C.white },
  L4: { bg: "B7791F", text: C.white },
  L3: { bg: "B7791F", text: C.white },
  L2: { bg: "9B2C2C", text: C.white },
  L1: { bg: "9B2C2C", text: C.white },
};

// ============================================================
// CONTAINER SYSTEM
// ============================================================
function parseImageDimensions(path) {
  const match = path.match(/_(\d+)x(\d+)\.(png|jpg|jpeg|gif|webp)$/i);
  if (match) return { width: parseInt(match[1]), height: parseInt(match[2]) };
  return null;
}
function calculateScaledImageOpts(opts) {
  const { path, w: targetW, h: targetH, x = 0, y = 0, ...rest } = opts;
  if (!path || !targetW || !targetH) return opts;
  const dims = parseImageDimensions(path);
  if (!dims) return opts;
  const imgAspect = dims.width / dims.height;
  const targetAspect = targetW / targetH;
  let scaledW, scaledH, offsetX = 0, offsetY = 0;
  if (imgAspect > targetAspect) { scaledW = targetW; scaledH = targetW / imgAspect; offsetY = (targetH - scaledH) / 2; }
  else { scaledH = targetH; scaledW = targetH * imgAspect; offsetX = (targetW - scaledW) / 2; }
  return { path, x: x + offsetX, y: y + offsetY, w: scaledW, h: scaledH, ...rest };
}
function createVirtualNode(type, data, parentX = 0, parentY = 0) {
  const opts = data.opts || {};
  const node = { type, data, absX: parentX + (opts.x || 0), absY: parentY + (opts.y || 0), w: opts.w || 0, h: opts.h || 0, children: [] };
  node.addShape = function(shapeType, opts = {}) { const child = createVirtualNode("shape", { shapeType, opts }, node.absX, node.absY); node.children.push(child); return child; };
  node.addText = function(text, opts = {}) {
    const safeOpts = { fit: "shrink", ...opts };
    const child = createVirtualNode("text", { text, opts: safeOpts }, node.absX, node.absY);
    node.children.push(child); return child;
  };
  node.addImage = function(opts = {}) { const scaledOpts = calculateScaledImageOpts(opts); const child = createVirtualNode("image", { opts: scaledOpts }, node.absX, node.absY); node.children.push(child); return child; };
  node.addTable = function(tableData, opts = {}) { const child = createVirtualNode("table", { tableData, opts }, node.absX, node.absY); node.children.push(child); return child; };
  return node;
}
function flattenNode(node, realSlide) {
  const absOpts = { ...node.data.opts, x: node.absX, y: node.absY };
  if (node.type === "shape") realSlide.addShape(node.data.shapeType, absOpts);
  else if (node.type === "text") realSlide.addText(node.data.text, absOpts);
  else if (node.type === "image") realSlide.addImage(absOpts);
  else if (node.type === "table") realSlide.addTable(node.data.tableData, absOpts);
  node.children.forEach(child => flattenNode(child, realSlide));
}
const originalAddSlide = pres.addSlide.bind(pres);
pres.addSlide = function(options) {
  const realSlide = originalAddSlide(options);
  const virtualSlide = {
    children: [], _realSlide: realSlide,
    set background(val) { realSlide.background = val; },
    get background() { return realSlide.background; },
    addShape: function(shapeType, opts = {}) { const node = createVirtualNode("shape", { shapeType, opts }, 0, 0); this.children.push(node); return node; },
    addText: function(text, opts = {}) { const safeOpts = { fit: "shrink", ...opts }; const node = createVirtualNode("text", { text, opts: safeOpts }, 0, 0); this.children.push(node); return node; },
    addImage: function(opts = {}) { const scaledOpts = calculateScaledImageOpts(opts); const node = createVirtualNode("image", { opts: scaledOpts }, 0, 0); this.children.push(node); return node; },
    addTable: function(tableData, opts = {}) { const node = createVirtualNode("table", { tableData, opts }, 0, 0); this.children.push(node); return node; },
    addChart: function(chartType, data, opts = {}) { realSlide.addChart(chartType, data, opts); },
    addNotes: function(notes) { realSlide.addNotes(notes); },
    render: function() { this.children.forEach(child => flattenNode(child, realSlide)); }
  };
  return virtualSlide;
};

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function addSlideHeader(slide, kicker, title) {
  slide.addText(kicker, {
    x: CONTENT_X, y: 0.3, w: CONTENT_W, h: 0.3,
    fontSize: 11, color: C.navy, bold: true, charSpacing: 3,
    fontFace: "Arial"
  });
  slide.addText(title, {
    x: CONTENT_X, y: 0.55, w: CONTENT_W, h: 0.55,
    fontSize: 28, bold: true, color: C.text,
    fontFace: "Georgia", charSpacing: 1.5
  });
  // Accent line
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: 1.15, w: 0.8, h: 0.04,
    fill: { color: C.navy }, line: { color: C.navy }
  });
}

// ============================================================
// SLIDE 1: COVER
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.navyDark };

  // Decorative shape
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.15,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.15, w: SLIDE_W, h: 0.15,
    fill: { color: C.blue }, line: { color: C.blue }
  });

  slide.addText("Unified Data Platform Architecture", {
    x: CONTENT_X, y: 1.5, w: CONTENT_W, h: 0.4,
    fontSize: 14, color: C.blue, bold: true, charSpacing: 4,
    fontFace: "Arial"
  });

  slide.addText("统一数据平台\n七层架构", {
    x: CONTENT_X, y: 1.9, w: CONTENT_W, h: 1.8,
    fontSize: 42, bold: true, color: C.white,
    fontFace: "Georgia", charSpacing: 2.5, lineSpacingMultiple: 1.1
  });

  slide.addText("从数据源到应用输出，一层共享、多层演绎\n支撑法定财报 · 阿米巴管报 · BG经营报表", {
    x: CONTENT_X, y: 3.8, w: CONTENT_W, h: 0.8,
    fontSize: 16, color: C.grayMid,
    fontFace: "Arial", lineSpacingMultiple: 1.5
  });

  // Stats row
  const statsY = 4.8;
  const statW = (CONTENT_W - 0.6) / 3;
  [
    { n: "7", l: "层架构" },
    { n: "3", l: "套报表" },
    { n: "1", l: "个数据底座" },
  ].forEach((s, i) => {
    const sx = CONTENT_X + i * (statW + 0.3);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: statsY, w: statW, h: 0.5,
      fill: { color: "223A5C" }, line: { color: C.navyLight }
    });
    slide.addText(s.n, {
      x: sx, y: statsY + 0.05, w: statW, h: 0.25,
      fontSize: 20, bold: true, color: C.blue, align: "center",
      fontFace: "Georgia"
    });
    slide.addText(s.l, {
      x: sx, y: statsY + 0.28, w: statW, h: 0.2,
      fontSize: 11, color: C.grayMid, align: "center",
      fontFace: "Arial"
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 2: TOC
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  addSlideHeader(slide, "AGENDA", "今天讲什么");

  const items = [
    { n: "01", t: "架构总览", d: "七层架构一张图看懂" },
    { n: "02", t: "底三层详解", d: "数据源 · ODS贴源层 · 数仓层" },
    { n: "03", t: "上四层详解", d: "ADS · 指标中心 · 数据服务 · 应用输出" },
    { n: "04", t: "核心价值", d: "为什么要这样设计" },
  ];

  items.forEach((it, i) => {
    const y = 1.5 + i * 0.95;
    // Number circle
    slide.addShape(pres.shapes.OVAL, {
      x: CONTENT_X, y: y, w: 0.65, h: 0.65,
      fill: { color: C.navy }, line: { color: C.navy }
    });
    slide.addText(it.n, {
      x: CONTENT_X, y: y + 0.1, w: 0.65, h: 0.45,
      fontSize: 20, bold: true, color: C.white, align: "center",
      fontFace: "Georgia"
    });
    // Title
    slide.addText(it.t, {
      x: CONTENT_X + 0.9, y: y + 0.05, w: CONTENT_W - 0.9, h: 0.35,
      fontSize: 20, bold: true, color: C.text,
      fontFace: "Georgia"
    });
    // Description
    slide.addText(it.d, {
      x: CONTENT_X + 0.9, y: y + 0.4, w: CONTENT_W - 0.9, h: 0.3,
      fontSize: 13, color: C.grayDark,
      fontFace: "Arial"
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 3: 七层架构总览
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  addSlideHeader(slide, "ARCHITECTURE OVERVIEW · 架构总览", "统一数据平台 · 七层架构");
  slide.addText("（从上到下 = 第7层 → 第1层，最底层是原始数据源）", {
    x: CONTENT_X, y: 1.2, w: CONTENT_W, h: 0.3,
    fontSize: 10, color: C.grayDark, italic: true,
    fontFace: "Arial"
  });

  // Layer data (top to bottom: L7 → L1)
  const layers = [
    {
      key: "L7", num: "第7层", name: "应用输出层", sub: "3套报表 + AI应用",
      items: ["法定财报合并", "阿米巴考核管报", "BG经营报表", "智能指标助手", "自然语言取数"]
    },
    {
      key: "L6", num: "第6层", name: "数据服务层", sub: "API / 指标服务",
      items: ["统一数据API网关", "指标服务接口", "报表服务", "权限控制"]
    },
    {
      key: "L5", num: "第5层", name: "指标中心层", sub: "规则引擎",
      items: ["原子指标池", "映射中心", "分摊引擎", "口径版本管理", "血缘追踪"]
    },
    {
      key: "L4", num: "第4层", name: "数据仓库层", sub: "ADS 应用层",
      items: ["阿米巴主题宽表", "财报主题宽表", "BG经营主题宽表"]
    },
    {
      key: "L3", num: "第3层", name: "数仓汇总+明细", sub: "DWS 汇总层 + DWD 明细层",
      items: ["DWS：收入/成本/费用/人力域", "DWD：事实表+维度表，清洗后业务明细"]
    },
    {
      key: "L2", num: "第2层", name: "贴源层 ODS", sub: "",
      items: ["业务系统原始数据落地，保留原始语义，不做加工，只做格式统一"]
    },
    {
      key: "L1", num: "第1层", name: "数据源层", sub: "最底层",
      items: ["ERP", "总账系统", "CRM", "HR系统", "项目管理", "工时系统", "手工Excel"]
    },
  ];

  const labelW = 1.3;
  const startY = 1.5;
  const layerH = 0.52;
  const gap = 0.05;

  layers.forEach((l, i) => {
    const y = startY + i * (layerH + gap);
    const lc = LAYER_COLORS[l.key];

    // Label bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: CONTENT_X, y, w: labelW, h: layerH,
      fill: { color: lc.bg }, line: { color: lc.bg }
    });
    slide.addText(l.num + (l.sub ? " · " + l.sub : ""), {
      x: CONTENT_X + 0.05, y: y + 0.02, w: labelW - 0.1, h: layerH - 0.04,
      fontSize: 9, bold: true, color: lc.text, align: "center",
      fontFace: "Arial", valign: "middle"
    });

    // Content bar
    const contentX = CONTENT_X + labelW + 0.08;
    const contentW = CONTENT_W - labelW - 0.08;

    slide.addShape(pres.shapes.RECTANGLE, {
      x: contentX, y, w: contentW, h: layerH,
      fill: { color: "FFFFFF" }, line: { color: C.grayLight }
    });

    // Items
    const n = l.items.length;
    const itemW = (contentW - 0.1) / n;
    l.items.forEach((item, j) => {
      const ix = contentX + 0.05 + j * itemW;
      slide.addText(item, {
        x: ix, y: y + 0.08, w: itemW - 0.05, h: layerH - 0.16,
        fontSize: 9, color: C.text, align: "center",
        fontFace: "Arial", valign: "middle",
        fit: "shrink"
      });
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 4: 底三层详解
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  addSlideHeader(slide, "BOTTOM LAYERS · 数据底座", "底四层：把数据\"收进来、洗干净、整理好\"");

  const cardW = (CONTENT_W - 0.3) / 2;
  const cardH = 1.95;
  const row1Y = 1.45;
  const row2Y = row1Y + cardH + 0.15;

  const cards = [
    { y: row1Y, x: CONTENT_X, color: C.red, title: "第1层 · 数据源层",
      items: ["7大来源：ERP、总账、CRM、HR、项目管理、工时系统、手工Excel", "原则：只取一次，不重复采集", "关键点：必须在这一层就做数据质量评估"] },
    { y: row1Y, x: CONTENT_X + cardW + 0.3, color: C.red, title: "第2层 · 贴源层 ODS",
      items: ["定位：业务系统原始数据原样落地", "原则：保留原始语义，不做加工，只做格式统一", "为什么要这层：出了问题可以回溯到最原始的数据"] },
    { y: row2Y, x: CONTENT_X, color: C.amber, title: "第3层 · 数仓 DWD + DWS",
      items: ["DWD明细层：清洗后的业务明细，统一建模的事实表+维度表", "DWS汇总层：按业务域汇总（收入域、成本域、费用域、人力域）", "核心动作：去重、补全、标准化编码", "价值：三套报表共用这一层，不用各算各的"] },
    { y: row2Y, x: CONTENT_X + cardW + 0.3, color: C.amber, title: "第4层 · ADS 应用层",
      items: ["定位：按报表主题预聚合的宽表", "三张核心宽表：阿米巴/财报/BG经营", "价值：报表直接查，不用每次重算"] },
  ];

  cards.forEach(c => {
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: cardW, h: cardH,
      fill: { color: "FFFFFF" }, line: { color: C.grayLight }
    });
    // Accent bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: 0.06, h: cardH,
      fill: { color: c.color }, line: { color: c.color }
    });
    // Title
    slide.addText(c.title, {
      x: c.x + 0.2, y: c.y + 0.12, w: cardW - 0.25, h: 0.35,
      fontSize: 16, bold: true, color: c.color,
      fontFace: "Georgia"
    });
    // Items
    c.items.forEach((item, i) => {
      slide.addText("• " + item, {
        x: c.x + 0.2, y: c.y + 0.52 + i * 0.3, w: cardW - 0.3, h: 0.28,
        fontSize: 11, color: C.text,
        fontFace: "Arial", lineSpacingMultiple: 1.2
      });
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 5: 上四层详解
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  addSlideHeader(slide, "TOP LAYERS · 规则与应用", "上三层：把规则\"管起来\"，把数据\"用出去\"");

  const cardW = (CONTENT_W - 0.3) / 2;
  const cardH = 1.85;
  const row1Y = 1.5;
  const row2Y = row1Y + cardH + 0.2;

  const cards = [
    { y: row1Y, x: CONTENT_X, color: C.green, title: "第5层 · 指标中心层",
      items: ["原子指标池：统一定义所有指标是什么、怎么算", "映射中心：人→岗→部门→法人→阿米巴→BG的多维关系", "分摊引擎：5种分摊方式（固定比例、动因、阶梯、直接归属、工时/面积）", "口径版本管理：规则改了能回到旧口径对比", "血缘追踪：每个指标能追到最底层的来源字段"] },
    { y: row1Y, x: CONTENT_X + cardW + 0.3, color: C.navyLight, title: "第6层 · 数据服务层",
      items: ["统一API网关：所有外部系统都从这里取数，不直接连数仓", "指标服务接口：按标准接口查询任何指标", "报表服务：定时生成、推送、订阅报表", "权限控制：数据权限、功能权限分离，支持行级隔离"] },
    { y: row2Y, x: CONTENT_X, color: C.navy, title: "第7层 · 3套报表",
      items: ["法定财报合并：面向监管和审计，口径严格不变", "阿米巴考核管报：面向内部经营，口径灵活可调整", "BG经营报表：面向战略决策，BG可自定义口径"] },
    { y: row2Y, x: CONTENT_X + cardW + 0.3, color: C.blue, title: "第7层 · AI应用",
      items: ["智能指标口径助手：问AI就能查到指标定义、公式、来源", "自然语言取数：用中文提问，AI自动查数据出表", "智能血缘与根因：指标异常时自动找原因", "规则变更评估：改规则前先模拟影响"] },
  ];

  cards.forEach(c => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: cardW, h: cardH,
      fill: { color: "FFFFFF" }, line: { color: C.grayLight }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: 0.06, h: cardH,
      fill: { color: c.color }, line: { color: c.color }
    });
    slide.addText(c.title, {
      x: c.x + 0.2, y: c.y + 0.12, w: cardW - 0.25, h: 0.35,
      fontSize: 16, bold: true, color: c.color,
      fontFace: "Georgia"
    });
    c.items.forEach((item, i) => {
      slide.addText("• " + item, {
        x: c.x + 0.2, y: c.y + 0.5 + i * 0.27, w: cardW - 0.3, h: 0.25,
        fontSize: 10.5, color: C.text,
        fontFace: "Arial", lineSpacingMultiple: 1.2
      });
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 6: 核心价值
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  addSlideHeader(slide, "CORE VALUE · 核心价值", "为什么要这样设计？");

  const values = [
    { icon: "🎯", title: "口径统一", desc: "同一指标在不同报表中的定义被记录和管理，不再各说各话" },
    { icon: "⚡", title: "变更提速", desc: "阿米巴规则调整从\"等2-3周\"变成\"配置一下当天生效\"" },
    { icon: "🔍", title: "血缘可追", desc: "每个数都能追到最底层的来源，出了问题快速定位" },
    { icon: "💰", title: "成本降低", desc: "数据只采一次、不用重复建设ETL，省人省钱省时间" },
    { icon: "🛡️", title: "质量可控", desc: "统一的质量监控和校验规则，数据可信度大幅提升" },
    { icon: "🤖", title: "AI赋能", desc: "为AI应用提供干净、标准、可追溯的数据底座" },
  ];

  const cols = 3;
  const rows = 2;
  const cardW = (CONTENT_W - 0.4) / cols;
  const cardH = 1.6;
  const gapX = 0.2;
  const gapY = 0.2;
  const startY = 1.55;

  values.forEach((v, i) => {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = CONTENT_X + c * (cardW + gapX);
    const y = startY + r * (cardH + gapY);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: "FFFFFF" }, line: { color: C.grayLight }
    });
    // Top accent
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: 0.04,
      fill: { color: C.navy }, line: { color: C.navy }
    });

    slide.addText(v.icon, {
      x, y: y + 0.15, w: cardW, h: 0.5,
      fontSize: 28, align: "center",
      fontFace: "Arial"
    });
    slide.addText(v.title, {
      x, y: y + 0.65, w: cardW, h: 0.3,
      fontSize: 15, bold: true, color: C.navy, align: "center",
      fontFace: "Georgia"
    });
    slide.addText(v.desc, {
      x: x + 0.15, y: y + 0.95, w: cardW - 0.3, h: 0.6,
      fontSize: 10.5, color: C.grayDark, align: "center",
      fontFace: "Arial", lineSpacingMultiple: 1.3
    });
  });

  slide.render();
}

// ============================================================
// SLIDE 7: CLOSING
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.navyDark };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.15,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.15, w: SLIDE_W, h: 0.15,
    fill: { color: C.blue }, line: { color: C.blue }
  });

  slide.addText("谢谢", {
    x: 0, y: 1.8, w: SLIDE_W, h: 1.2,
    fontSize: 52, bold: true, color: C.white, align: "center",
    fontFace: "Georgia", charSpacing: 3
  });

  slide.addText("一个数据底座 · 七层架构 · 支撑三套报表\n\n数据只取一次，上层各算各的", {
    x: CONTENT_X + 1, y: 3.3, w: CONTENT_W - 2, h: 1.2,
    fontSize: 16, color: C.grayMid, align: "center",
    fontFace: "Arial", lineSpacingMultiple: 1.6
  });

  slide.render();
}

// ============================================================
// SAVE
// ============================================================
const outPath = "C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\统一数据平台七层架构.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("Saved: " + outPath);
}).catch(e => {
  console.error("Error:", e);
});
