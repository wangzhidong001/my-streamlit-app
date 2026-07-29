const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
const SW = 10, SH = 5.625, M = 0.5;
const CX = M, CY = M, CW = SW - 2*M, CH = SH - 2*M;

// ========== 精致配色：Midnight Executive + Gold Accent ==========
const C = {
  navyDk:   "0B1E3F",   // 深海蓝 - 深色背景
  navy:     "1E3A5F",    // 主色
  navyMd:   "2C5282",    // 中蓝
  blue:     "3B82F6",    // 亮蓝
  gold:     "D4A843",      // 金色强调色
  goldLt:    "F5E6C8",     // 浅金
  white:    "FFFFFF",
  offW:   "F5F7FA",       // 浅灰背景
  grayL:  "E2E8F0",
  grayM:  "94A3B8",
  grayD:  "475569",
  text:   "1A2332",         // 深文字
  textLt:  "64748B",
  green:  "0D9488",
  red:    "DC2626",
  amber:  "D97706",
};

// 层配色（更和谐的渐变色）
const LC = {
  L7: { bg: "1E3A5F", sub: "60A5FA" },
  L6: { bg: "1E40AF", sub: "93C5FD" },
  L5: { bg: "0369A1", sub: "7DD3FC" },
  L4: { bg: "0D9488", sub: "5EEAD4" },
  L3: { bg: "059669", sub: "6EE7B7" },
  L2: { bg: "D97706", sub: "FCD34D" },
  L1: { bg: "B45309", sub: "FBBF24" },
};

// ========== Container System ==========
function sImg(o){const m=o.path.match(/_(\d+)x(\d+)\./);if(!m)return o;
  const iw=+m[1],ih=+m[2],ia=iw/ih,ta=o.w/o.h;
  let sw,sh,ox=0,oy=0;if(ia>ta){sw=o.w;sh=o.w/ia;oy=(o.h-sh)/2}else{sh=o.h;sw=o.h*ia;ox=(o.w-sw)/2}
  return{path:o.path,x:o.x+ox,y:o.y+oy,w:sw,h:sh,...o}}
function vNode(t,d,px=0,py=0){const o=d.opts||{};const n={type:t,d,absX:px+(o.x||0),absY:py+(o.y||0),w:o.w||0,h:o.h||0,children:[]};
n.addShape=(st,op={})=>{const c=vNode("shape",{shapeType:st,opts:op},n.absX,n.absY);n.children.push(c);return c};
n.addText=(tx,op={})=>{const so={fit:"shrink",...op};const c=vNode("text",{text:tx,opts:so},n.absX,n.absY);n.children.push(c);return c};
n.addImage=(op={})=>{const so=sImg(op);const c=vNode("image",{opts:so},n.absX,n.absY);n.children.push(c);return c};
n.addTable=(td,op={})=>{const c=vNode("table",{tableData:td,opts:op},n.absX,n.absY);n.children.push(c);return c};return n}
function fNode(n,s){const o={...n.d.opts,x:n.absX,y:n.absY};
if(n.type==="shape")s.addShape(n.d.shapeType,o);else if(n.type==="text")s.addText(n.d.text,o);
else if(n.type==="image")s.addImage(o);else if(n.type==="table")s.addTable(n.d.tableData,o);
n.children.forEach(c=>fNode(c,s))}
const oAdd=pres.addSlide.bind(pres);
pres.addSlide=function(opt){const rs=oAdd(opt);const vs={children:[],_realSlide:rs,
set background(v){rs.background=v},get background(){return rs.background},
addShape:(st,op={})=>{const n=vNode("shape",{shapeType:st,opts:op},0,0);vs.children.push(n);return n},
addText:(tx,op={})=>{const so={fit:"shrink",...op};const n=vNode("text",{text:tx,opts:so},0,0);vs.children.push(n);return n},
addImage:(op={})=>{const so=sImg(op);const n=vNode("image",{opts:so},0,0);vs.children.push(n);return n},
addTable:(td,op={})=>{const n=vNode("table",{tableData:td,opts:op},0,0);vs.children.push(n);return n},
addNotes:(n)=>{rs.addNotes(n)},render:function(){this.children.forEach(c=>fNode(c,rs))}};return vs};

// 全局表格样式
const gTd = { fontSize: 9.5, fontFace: "Arial", color: C.text, fill: { color: "FFFFFF" }, valign: "middle" };
const gTd2 = { ...gTd, fill: { color: "F8FAFC" }};
const gTh = { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center", valign: "middle" };

// ========== SLIDE 1: COVER ==========
{const s=pres.addSlide();
s.background={color:C.navyDk};
// 顶部装饰条
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.08,fill:{color:C.gold},line:{color:C.gold}});
// 底部渐变装饰
s.addShape(pres.shapes.RECTANGLE,{x:0,y:SH-0.08,w:SW,h:0.08,fill:{color:C.gold},line:{color:C.gold}});
// 左侧金色竖条
s.addShape(pres.shapes.RECTANGLE,{x:0.8,y:1.2,w:0.06,h:3.2,fill:{color:C.gold},line:{color:C.gold}});
// 装饰圆形
s.addShape(pres.shapes.OVAL,{x:SW-2.2,y:0.8,w:1.8,h:1.8,fill:{color:"12284D"},line:{color:"12284D"}});
s.addShape(pres.shapes.OVAL,{x:SW-1.8,y:SH-2.0,w:1.2,h:1.2,fill:{color:"12284D"},line:{color:"12284D"}});

s.addText("DATA PLATFORM & METRIC CENTER",{x:1.1,y:1.3,w:SW-2,h:0.35,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("统一数据平台\n与指标中心",{x:1.1,y:1.75,w:SW-2,h:1.5,fontSize:40,bold:true,color:C.white,fontFace:"Georgia",charSpacing:2,lineSpacingMultiple:1.15});
s.addText("建设方案",{x:1.1,y:3.3,w:SW-2,h:0.5,fontSize:22,color:C.grayM,fontFace:"Georgia"});
// 分隔线
s.addShape(pres.shapes.RECTANGLE,{x:1.1,y:3.95,w:1.5,h:0.03,fill:{color:C.gold},line:{color:C.gold}});
s.addText("数据底座 · 指标中台 · 智能应用  三位一体",{x:1.1,y:4.15,w:SW-2,h:0.35,fontSize:13,color:C.grayM,fontFace:"Arial"});
s.addText("2026 年 7 月",{x:0.5,y:SH-0.6,w:SW-1,h:0.3,fontSize:10,color:C.grayM,align:"right",fontFace:"Arial"});
s.render()}

// ========== SLIDE 2: TOC ==========
{const s=pres.addSlide();
s.background={color:C.offW};
// 顶部装饰
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
// 标题
s.addText("CONTENTS",{x:CX,y:0.35,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("目录",{x:CX,y:0.65,w:CW,h:0.55,fontSize:28,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.22,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

const items=[
  {n:"01",t:"方案总览",d:"背景 · 痛点 · 一句话方案"},
  {n:"02",t:"七层架构",d:"统一数据平台 · 分层详解"},
  {n:"03",t:"指标中心",d:"L1-L4四层体系 · 三种层间关系"},
  {n:"04",t:"核心能力",d:"映射中心 · 分摊引擎 · 三套报表"},
  {n:"05",t:"AI应用场景",d:"6个场景 · P0/P1/P2"},
  {n:"06",t:"实施路径",d:"三期12个月 · 5个成功关键"},
];
items.forEach((it,i)=>{
  const col=i%3,row=Math.floor(i/3);
  const x=CX+col*(CW/3+0.15),y=1.55+row*1.55;
  // 卡片
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w:CW/3-0.1,h:1.35,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.08});
  // 左侧色条
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w:0.08,h:1.35,fill:{color:C.navy},line:{color:C.navy},rectRadius:0.04});
  // 序号
  s.addText(it.n,{x:x+0.2,y:y+0.15,w:0.6,h:0.5,fontSize:22,bold:true,color:C.gold,fontFace:"Georgia"});
  s.addText(it.t,{x:x+0.85,y:y+0.2,w:CW/3-1.1,h:0.35,fontSize:15,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(it.d,{x:x+0.85,y:y+0.55,w:CW/3-1.1,h:0.3,fontSize:10.5,color:C.grayD,fontFace:"Arial"});
});
s.render()}

// ========== SLIDE 3: 方案总览 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("OVERVIEW",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("方案总览",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

// 左侧：一句话方案卡片
const lx=CX,ly=1.35,lw=CW*0.42;
s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:lx,y:ly,w:lw,h:3.85,fill:{color:C.navy},line:{color:C.navy},rectRadius:0.1});
// 内部装饰
s.addShape(pres.shapes.RECTANGLE,{x:lx,y:ly,w:lw,h:0.55,fill:{color:C.gold},line:{color:C.gold}});
s.addText("一句话方案",{x:lx+0.25,y:ly+0.12,w:lw-0.5,h:0.32,fontSize:13,bold:true,color:C.navyDk,charSpacing:2,fontFace:"Arial"});
s.addText("建一个",{x:lx+0.3,y:ly+0.75,w:lw-0.6,h:0.4,fontSize:16,color:C.white,fontFace:"Georgia"});
s.addText("统一数据平台",{x:lx+0.3,y:ly+1.1,w:lw-0.6,h:0.55,fontSize:24,bold:true,color:C.gold,fontFace:"Georgia"});
s.addText("（数据只取一次）",{x:lx+0.3,y:ly+1.6,w:lw-0.6,h:0.3,fontSize:12,color:C.grayM,fontFace:"Arial"});
s.addText("+ 一个",{x:lx+0.3,y:ly+2.0,w:lw-0.6,h:0.35,fontSize:16,color:C.white,fontFace:"Georgia"});
s.addText("指标中心",{x:lx+0.3,y:ly+2.32,w:lw-0.6,h:0.55,fontSize:24,bold:true,color:C.gold,fontFace:"Georgia"});
s.addText("（口径统一管理）",{x:lx+0.3,y:ly+2.82,w:lw-0.6,h:0.3,fontSize:12,color:C.grayM,fontFace:"Arial"});
// 分隔线
s.addShape(pres.shapes.RECTANGLE,{x:lx+0.3,y:ly+3.25,w:lw-0.6,h:0.02,fill:{color:"334155"},line:{color:"334155"}});
s.addText("底层共享 · 上层各算各的",{x:lx+0.3,y:ly+3.4,w:lw-0.6,h:0.35,fontSize:14,bold:true,color:C.white,fontFace:"Georgia"});

// 右侧：4个痛点
const px=CX+CW*0.42+0.2;
const pains=[
  {n:"01",t:"数据重复采集",d:"三套报表各自从ERP/CRM/HR取数，相同数据被抽3次"},
  {n:"02",t:"口径对不齐",d:"同一指标在不同报表中定义不同，数据打架"},
  {n:"03",t:"变更响应慢",d:"阿米巴考核规则每季度调整，传统开发要等2-3周"},
  {n:"04",t:"问题找不到根",d:"数据血缘断裂，报表出错了不知道源头在哪"},
];
pains.forEach((p,i)=>{
  const y=1.35+i*0.95;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:px,y,w:CW*0.56-0.2,h:0.8,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.06});
  // 红色左边
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:px,y,w:0.08,h:0.8,fill:{color:C.red},line:{color:C.red},rectRadius:0.04});
  s.addText(p.n,{x:px+0.2,y:y+0.12,w:0.5,h:0.55,fontSize:20,bold:true,color:C.red,fontFace:"Georgia"});
  s.addText(p.t,{x:px+0.75,y:y+0.12,w:CW*0.56-1.1,h:0.3,fontSize:14,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(p.d,{x:px+0.75,y:y+0.42,w:CW*0.56-1.1,h:0.3,fontSize:10.5,color:C.grayD,fontFace:"Arial"});
});
s.render()}

// ========== SLIDE 4: 七层架构整合页 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("ARCHITECTURE",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("七层架构总览 · 分层详解",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

// 左侧：七层架构图
const leftW = CW * 0.52;
const layers=[
  {k:"L7",n:"第7层",name:"应用输出层",sub:"3套报表+AI应用",items:["法定财报","阿米巴管报","BG经营","智能助手","自然语言"]},
  {k:"L6",n:"第6层",name:"数据服务层",sub:"统一API出口",items:["API网关","指标服务","报表服务","权限控制"]},
  {k:"L5",n:"第5层",name:"指标中心层",sub:"规则引擎",items:["原子指标","映射中心","分摊引擎","版本管理","血缘追踪"]},
  {k:"L4",n:"第4层",name:"数据仓库层",sub:"ADS 应用层",items:["阿米巴宽表","财报宽表","BG经营宽表"]},
  {k:"L3",n:"第3层",name:"数仓汇总+明细",sub:"DWS + DWD",items:["DWS业务域","DWD事实维度"]},
  {k:"L2",n:"第2层",name:"贴源层 ODS",sub:"原始数据落地",items:["原样落地，保留语义"]},
  {k:"L1",n:"第1层",name:"数据源层",sub:"7大来源",items:["ERP","总账","CRM","HR","项目","工时","Excel"]},
];
const lw=1.1,sy=1.35,lh=0.44,gp=0.03;
layers.forEach((l,i)=>{
  const y=sy+i*(lh+gp),lc=LC[l.k];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:CX,y,w:lw,h:lh,fill:{color:lc.bg},line:{color:lc.bg},rectRadius:0.04});
  s.addText(l.n,{x:CX+0.05,y:y+0.02,w:lw-0.1,h:lh*0.5,fontSize:8,bold:true,color:C.white,align:"center",fontFace:"Arial",valign:"middle"});
  s.addText(l.sub,{x:CX+0.05,y:y+lh*0.5,w:lw-0.1,h:lh*0.45,fontSize:7,color:lc.sub,align:"center",fontFace:"Arial",valign:"middle",fit:"shrink"});
  const cx2=CX+lw+0.05,cw2=leftW-lw-0.05;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:cx2,y,w:cw2,h:lh,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.04});
  const n=l.items.length,iw=(cw2-0.06)/n;
  l.items.forEach((it,j)=>{
    const ix=cx2+0.03+j*iw;
    s.addText(it,{x:ix,y:y+0.06,w:iw-0.02,h:lh-0.12,fontSize:7.5,color:C.text,align:"center",fontFace:"Arial",valign:"middle",fit:"shrink"});
  });
});

// 右侧：分层详解
const rx = CX + leftW + 0.15;
const rw = CW - leftW - 0.15;

// 上：底四层
s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:rx,y:1.35,w:rw,h:2.1,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.08});
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:1.35,w:rw,h:0.42,fill:{color:C.amber},line:{color:C.amber}});
s.addText("底四层 · 数据底座",{x:rx+0.15,y:1.38,w:rw-0.3,h:0.36,fontSize:12,bold:true,color:C.white,fontFace:"Georgia"});
s.addText("收进来 · 洗干净 · 整理好",{x:rx+rw-2.5,y:1.42,w:2.3,h:0.3,fontSize:9,color:"FEF3C7",align:"right",fontFace:"Arial"});

const bItems=[
  {t:"L1 数据源层",d:"7大来源；只取一次，不重复采集",c:C.amber},
  {t:"L2 贴源层 ODS",d:"原始数据原样落地；出问题可回溯",c:C.amber},
  {t:"L3 DWD+DWS",d:"清洗建模统一编码；三套报表共用",c:C.amber},
  {t:"L4 ADS应用层",d:"按报表主题预聚合；报表直接查",c:C.amber},
];
bItems.forEach((b,i)=>{
  const y=1.85+i*0.38;
  s.addShape(pres.shapes.OVAL,{x:rx+0.12,y:y+0.04,w:0.22,h:0.22,fill:{color:b.c},line:{color:b.c}});
  s.addText((i+1).toString(),{x:rx+0.12,y:y+0.05,w:0.22,h:0.2,fontSize:10,bold:true,color:C.white,align:"center",fontFace:"Arial"});
  s.addText(b.t,{x:rx+0.45,y:y,w:1.5,h:0.3,fontSize:10,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(b.d,{x:rx+1.95,y:y+0.02,w:rw-2.1,h:0.28,fontSize:9,color:C.grayD,fontFace:"Arial",lineSpacingMultiple:1.2});
});

// 下：上三层
s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:rx,y:3.55,w:rw,h:1.85,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.08});
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:3.55,w:rw,h:0.42,fill:{color:C.navy},line:{color:C.navy}});
s.addText("上三层 · 规则与应用",{x:rx+0.15,y:3.58,w:rw-0.3,h:0.36,fontSize:12,bold:true,color:C.white,fontFace:"Georgia"});
s.addText("管起来 · 用出去",{x:rx+rw-1.8,y:3.62,w:1.6,h:0.3,fontSize:9,color:"93C5FD",align:"right",fontFace:"Arial"});

const tItems=[
  {t:"L5 指标中心",d:"指标池+映射+分摊+版本+血缘",c:C.navy},
  {t:"L6 数据服务",d:"统一API网关+服务+权限",c:C.navy},
  {t:"L7 应用输出",d:"3套报表 + 2个AI应用",c:C.navy},
];
tItems.forEach((b,i)=>{
  const y=4.05+i*0.38;
  s.addShape(pres.shapes.OVAL,{x:rx+0.12,y:y+0.04,w:0.22,h:0.22,fill:{color:b.c},line:{color:b.c}});
  s.addText((i+1).toString(),{x:rx+0.12,y:y+0.05,w:0.22,h:0.2,fontSize:10,bold:true,color:C.white,align:"center",fontFace:"Arial"});
  s.addText(b.t,{x:rx+0.45,y:y,w:1.5,h:0.3,fontSize:10,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(b.d,{x:rx+1.95,y:y+0.02,w:rw-2.1,h:0.28,fontSize:9,color:C.grayD,fontFace:"Arial",lineSpacingMultiple:1.2});
});
s.render()}

// ========== SLIDE 5: 指标中心 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("METRIC CENTER",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("指标中心：四层指标体系",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

// 表格
const th={...gTh,fill:{color:C.navy}};
const rows=[
  [{text:"层级",options:th},{text:"定位",options:th},{text:"示例",options:th},{text:"核心特征",options:th}],
  [{text:"L4 分析层（最上）",options:{...gTd,fill:{color:"FEF2F2"}}},{text:"经营决策用的分析指标",options:{...gTd,fill:{color:"FEF2F2"}}},{text:"毛利率、人效、回款率、单客价值、库存周转天数",options:{...gTd,fill:{color:"FEF2F2"}}},{text:"由下层组合计算、口径灵活",options:{...gTd,fill:{color:"FEF2F2"}}}],
  [{text:"L3 报表层",options:{...gTd2,fill:{color:"EFF6FF"}}},{text:"监管/管理报表的格式化项目",options:{...gTd2,fill:{color:"EFF6FF"}}},{text:"利润表项目、资产负债表项目、管理报表项目",options:{...gTd2,fill:{color:"EFF6FF"}}},{text:"按模板聚合、有合并抵消、重分类",options:{...gTd2,fill:{color:"EFF6FF"}}}],
  [{text:"L2 财务核算层",options:{...gTd,fill:{color:"ECFDF5"}}},{text:"按会计准则确认的收入/成本/费用",options:{...gTd,fill:{color:"ECFDF5"}}},{text:"营业收入、营业成本、期间费用、应收账款",options:{...gTd,fill:{color:"ECFDF5"}}},{text:"口径严格、受准则约束",options:{...gTd,fill:{color:"ECFDF5"}}}],
  [{text:"L1 业务运营层（最下）",options:{...gTd2,fill:{color:"FFFBEB"}}},{text:"业务发生的原始事实",options:{...gTd2,fill:{color:"FFFBEB"}}},{text:"订单量、发货量、签约金额、客户拜访次数",options:{...gTd2,fill:{color:"FFFBEB"}}},{text:"粒度最细、实时性高、不做财务加工",options:{...gTd2,fill:{color:"FFFBEB"}}}],
];
s.addTable(rows,{x:CX,y:1.3,w:CW,h:2.3,colW:[CW*0.2,CW*0.22,CW*0.35,CW*0.23],rowH:[0.38,0.44,0.44,0.44,0.44],border:{pt:0.5,color:C.grayL},rowBorders:true});

// 三种模式标题
s.addText("三种层间关系模式",{x:CX,y:3.75,w:CW,h:0.32,fontSize:15,bold:true,color:C.text,fontFace:"Georgia"});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:4.08,w:0.5,h:0.03,fill:{color:C.gold},line:{color:C.gold}});

const modes=[
  {t:"一对一映射",d:"L1→L2，有明确转换规则",e:"订单金额→【剔除退款+履约进度+价税分离】→营业收入",c:"FEF3C7",tc:C.amber},
  {t:"一对多拆解",d:"L1→多个L2，按维度拆分",e:"市场部总费用→【按项目工时分摊】→各产品线费用",c:"DBEAFE",tc:C.navy},
  {t:"多对一聚合",d:"多下层→L4，组合计算",e:"单客贡献毛利 = 营业收入÷活跃客户数 - 营业成本÷活跃客户数",c:"D1FAE5",tc:C.green},
];
const mw=(CW-0.2)/3;
modes.forEach((m,i)=>{
  const x=CX+i*(mw+0.1);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:4.2,w:mw,h:1.2,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.08});
  s.addShape(pres.shapes.RECTANGLE,{x,y:4.2,w:mw,h:0.05,fill:{color:m.tc},line:{color:m.tc}});
  s.addText(m.t,{x:x+0.12,y:4.32,w:mw-0.24,h:0.28,fontSize:12,bold:true,color:m.tc,fontFace:"Georgia"});
  s.addText(m.d,{x:x+0.12,y:4.6,w:mw-0.24,h:0.2,fontSize:9,color:C.grayD,fontFace:"Arial"});
  s.addText(m.e,{x:x+0.12,y:4.82,w:mw-0.24,h:0.5,fontSize:9,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.25});
});
s.render()}

// ========== SLIDE 6: 核心能力 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("CORE MODULES",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("三大核心能力模块",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

// 4.1 映射中心
s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:CX,y:1.35,w:CW,h:1.1,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.08});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.35,w:0.12,h:1.1,fill:{color:C.navy},line:{color:C.navy}});
s.addText("4.1 映射中心 · 解决维度差异",{x:CX+0.3,y:1.42,w:CW-0.5,h:0.32,fontSize:14,bold:true,color:C.navy,fontFace:"Georgia"});
s.addText("人 → 岗位 → 部门 → 法人公司 → 阿米巴单元 → BG事业群",{x:CX+0.3,y:1.75,w:CW-0.5,h:0.28,fontSize:11,color:C.text,fontFace:"Arial"});
s.addText("每个关系都有版本（历史可追溯）· 变更走审批 · 历史版本永久保留",{x:CX+0.3,y:2.05,w:CW-0.5,h:0.28,fontSize:10.5,color:C.grayD,fontFace:"Arial"});

// 4.2 分摊引擎标题
s.addText("4.2 分摊引擎 · 5种分摊类型",{x:CX,y:2.6,w:CW,h:0.32,fontSize:14,bold:true,color:C.text,fontFace:"Georgia"});

// 分摊引擎表格
const th2={...gTh,fill:{color:C.navy}};
const alloc=[
  [{text:"类型",options:th2},{text:"适用场景",options:th2},{text:"示例",options:th2}],
  [{text:"固定比例分摊",options:gTd},{text:"总部职能费按预定比例",options:gTd},{text:"HR费用：A巴30人摊30%，B巴70人摊70%",options:gTd}],
  [{text:"动因分摊",options:gTd2},{text:"IT费按实际使用量",options:gTd2},{text:"IT运维费：A巴1000次API调用摊40%，B巴1500次摊60%",options:gTd2}],
  [{text:"阶梯分摊",options:gTd},{text:"收入越大分摊比例越低",options:gTd},{text:"总部管理费：<1000万分摊5%，1000-5000万3%，>5000万1%",options:gTd}],
  [{text:"直接归属+剩余分摊",options:gTd2},{text:"能直接归的先归，剩下再分",options:gTd2},{text:"市场部费用：A项目60万直接归A巴，剩余40万按收入比例分",options:gTd2}],
  [{text:"工时/面积分摊",options:gTd},{text:"房租按面积、共享人力按工时",options:gTd},{text:"办公房租：A巴300㎡摊30%，B巴700㎡摊70%",options:gTd}],
];
s.addTable(alloc,{x:CX,y:2.95,w:CW,h:2.45,colW:[CW*0.2,CW*0.28,CW*0.52],rowH:[0.36,0.36,0.36,0.36,0.36,0.36],border:{pt:0.5,color:C.grayL},rowBorders:true});
s.render()}

// ========== SLIDE 7: 三套报表差异化 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("REPORT DIFFERENTIATION",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("三套报表差异化支撑",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

const th3={fill:{color:C.navy},color:C.white,bold:true,fontSize:10,fontFace:"Arial",align:"center",valign:"middle"};
const rpt=[
  [{text:"维度",options:th3},{text:"法定财报",options:th3},{text:"阿米巴管报",options:th3},{text:"BG经营报表",options:th3}],
  [{text:"给谁看",options:gTd},{text:"CFO、审计、投资者",options:gTd},{text:"各巴负责人、财务BP",options:gTd},{text:"BG总裁、运营总监",options:gTd}],
  [{text:"组织维度",options:gTd2},{text:"法人公司",options:gTd2},{text:"虚拟阿米巴单元（可跨法人）",options:gTd2},{text:"事业群/事业部",options:gTd2}],
  [{text:"收入确认",options:gTd},{text:"会计准则（权责发生制）",options:gTd},{text:"内部交易价+外部收入拆分",options:gTd},{text:"BG自定义确认节点",options:gTd}],
  [{text:"更新频率",options:gTd2},{text:"月/季/年",options:gTd2},{text:"周/月",options:gTd2},{text:"月/季",options:gTd2}],
  [{text:"规则变更",options:gTd},{text:"极少（受准则约束）",options:gTd},{text:"频繁（每季度可能调）",options:gTd},{text:"中等（随组织架构）",options:gTd}],
  [{text:"内部交易",options:gTd2},{text:"合并抵消（消掉）",options:gTd2},{text:"模拟内部定价，各巴都算损益",options:gTd2},{text:"视BG需要处理",options:gTd2}],
];
s.addTable(rpt,{x:CX,y:1.3,w:CW,h:3.8,colW:[CW*0.16,CW*0.28,CW*0.3,CW*0.26],rowH:[0.42,0.42,0.42,0.42,0.42,0.42,0.42],border:{pt:0.5,color:C.grayL},rowBorders:true});
s.render()}

// ========== SLIDE 8: AI应用 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("AI SCENARIOS",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("AI应用场景（按优先级）",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

const th4={fill:{color:C.navy},color:C.white,bold:true,fontSize:10,fontFace:"Arial",align:"center",valign:"middle"};
const p0={bg:"FEE2E2"},p1={bg:"FEF3C7"},p2={bg:"DBEAFE"};
const ai=[
  [{text:"优先级",options:th4},{text:"场景",options:th4},{text:"说明",options:th4},{text:"示例",options:th4}],
  [{text:"P0",options:{...gTd,fill:p0.bg,color:C.red,bold:true}},{text:"智能指标口径助手",options:gTd},{text:"输入指标名，查到定义、公式、来源、责任部门",options:gTd},{text:"问：\"阿米巴收入和法定收入有什么不一样？\"→自动对比",options:gTd}],
  [{text:"P0",options:{...gTd2,fill:p0.bg,color:C.red,bold:true}},{text:"智能数据质量监控",options:gTd2},{text:"自动检测数据异常（突增/突降/缺失）",options:gTd2},{text:"交通费比上月涨300%→自动告警+定位源系统录错",options:gTd2}],
  [{text:"P1",options:{...gTd,fill:p1.bg,color:C.amber,bold:true}},{text:"智能血缘与根因分析",options:gTd},{text:"指标异常时自动沿血缘链追原因",options:gTd},{text:"毛利率降了→自动追溯：哪个产品线？收入降还是成本涨？",options:gTd}],
  [{text:"P1",options:{...gTd2,fill:p1.bg,color:C.amber,bold:true}},{text:"规则变更影响评估",options:gTd2},{text:"改规则前模拟：哪些指标会变、变多少",options:gTd2},{text:"IT费从\"按人数\"改\"按使用量\"→模拟A巴增15万B巴减15万",options:gTd2}],
  [{text:"P1",options:{...gTd,fill:p1.bg,color:C.amber,bold:true}},{text:"自然语言取数",options:gTd},{text:"用中文问问题，AI自动查数据出表",options:gTd},{text:"\"上个月EBG前5大客户收入排行\"→自动出表",options:gTd}],
  [{text:"P2",options:{...gTd2,fill:p2.bg,color:C.navy,bold:true}},{text:"经营分析自动生成",options:gTd2},{text:"每月自动写经营分析初稿",options:gTd2},{text:"自动生成：\"本月收入同比增15%，主要来自SBG增长23%...\"",options:gTd2}],
];
s.addTable(ai,{x:CX,y:1.3,w:CW,h:3.9,colW:[CW*0.07,CW*0.18,CW*0.3,CW*0.45],rowH:[0.38,0.44,0.44,0.44,0.44,0.44,0.44],border:{pt:0.5,color:C.grayL},rowBorders:true});

s.addText("⚠ 风控：AI是辅助不是替代。高敏感数据不开放AI查询，AI结果需交叉验证。",{x:CX,y:5.25,w:CW,h:0.28,fontSize:10,bold:true,color:C.red,fontFace:"Arial"});
s.render()}

// ========== SLIDE 9: 实施路径 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("ROADMAP",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("实施路径：三期 12 个月",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

const phases=[
  {n:"第一期",t:"0-4个月",sub:"数据底座 + 法定财报跑通",c:C.red,items:["接入ERP、总账、HR等核心系统数据","搭建ODS→DWD→DWS→ADS四层数仓","迁移法定财报到平台，证明平台能力","建指标中台基础框架","AI落地：智能数据质量监控试点"]},
  {n:"第二期",t:"4-8个月",sub:"阿米巴管报 + 映射中心",c:C.amber,items:["建好映射中心（人-岗-巴-法人-BG关系）","迁移阿米巴核心逻辑（收入拆分、成本分摊）","上线口径版本管理和模拟测算","AI落地：智能指标口径助手、规则变更评估"]},
  {n:"第三期",t:"8-12个月",sub:"BG经营报表 + 全面智能化",c:C.green,items:["迁移BG经营报表，支持BG自定义口径","建统一报表门户和数据API服务层","试点开放自然语言取数","全面推广AI辅助的质量监控和根因分析"]},
];
const pw=(CW-0.3)/3;
phases.forEach((ph,i)=>{
  const x=CX+i*(pw+0.15);
  // 卡片
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:1.35,w:pw,h:3.9,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.1});
  // 头部
  s.addShape(pres.shapes.RECTANGLE,{x,y:1.35,w:pw,h:0.85,fill:{color:ph.c},line:{color:ph.c}});
  s.addText(ph.n,{x,y:1.42,w:pw,h:0.35,fontSize:16,bold:true,color:C.white,align:"center",fontFace:"Georgia"});
  s.addText(ph.t,{x,y:1.78,w:pw,h:0.25,fontSize:11,color:C.white,align:"center",fontFace:"Arial"});
  s.addText(ph.sub,{x,y:2.0,w:pw,h:0.25,fontSize:10,color:C.navyDk,bold:true,align:"center",fontFace:"Arial"});
  // 分隔线
  s.addShape(pres.shapes.RECTANGLE,{x:x+0.2,y:2.28,w:pw-0.4,h:0.02,fill:{color:C.grayL},line:{color:C.grayL}});
  // 列表
  ph.items.forEach((it,j)=>{
    const y=2.4+j*0.52;
    s.addShape(pres.shapes.OVAL,{x:x+0.2,y:y+0.06,w:0.18,h:0.18,fill:{color:ph.c},line:{color:ph.c}});
    s.addText((j+1).toString(),{x:x+0.2,y:y+0.07,w:0.18,h:0.16,fontSize:9,bold:true,color:C.white,align:"center",fontFace:"Arial"});
    s.addText(it,{x:x+0.48,y:y+0.02,w:pw-0.6,h:0.48,fontSize:9.5,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.25});
  });
});
s.render()}

// ========== SLIDE 10: 关键成功因素 ==========
{const s=pres.addSlide();
s.background={color:C.offW};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.06,fill:{color:C.navy},line:{color:C.navy}});
s.addText("KEY SUCCESS FACTORS",{x:CX,y:0.3,w:CW,h:0.3,fontSize:11,color:C.gold,bold:true,charSpacing:4,fontFace:"Arial"});
s.addText("5个关键成功因素",{x:CX,y:0.58,w:CW,h:0.5,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.gold},line:{color:C.gold}});

const kfs=[
  {i:"①",t:"高层背书",d:"CFO/COO牵头\n这是\"一把手工程\""},
  {i:"②",t:"先跑通一套",d:"先让法定财报稳定运行\n再迁阿米巴和BG"},
  {i:"③",t:"指标中台是灵魂",d:"投入资源打磨规则引擎"},
  {i:"④",t:"业务人员参与",d:"口径是财务BP定的\n不是IT定的"},
  {i:"⑤",t:"AI适度引入",d:"AI提效\n不替代人做判断"},
];
const kw=(CW-0.5)/5;
kfs.forEach((k,i)=>{
  const x=CX+i*(kw+0.125);
  // 卡片
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:1.4,w:kw,h:3.7,fill:{color:C.white},line:{color:C.grayL},rectRadius:0.1});
  // 顶色条
  s.addShape(pres.shapes.RECTANGLE,{x,y:1.4,w:kw,h:0.08,fill:{color:C.gold},line:{color:C.gold}});
  s.addText(k.i,{x,y:1.6,w:kw,h:0.9,fontSize:40,bold:true,color:C.navy,align:"center",fontFace:"Georgia"});
  s.addText(k.t,{x:x+0.1,y:2.55,w:kw-0.2,h:0.4,fontSize:14,bold:true,color:C.text,align:"center",fontFace:"Georgia"});
  // 分隔线
  s.addShape(pres.shapes.RECTANGLE,{x:x+0.3,y:3.0,w:kw-0.6,h:0.02,fill:{color:C.grayL},line:{color:C.grayL}});
  s.addText(k.d,{x:x+0.15,y:3.15,w:kw-0.3,h:1.6,fontSize:11,color:C.grayD,align:"center",fontFace:"Arial",lineSpacingMultiple:1.5});
});
s.render()}

// ========== SLIDE 11: CLOSING ==========
{const s=pres.addSlide();
s.background={color:C.navyDk};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.08,fill:{color:C.gold},line:{color:C.gold}});
s.addShape(pres.shapes.RECTANGLE,{x:0,y:SH-0.08,w:SW,h:0.08,fill:{color:C.gold},line:{color:C.gold}});
// 左侧金色竖条
s.addShape(pres.shapes.RECTANGLE,{x:0.8,y:1.5,w:0.06,h:2.8,fill:{color:C.gold},line:{color:C.gold}});
// 装饰圆
s.addShape(pres.shapes.OVAL,{x:SW-2.0,y:1.0,w:1.6,h:1.6,fill:{color:"12284D"},line:{color:"12284D"}});

s.addText("谢谢",{x:1.1,y:1.55,w:SW-2,h:1.5,fontSize:52,bold:true,color:C.white,fontFace:"Georgia",charSpacing:3});
s.addShape(pres.shapes.RECTANGLE,{x:1.1,y:3.15,w:1.5,h:0.03,fill:{color:C.gold},line:{color:C.gold}});
s.addText("一个数据底座 · 七层架构 · 支撑三套报表\n\n数据只取一次 · 上层各算各的",{x:1.1,y:3.35,w:SW-2,h:1.2,fontSize:15,color:C.grayM,fontFace:"Arial",lineSpacingMultiple:1.6});
s.render()}

// ========== SAVE ==========
const out="C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\统一数据平台与指标中心建设方案_精美版.pptx";
pres.writeFile({fileName:out}).then(()=>console.log("Saved: "+out)).catch(e=>console.error(e));
