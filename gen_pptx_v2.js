const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
const SW = 10, SH = 5.625, M = 0.5;
const CX = M, CY = M, CW = SW - 2*M, CH = SH - 2*M;

const C = {
  navy: "1F4E79", navyDk: "163A5C", navyLt: "2C5282",
  blue: "3B82F6", teal: "0D9488", green: "16A34A",
  amber: "D97706", red: "DC2626",
  white: "FFFFFF", offW: "F8FAFC", grayL: "E2E8F0",
  grayM: "94A3B8", grayD: "475569", text: "0F172A",
};
const LC = {
  L7: { bg: "1E3A5F" }, L6: { bg: "2B4A6F" }, L5: { bg: "276749" },
  L4: { bg: "B7791F" }, L3: { bg: "B7791F" }, L2: { bg: "9B2C2C" }, L1: { bg: "9B2C2C" },
};
// 表格单元格样式（全局）
const td = { fontSize: 9.5, fontFace: "Arial", color: C.text, fill: { color: "FFFFFF" } };
const td2 = { fontSize: 9.5, fontFace: "Arial", color: C.text, fill: { color: C.offW } };

// ===== Container System =====
function pDim(p){const m=p.match(/_(\d+)x(\d+)\./);return m?{w:+m[1],h:+m[2]}:null}
function sImg(o){const{p:w1,h:h1}=o,x=o.x||0,y=o.y||0;const d=pDim(o.path);if(!d)return o;
const a=d.w/d.h,ta=w1/h1;let sw,sh,ox=0,oy=0;if(a>ta){sw=w1;sh=w1/a;oy=(h1-sh)/2}else{sh=h1;sw=h1*a;ox=(w1-sw)/2}
return{path:o.path,x:x+ox,y:y+oy,w:sw,h:sh,...o}}
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
addTable:(tableData,op={})=>{const c=vNode("table",{tableData,opts:op},0,0);vs.children.push(c);return c},
addNotes:(n)=>{rs.addNotes(n)},render:function(){this.children.forEach(c=>fNode(c,rs))}};return vs};

// ===== Helpers =====
function addHeader(s,k,t){
  s.addText(k,{x:CX,y:0.28,w:CW,h:0.28,fontSize:10,color:C.navy,bold:true,charSpacing:3,fontFace:"Arial"});
  s.addText(t,{x:CX,y:0.52,w:CW,h:0.55,fontSize:26,bold:true,color:C.text,fontFace:"Georgia",charSpacing:1.5});
  s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.1,w:0.7,h:0.04,fill:{color:C.navy},line:{color:C.navy}});
}

// ===== SLIDE 1: COVER =====
{const s=pres.addSlide();s.background={color:C.navyDk};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.12,fill:{color:C.blue},line:{color:C.blue}});
s.addShape(pres.shapes.RECTANGLE,{x:0,y:SH-0.12,w:SW,h:0.12,fill:{color:C.blue},line:{color:C.blue}});
s.addText("Unified Data Platform & Metric Center",{x:CX,y:1.2,w:CW,h:0.35,fontSize:13,color:C.blue,bold:true,charSpacing:3,fontFace:"Arial"});
s.addText("统一数据平台\n与指标中心建设方案",{x:CX,y:1.6,w:CW,h:1.6,fontSize:40,bold:true,color:C.white,fontFace:"Georgia",charSpacing:2.5,lineSpacingMultiple:1.1});
s.addText("数据底座 · 指标中台 · 智能应用 三位一体",{x:CX,y:3.4,w:CW,h:0.4,fontSize:16,color:C.grayM,fontFace:"Arial"});
const sy=4.3,sw=(CW-0.6)/3;
["7层架构","3套报表","1个数据底座"].forEach((t,i)=>{
  const x=CX+i*(sw+0.3);
  s.addShape(pres.shapes.RECTANGLE,{x,y:sy,w:sw,h:0.55,fill:{color:"223A5C"},line:{color:C.navyLt}});
  s.addText(t.charAt(0),{x,y:sy+0.05,w:sw,h:0.3,fontSize:20,bold:true,color:C.blue,align:"center",fontFace:"Georgia"});
  s.addText(t.slice(1),{x,y:sy+0.32,w:sw,h:0.2,fontSize:11,color:C.grayM,align:"center",fontFace:"Arial"});
});
s.addText("2026年7月",{x:CX,y:5.05,w:CW,h:0.3,fontSize:11,color:C.grayM,align:"right",fontFace:"Arial"});
s.render()}

// ===== SLIDE 2: TOC =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"CONTENTS","目录");
const items=[
  {n:"01",t:"方案总览",d:"背景 · 痛点 · 一句话方案"},
  {n:"02",t:"七层架构总览",d:"统一数据平台 · 从数据源到应用输出"},
  {n:"03",t:"底四层详解",d:"数据源 · ODS · DWD/DWS · ADS"},
  {n:"04",t:"上三层详解",d:"指标中心 · 数据服务 · 应用输出"},
  {n:"05",t:"指标中心体系",d:"L1-L4四层指标 · 三种层间关系模式"},
  {n:"06",t:"核心能力模块",d:"映射中心 · 分摊引擎 · 三套报表差异化支撑"},
  {n:"07",t:"AI应用场景",d:"6个场景 · 按P0/P1/P2优先级"},
  {n:"08",t:"实施路径",d:"三期12个月 · 5个成功关键因素"},
];
const colW=(CW-0.3)/2;
items.forEach((it,i)=>{
  const col=i%2,row=Math.floor(i/2);
  const x=CX+col*(colW+0.3),y=1.35+row*0.52;
  s.addShape(pres.shapes.OVAL,{x,y,w:0.45,h:0.45,fill:{color:C.navy},line:{color:C.navy}});
  s.addText(it.n,{x,y:y+0.06,w:0.45,h:0.33,fontSize:14,bold:true,color:C.white,align:"center",fontFace:"Georgia"});
  s.addText(it.t,{x:x+0.55,y:y+0.02,w:colW-0.6,h:0.25,fontSize:14,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(it.d,{x:x+0.55,y:y+0.26,w:colW-0.6,h:0.2,fontSize:10,color:C.grayD,fontFace:"Arial"});
});
s.render()}

// ===== SLIDE 3: 方案总览 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"OVERVIEW","方案总览");
// Left: 一句话方案
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.35,w:CW*0.48,h:1.55,fill:{color:C.navy},line:{color:C.navy}});
s.addText("一句话方案",{x:CX+0.2,y:1.5,w:CW*0.48-0.4,h:0.3,fontSize:12,color:C.blue,bold:true,fontFace:"Arial",charSpacing:2});
s.addText("建一个统一数据平台\n（数据只取一次）\n+ 一个指标中心\n（口径统一管理）",{x:CX+0.2,y:1.82,w:CW*0.48-0.4,h:0.85,fontSize:12,bold:true,color:C.white,fontFace:"Georgia",lineSpacingMultiple:1.15});
s.addText("底层共享、上层各算各的",{x:CX+0.2,y:2.7,w:CW*0.48-0.4,h:0.25,fontSize:10,color:C.grayM,fontFace:"Arial"});
// Right: 4个痛点
const px=CX+CW*0.52;
const pains=[
  {n:"01",t:"数据重复采集",d:"三套报表各自从ERP/CRM/HR取数，相同数据被抽3次"},
  {n:"02",t:"口径对不齐",d:"同一指标在不同报表中定义不同"},
  {n:"03",t:"变更响应慢",d:"阿米巴考核规则每季度调整，传统开发要等2-3周"},
  {n:"04",t:"问题找不到根",d:"数据血缘断裂，报表出错了不知道源头在哪"},
];
pains.forEach((p,i)=>{
  const y=1.35+i*0.78;
  s.addShape(pres.shapes.RECTANGLE,{x:px,y,w:CW*0.48,h:0.7,fill:{color:"FFFFFF"},line:{color:C.grayL}});
  s.addShape(pres.shapes.RECTANGLE,{x:px,y,w:0.05,h:0.7,fill:{color:C.red},line:{color:C.red}});
  s.addText(p.n,{x:px+0.15,y:y+0.08,w:0.4,h:0.55,fontSize:16,bold:true,color:C.red,fontFace:"Georgia"});
  s.addText(p.t,{x:px+0.6,y:y+0.1,w:CW*0.48-0.8,h:0.28,fontSize:13,bold:true,color:C.text,fontFace:"Georgia"});
  s.addText(p.d,{x:px+0.6,y:y+0.38,w:CW*0.48-0.8,h:0.28,fontSize:10,color:C.grayD,fontFace:"Arial"});
});
s.render()}

// ===== SLIDE 4: 七层架构 + 分层详解（整合页） =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"ARCHITECTURE","七层架构总览 · 分层详解");

// ========== 左侧：七层架构图（55%宽度） ==========
const leftW = CW * 0.54;
const layers=[
  {k:"L7",n:"第7层",name:"应用输出层",sub:"3套报表+AI",items:["法定财报","阿米巴管报","BG经营报表","智能助手","自然语言"]},
  {k:"L6",n:"第6层",name:"数据服务层",sub:"API/指标服务",items:["API网关","指标服务","报表服务","权限控制"]},
  {k:"L5",n:"第5层",name:"指标中心层",sub:"规则引擎",items:["原子指标","映射中心","分摊引擎","版本管理","血缘追踪"]},
  {k:"L4",n:"第4层",name:"数据仓库层",sub:"ADS 应用层",items:["阿米巴宽表","财报宽表","BG经营宽表"]},
  {k:"L3",n:"第3层",name:"数仓汇总+明细",sub:"DWS+DWD",items:["DWS汇总域","DWD事实+维度"]},
  {k:"L2",n:"第2层",name:"贴源层 ODS",sub:"",items:["原始数据落地，不做加工"]},
  {k:"L1",n:"第1层",name:"数据源层",sub:"最底层",items:["ERP","总账","CRM","HR","项目","工时","Excel"]},
];
const lw=1.0,sy=1.3,lh=0.44,gp=0.03;
layers.forEach((l,i)=>{
  const y=sy+i*(lh+gp),lc=LC[l.k];
  s.addShape(pres.shapes.RECTANGLE,{x:CX,y,w:lw,h:lh,fill:{color:lc.bg},line:{color:lc.bg}});
  s.addText(l.n+(l.sub?"·"+l.sub:""),{x:CX+0.03,y:y+0.02,w:lw-0.06,h:lh-0.04,fontSize:7.5,bold:true,color:C.white,align:"center",fontFace:"Arial",valign:"middle",fit:"shrink"});
  const cx2=CX+lw+0.04,cw2=leftW-lw-0.04;
  s.addShape(pres.shapes.RECTANGLE,{x:cx2,y,w:cw2,h:lh,fill:{color:"FFFFFF"},line:{color:C.grayL}});
  const n=l.items.length,iw=(cw2-0.04)/n;
  l.items.forEach((it,j)=>{
    const ix=cx2+0.02+j*iw;
    s.addText(it,{x:ix,y:y+0.06,w:iw-0.02,h:lh-0.12,fontSize:7.5,color:C.text,align:"center",fontFace:"Arial",valign:"middle",fit:"shrink"});
  });
});

// ========== 右侧：分层详解（45%宽度） ==========
const rx = CX + leftW + 0.12;
const rw = CW - leftW - 0.12;

// 上半部分：底四层
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:1.3,w:rw,h:2.1,fill:{color:"FFFFFF"},line:{color:C.grayL}});
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:1.3,w:rw,h:0.35,fill:{color:C.red},line:{color:C.red}});
s.addText("底四层 · 数据底座（收进来、洗干净、整理好）",{x:rx+0.1,y:1.33,w:rw-0.2,h:0.3,fontSize:10.5,bold:true,color:C.white,fontFace:"Georgia"});

const bottomItems=[
  {t:"L1 数据源层",d:"7大来源：ERP/总账/CRM/HR/项目/工时/Excel；只取一次"},
  {t:"L2 贴源层 ODS",d:"原始数据原样落地；保留语义；出问题可回溯"},
  {t:"L3 DWD+DWS",d:"清洗建模统一编码；按业务域汇总；三套报表共用"},
  {t:"L4 ADS应用层",d:"按报表主题预聚合；三张核心宽表；报表直接查"},
];
bottomItems.forEach((b,i)=>{
  const y=1.72+i*0.42;
  s.addText(b.t,{x:rx+0.1,y,w:1.4,h:0.3,fontSize:9.5,bold:true,color:C.red,fontFace:"Georgia"});
  s.addText(b.d,{x:rx+1.5,y,w:rw-1.6,h:0.35,fontSize:8.5,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.2});
});

// 下半部分：上三层
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:3.5,w:rw,h:1.8,fill:{color:"FFFFFF"},line:{color:C.grayL}});
s.addShape(pres.shapes.RECTANGLE,{x:rx,y:3.5,w:rw,h:0.35,fill:{color:C.navy},line:{color:C.navy}});
s.addText("上三层 · 规则与应用（管起来、用出去）",{x:rx+0.1,y:3.53,w:rw-0.2,h:0.3,fontSize:10.5,bold:true,color:C.white,fontFace:"Georgia"});

const topItems=[
  {t:"L5 指标中心",d:"原子指标池+映射中心+分摊引擎+版本管理+血缘追踪"},
  {t:"L6 数据服务",d:"统一API网关+指标服务+报表服务+行级权限控制"},
  {t:"L7 应用输出",d:"3套报表（法定/阿米巴/BG）+ 2个AI应用"},
];
topItems.forEach((b,i)=>{
  const y=3.9+i*0.42;
  s.addText(b.t,{x:rx+0.1,y,w:1.4,h:0.3,fontSize:9.5,bold:true,color:C.navy,fontFace:"Georgia"});
  s.addText(b.d,{x:rx+1.5,y,w:rw-1.6,h:0.35,fontSize:8.5,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.2});
});

s.render()}

// ===== SLIDE 5: 指标中心四层体系 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"METRIC CENTER","指标中心：四层指标体系");
const th={fill:{color:C.navy},color:C.white,bold:true,fontSize:10,fontFace:"Arial",align:"center"};
const td={fontSize:9.5,fontFace:"Arial",color:C.text,fill:{color:"FFFFFF"}};
const td2={fontSize:9.5,fontFace:"Arial",color:C.text,fill:{color:C.offW}};
const rows=[
  [{text:"层级",options:th},{text:"定位",options:th},{text:"示例",options:th},{text:"核心特征",options:th}],
  [{text:"L4 分析层（最上）",options:{...td,fill:{color:"FCE4EC"}}},{text:"经营决策用的分析指标",options:{...td,fill:{color:"FCE4EC"}}},{text:"毛利率、人效、回款率、单客价值、库存周转天数",options:{...td,fill:{color:"FCE4EC"}}},{text:"由下层组合计算、口径灵活",options:{...td,fill:{color:"FCE4EC"}}}],
  [{text:"L3 报表层",options:{...td2,fill:{color:"EBF5FF"}}},{text:"监管/管理报表的格式化项目",options:{...td2,fill:{color:"EBF5FF"}}},{text:"利润表项目、资产负债表项目、管理报表项目",options:{...td2,fill:{color:"EBF5FF"}}},{text:"按模板聚合、有合并抵消、重分类",options:{...td2,fill:{color:"EBF5FF"}}}],
  [{text:"L2 财务核算层",options:{...td,fill:{color:"E6F4EA"}}},{text:"按会计准则确认的收入/成本/费用",options:{...td,fill:{color:"E6F4EA"}}},{text:"营业收入、营业成本、期间费用、应收账款",options:{...td,fill:{color:"E6F4EA"}}},{text:"口径严格、受准则约束",options:{...td,fill:{color:"E6F4EA"}}}],
  [{text:"L1 业务运营层（最下）",options:{...td2,fill:{color:"FFF7E6"}}},{text:"业务发生的原始事实",options:{...td2,fill:{color:"FFF7E6"}}},{text:"订单量、发货量、签约金额、客户拜访次数",options:{...td2,fill:{color:"FFF7E6"}}},{text:"粒度最细、实时性高、不做财务加工",options:{...td2,fill:{color:"FFF7E6"}}}],
];
s.addTable(rows,{x:CX,y:1.3,w:CW,h:2.2,colW:[CW*0.18,CW*0.22,CW*0.35,CW*0.25],rowH:[0.35,0.5,0.5,0.5,0.5],border:{pt:1,color:C.grayL}});

// 三种模式
s.addText("三种层间关系模式",{x:CX,y:3.65,w:CW,h:0.3,fontSize:14,bold:true,color:C.navy,fontFace:"Georgia"});
const modes=[
  {t:"一对一映射",d:"L1→L2，有明确转换规则",e:"订单金额（L1）→【剔除退款+按履约进度+价税分离】→营业收入（L2）"},
  {t:"一对多拆解",d:"L1→多个L2，按维度拆分",e:"市场部总费用100万（L1）→【按项目工时分摊】→A产品线40万+B产品线35万+总部25万"},
  {t:"多对一聚合",d:"多下层→L4，组合计算",e:"单客贡献毛利（L4）= 营业收入（L2）÷活跃客户数（L1）- 营业成本（L2）÷活跃客户数（L1）"},
];
const mw=(CW-0.2)/3;
modes.forEach((m,i)=>{
  const x=CX+i*(mw+0.1);
  s.addShape(pres.shapes.RECTANGLE,{x,y:4.0,w:mw,h:1.35,fill:{color:"FFFFFF"},line:{color:C.grayL}});
  s.addText(m.t,{x:x+0.1,y:4.1,w:mw-0.2,h:0.28,fontSize:12,bold:true,color:C.navy,fontFace:"Georgia"});
  s.addText(m.d,{x:x+0.1,y:4.38,w:mw-0.2,h:0.2,fontSize:9.5,color:C.grayD,fontFace:"Arial"});
  s.addText(m.e,{x:x+0.1,y:4.6,w:mw-0.2,h:0.7,fontSize:9.5,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.3});
});
s.render()}

// ===== SLIDE 8: 核心能力模块 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"CORE MODULES","三大核心能力模块");

// 4.1 映射中心
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.35,w:CW,h:1.05,fill:{color:"FFFFFF"},line:{color:C.grayL}});
s.addShape(pres.shapes.RECTANGLE,{x:CX,y:1.35,w:CW,h:0.04,fill:{color:C.navy},line:{color:C.navy}});
s.addText("4.1 映射中心：解决维度差异",{x:CX+0.15,y:1.42,w:CW-0.3,h:0.3,fontSize:14,bold:true,color:C.navy,fontFace:"Georgia"});
s.addText("• 人 → 岗位 → 部门 → 法人公司 → 阿米巴单元 → BG事业群",{x:CX+0.15,y:1.75,w:CW-0.3,h:0.22,fontSize:11,color:C.text,fontFace:"Arial"});
s.addText("• 每个关系都有版本：张三2025年1-6月归东区巴，7月起归西区巴（历史可追溯）",{x:CX+0.15,y:1.97,w:CW-0.3,h:0.22,fontSize:11,color:C.text,fontFace:"Arial"});
s.addText("• 变更走审批，历史版本永久保留",{x:CX+0.15,y:2.19,w:CW-0.3,h:0.22,fontSize:11,color:C.text,fontFace:"Arial"});

// 4.2 分摊引擎（表格）
s.addText("4.2 分摊引擎：解决费用分配（5种类型）",{x:CX,y:2.55,w:CW,h:0.3,fontSize:14,bold:true,color:C.navy,fontFace:"Georgia"});
const th2={fill:{color:C.navy},color:C.white,bold:true,fontSize:9,fontFace:"Arial",align:"center"};
const allocRows=[
  [{text:"类型",options:th2},{text:"适用场景",options:th2},{text:"示例",options:th2}],
  [{text:"固定比例分摊",options:td},{text:"总部职能费按预定比例",options:td},{text:"HR费用：A巴30人摊30%，B巴70人摊70%",options:td}],
  [{text:"动因分摊",options:td2},{text:"IT费按实际使用量",options:td2},{text:"IT运维费：A巴1000次API调用摊40%，B巴1500次摊60%",options:td2}],
  [{text:"阶梯分摊",options:td},{text:"收入越大分摊比例越低",options:td},{text:"总部管理费：<1000万分摊5%，1000-5000万3%，>5000万1%",options:td}],
  [{text:"直接归属+剩余分摊",options:td2},{text:"能直接归的先归，剩下再分",options:td2},{text:"市场部费用：A项目的60万直接归A巴，剩余40万按收入比例分",options:td2}],
  [{text:"工时/面积分摊",options:td},{text:"房租按面积、共享人力按工时",options:td},{text:"办公房租：A巴300㎡摊30%，B巴700㎡摊70%",options:td}],
];
s.addTable(allocRows,{x:CX,y:2.88,w:CW,h:2.4,colW:[CW*0.18,CW*0.28,CW*0.54],rowH:[0.3,0.32,0.32,0.32,0.32,0.32],border:{pt:1,color:C.grayL}});
s.render()}

// ===== SLIDE 9: 三套报表差异化 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"REPORT DIFFERENTIATION","三套报表差异化支撑");
const th3={fill:{color:C.navy},color:C.white,bold:true,fontSize:10,fontFace:"Arial",align:"center"};
const rptRows=[
  [{text:"维度",options:th3},{text:"法定财报",options:th3},{text:"阿米巴管报",options:th3},{text:"BG经营报表",options:th3}],
  [{text:"给谁看",options:td},{text:"CFO、审计、投资者",options:td},{text:"各巴负责人、财务BP",options:td},{text:"BG总裁、运营总监",options:td}],
  [{text:"组织维度",options:td2},{text:"法人公司",options:td2},{text:"虚拟阿米巴单元（可跨法人）",options:td2},{text:"事业群/事业部",options:td2}],
  [{text:"收入确认",options:td},{text:"会计准则（权责发生制）",options:td},{text:"内部交易价+外部收入拆分",options:td},{text:"BG自定义确认节点",options:td}],
  [{text:"更新频率",options:td2},{text:"月/季/年",options:td2},{text:"周/月",options:td2},{text:"月/季",options:td2}],
  [{text:"规则变更",options:td},{text:"极少（受准则约束）",options:td},{text:"频繁（每季度可能调）",options:td},{text:"中等（随组织架构）",options:td}],
  [{text:"内部交易",options:td2},{text:"合并抵消（消掉）",options:td2},{text:"模拟内部定价，各巴都算损益",options:td2},{text:"视BG需要处理",options:td2}],
];
s.addTable(rptRows,{x:CX,y:1.35,w:CW,h:3.2,colW:[CW*0.15,CW*0.28,CW*0.3,CW*0.27],rowH:[0.38,0.38,0.38,0.38,0.38,0.38,0.38],border:{pt:1,color:C.grayL}});
s.render()}

// ===== SLIDE 10: AI应用场景 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"AI SCENARIOS","AI应用场景（按优先级）");
const th4={fill:{color:C.navy},color:C.white,bold:true,fontSize:10,fontFace:"Arial",align:"center"};
const aiRows=[
  [{text:"优先级",options:th4},{text:"场景",options:th4},{text:"说明",options:th4},{text:"示例",options:th4}],
  [{text:"P0",options:{...td,color:C.red,bold:true}},{text:"智能指标口径助手",options:td},{text:"输入指标名，查到定义、公式、来源、责任部门",options:td},{text:"问：\"阿米巴收入和法定收入有什么不一样？\" → 自动对比",options:td}],
  [{text:"P0",options:{...td2,color:C.red,bold:true}},{text:"智能数据质量监控",options:td2},{text:"自动检测数据异常（突增/突降/缺失）",options:td2},{text:"交通费比上月涨300% → 自动告警+定位源系统录错",options:td2}],
  [{text:"P1",options:td},{text:"智能血缘与根因分析",options:td},{text:"指标异常时自动沿血缘链追原因",options:td},{text:"毛利率降了 → 自动追溯：哪个产品线？收入降还是成本涨？",options:td}],
  [{text:"P1",options:td2},{text:"规则变更影响评估",options:td2},{text:"改规则前模拟：哪些指标会变、变多少",options:td2},{text:"IT费从\"按人数\"改\"按使用量\" → 模拟A巴增15万B巴减15万",options:td2}],
  [{text:"P1",options:td},{text:"自然语言取数",options:td},{text:"用中文问问题，AI自动查数据出表",options:td},{text:"\"上个月EBG前5大客户收入排行\" → 自动出表",options:td}],
  [{text:"P2",options:td2},{text:"经营分析自动生成",options:td2},{text:"每月自动写经营分析初稿",options:td2},{text:"自动生成：\"本月收入同比增15%，主要来自SBG增长23%...\"",options:td2}],
];
s.addTable(aiRows,{x:CX,y:1.35,w:CW,h:3.8,colW:[CW*0.07,CW*0.18,CW*0.3,CW*0.45],rowH:[0.38,0.45,0.45,0.45,0.45,0.45,0.45],border:{pt:1,color:C.grayL}});
s.addText("⚠️ 风控：AI是辅助不是替代。高敏感数据不开放AI查询，AI结果要和关键指标交叉验证。",{x:CX,y:5.1,w:CW,h:0.3,fontSize:10,bold:true,color:C.red,fontFace:"Arial"});
s.render()}

// ===== SLIDE 11: 实施路径 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"ROADMAP","实施路径：三期，12个月");
const phases=[
  {n:"第一期",t:"0-4个月",sub:"数据底座 + 法定财报跑通",color:C.red,items:["接入ERP、总账、HR等核心系统数据","搭建ODS→DWD→DWS→ADS四层数仓，统一主数据编码","迁移法定财报到平台，证明平台能力","建指标中台基础框架（原子指标池、规则引擎雏形）","AI落地：智能数据质量监控试点"]},
  {n:"第二期",t:"4-8个月",sub:"阿米巴管报 + 映射中心",color:C.amber,items:["建好映射中心（人-岗-巴-法人-BG关系）","迁移阿米巴核心逻辑（收入拆分、成本分摊）","上线口径版本管理和模拟测算","AI落地：智能指标口径助手、规则变更影响评估"]},
  {n:"第三期",t:"8-12个月",sub:"BG经营报表 + 全面智能化",color:C.green,items:["迁移BG经营报表，支持BG自定义口径","建统一报表门户和数据API服务层","试点开放自然语言取数","全面推广AI辅助的质量监控和根因分析"]},
];
const pw=(CW-0.2)/3;
phases.forEach((ph,i)=>{
  const x=CX+i*(pw+0.1);
  s.addShape(pres.shapes.RECTANGLE,{x,y:1.35,w:pw,h:3.8,fill:{color:"FFFFFF"},line:{color:C.grayL}});
  s.addShape(pres.shapes.RECTANGLE,{x,y:1.35,w:pw,h:0.6,fill:{color:ph.color},line:{color:ph.color}});
  s.addText(ph.n,{x,y:1.4,w:pw,h:0.28,fontSize:14,bold:true,color:C.white,align:"center",fontFace:"Georgia"});
  s.addText(ph.t,{x,y:1.67,w:pw,h:0.22,fontSize:11,color:C.white,align:"center",fontFace:"Arial"});
  s.addText(ph.sub,{x:x+0.1,y:2.05,w:pw-0.2,h:0.3,fontSize:11,bold:true,color:ph.color,align:"center",fontFace:"Arial"});
  ph.items.forEach((it,j)=>{
    s.addText("• "+it,{x:x+0.12,y:2.42+j*0.3,w:pw-0.24,h:0.28,fontSize:10,color:C.text,fontFace:"Arial",lineSpacingMultiple:1.25});
  });
  // Arrow between phases - removed to avoid overlap
});
s.render()}

// ===== SLIDE 12: 关键成功因素 + 结束 =====
{const s=pres.addSlide();s.background={color:C.offW};
addHeader(s,"KEY SUCCESS FACTORS","5个关键成功因素");
const kfs=[
  {i:"①",t:"高层背书",d:"CFO/COO牵头，这是\"一把手工程\""},
  {i:"②",t:"先跑通一套",d:"先让法定财报稳定运行，再迁阿米巴和BG"},
  {i:"③",t:"指标中台是灵魂",d:"投入资源打磨规则引擎"},
  {i:"④",t:"业务人员参与",d:"口径是财务BP定的，不是IT定的"},
  {i:"⑤",t:"AI适度引入",d:"AI提效，不替代人做判断"},
];
const kw=(CW-0.4)/5;
kfs.forEach((k,i)=>{
  const x=CX+i*(kw+0.1);
  s.addShape(pres.shapes.RECTANGLE,{x,y:1.5,w:kw,h:3.2,fill:{color:"FFFFFF"},line:{color:C.grayL}});
  s.addText(k.i,{x,y:1.65,w:kw,h:0.8,fontSize:36,bold:true,color:C.navy,align:"center",fontFace:"Georgia"});
  s.addText(k.t,{x:x+0.1,y:2.5,w:kw-0.2,h:0.35,fontSize:13,bold:true,color:C.text,align:"center",fontFace:"Georgia"});
  s.addText(k.d,{x:x+0.1,y:2.9,w:kw-0.2,h:1.5,fontSize:10.5,color:C.grayD,align:"center",fontFace:"Arial",lineSpacingMultiple:1.4});
});
s.render()}

// ===== SLIDE 13: CLOSING =====
{const s=pres.addSlide();s.background={color:C.navyDk};
s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:0.12,fill:{color:C.blue},line:{color:C.blue}});
s.addShape(pres.shapes.RECTANGLE,{x:0,y:SH-0.12,w:SW,h:0.12,fill:{color:C.blue},line:{color:C.blue}});
s.addText("谢谢",{x:0,y:1.8,w:SW,h:1.4,fontSize:56,bold:true,color:C.white,align:"center",fontFace:"Georgia",charSpacing:3});
s.addText("一个数据底座 · 七层架构 · 支撑三套报表\n\n数据只取一次，上层各算各的",{x:CX+1,y:3.5,w:CW-2,h:1.2,fontSize:16,color:C.grayM,align:"center",fontFace:"Arial",lineSpacingMultiple:1.6});
s.render()}

// ===== SAVE =====
const out="C:\\Users\\ruijie\\AppData\\Roaming\\TRAE SOLO CN\\ModularData\\ai-agent\\work-mode-projects\\6a4f8645e4b01f7722a32ca4\\统一数据平台与指标中心建设方案_v2.pptx";
pres.writeFile({fileName:out}).then(()=>console.log("Saved: "+out)).catch(e=>console.error(e));
