/**
 * word_builder.js — 用 docx 库生成 Word 报告
 * 调用方式：node word_builder.js <payload.json路径>
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat, Header, Footer, SimpleField
} = require('docx');
const fs = require('fs');

// ── 读取 payload ──
const payload = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const { outPath, company, year, base, rptDate, rptTarget, ch1, ch2, ch3, ch4, ch5 } = payload;

// ── 颜色常量 ──
const C = {
  darkBlue: '1F3864', midBlue: '2E75B6', lightBlue: 'BDD7EE',
  darkGray: '404040', midGray: '808080', white: 'FFFFFF',
  red: 'C00000', orange: 'E36C0A', yellow: 'BF8F00',
  green: '375623', gold: 'C9A000',
};

// ── 工具函数 ──
const W = 9026; // A4内容宽度 DXA
const b = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
const borders = { top: b, bottom: b, left: b, right: b };

const fmt = (v, d=0) => {
  v = parseFloat(v) || 0;
  if (Math.abs(v) >= 1e8) return `${(v/1e8).toFixed(d)}亿`;
  return `${(v/1e4).toFixed(d)}万`;
};
const pct = (v, d=1) => `${((parseFloat(v)||0)*100).toFixed(d)}%`;
const ppt = (v, d=2) => {
  v = (parseFloat(v)||0)*100;
  return (v>=0?'+':'')+v.toFixed(d)+'ppt';
};
const chg = (v, d=1) => {
  v = parseFloat(v)||0;
  return (v>=0?'+':'')+v.toFixed(d)+'%';
};

// ── 段落构建 ──
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 160 },
    children: [new TextRun({ text, font: 'Arial', size: 36, bold: true, color: C.darkBlue })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 1 } },
    children: [new TextRun({ text, font: 'Arial', size: 28, bold: true, color: C.midBlue })]
  });
}
function p(text, opts={}) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 340 },
    children: [new TextRun({
      text: String(text), font: 'Arial', size: opts.size||21,
      bold: opts.bold||false, color: opts.color||C.darkGray, italics: opts.italic||false
    })]
  });
}
function sp(n=100) { return new Paragraph({ spacing: { before: n, after: 0 }, children: [new TextRun('')] }); }
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function nl(text) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: 0 },
    spacing: { before: 40, after: 40, line: 300 },
    children: [new TextRun({ text, font: 'Arial', size: 21, color: C.darkGray })]
  });
}
function divider() {
  return new Paragraph({
    spacing: { before: 100, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.lightBlue, space: 1 } },
    children: [new TextRun('')]
  });
}

// ── 单元格构建 ──
function cell(text, opts={}) {
  const fill = opts.fill || C.white;
  return new TableCell({
    borders, verticalAlign: VerticalAlign.CENTER,
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    columnSpan: opts.span || undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 0, after: 0 },
      children: [new TextRun({
        text: String(text ?? '—'), font: 'Arial', size: opts.size||20,
        bold: opts.bold||false, color: opts.color||C.darkGray
      })]
    })]
  });
}
function hCell(text, w, align=AlignmentType.CENTER) {
  return cell(text, { fill: C.darkBlue, bold: true, color: C.white, w, align, size: 19 });
}
function shCell(text, w) {
  return cell(text, { fill: C.midBlue, bold: true, color: C.white, w, align: AlignmentType.LEFT, size: 19 });
}

// ── KPI 双列表格 ──
function kpiTable(items) {
  const cw = Math.floor(W / 2);
  const rows = [];
  for (let i = 0; i < items.length; i += 2) {
    const mk = (item) => {
      if (!item) return new TableCell({ borders, width: { size: cw, type: WidthType.DXA }, shading: { fill: C.white, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [] })] });
      return new TableCell({
        borders, width: { size: cw, type: WidthType.DXA },
        shading: { fill: 'F0F4FA', type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: item.label, font: 'Arial', size: 18, color: C.midGray || '808080' })] }),
          new Paragraph({ spacing: { before: 0, after: 10 }, children: [new TextRun({ text: item.value, font: 'Arial', size: 28, bold: true, color: item.color || C.darkBlue })] }),
          item.sub ? new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: item.sub, font: 'Arial', size: 17, color: C.midGray || '808080' })] }) : new Paragraph({ children: [] }),
        ]
      });
    };
    rows.push(new TableRow({ children: [mk(items[i]), mk(items[i+1])] }));
  }
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [cw, cw], rows });
}

// ── 提示框 ──
function alertBox(title, body, style='blue') {
  const styles = {
    red:    { fill: 'FFCCCC', tc: C.red },
    yellow: { fill: 'FFF2CC', tc: C.yellow },
    green:  { fill: 'E2EFDA', tc: C.green },
    blue:   { fill: 'DEEAF1', tc: C.midBlue },
  };
  const s = styles[style] || styles.blue;
  const lines = body.split('\n');
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      width: { size: W, type: WidthType.DXA },
      shading: { fill: s.fill, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 100, left: 180, right: 180 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: title, font: 'Arial', size: 22, bold: true, color: s.tc })] }),
        ...lines.map(line => new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: line, font: 'Arial', size: 20, color: C.darkGray })] }))
      ]
    })] })]
  });
}

// ════════════════════════════════════
// 正文构建
// ════════════════════════════════════
const children = [];

// ── 封面 ──
children.push(
  sp(2400),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 140 }, children: [new TextRun({ text: company, font: 'Arial', size: 26, color: C.midGray })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: `${year}度综合财务分析报告`, font: 'Arial', size: 52, bold: true, color: C.darkBlue })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.midBlue, space: 1 } }, spacing: { before: 0, after: 280 }, children: [new TextRun({ text: `供${rptTarget}决策参考`, font: 'Arial', size: 28, color: C.midGray, italics: true })] }),
  sp(360),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `报告日期：${rptDate}　|　分析范围：${base}–${year}`, font: 'Arial', size: 22, color: C.midGray })] }),
  sp(160),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '本报告含敏感经营信息，仅限内部使用，禁止对外传阅', font: 'Arial', size: 18, color: C.midGray, italics: true })] }),
  pb()
);

// ── 执行摘要 ──
const rev   = parseFloat(ch1.rev_25)||0;
const da    = parseFloat(ch1.da_ratio_25)||0;
const gmChg = (parseFloat(ch1.gm_rate_25)||0) - (parseFloat(ch1.gm_rate_24)||0);
const adjNet = parseFloat(ch1.adj_net_25)||0;
const net   = parseFloat(ch1.net_25)||0;
const adjG  = net>0 ? (adjNet-(parseFloat(ch1.net_24)||0))/(parseFloat(ch1.net_24)||1)*100 : 0;
const oGain = parseFloat(ch1.onetime_gain)||0;

children.push(
  h1('执行摘要'),
  p('基于利润表、资产负债表、现金流量表、收入成本明细、应收/应付/预收账款底稿及项目级现金流数据，对年度经营状况进行全面穿透分析。'),
  sp(80),
  p('五个核心发现', { bold: true, size: 24, color: C.midBlue }),
  nl(`净利润账面增长存在水分——含一次性收益约${fmt(oGain)}，剔除后实质净利润${fmt(adjNet)}，实质增幅仅${adjG.toFixed(1)}%`),
  nl(`收入增长依赖新项目输血，存续老项目同店收入持续萎缩，"破桶效应"明显`),
  nl(`应收账款持续积累，净资金敞口扩大，是制约并购扩张的首要瓶颈`),
  nl(`财务结构${da<0.55?'已优化，':'仍有压力，'}资产负债率${pct(da)}，${da<0.55?'首次达到战略目标':'需继续降杠杆'}`),
  nl(`关键区域同时触发毛利和应收双重预警，需立即专项介入`),
  sp(80),
  alertBox(
    `${year}核心战略指令`,
    `公司财务结构${da<0.6?'已安全':'仍有压力'}，但经营质量需要持续提升。\n${year}的首要任务：压缩应收账款 → 释放现金 → 修复问题区域盈利能力 → 再以更健康的基本面推进扩张`,
    'yellow'
  ),
  pb()
);

// ── 第一章 ──
const sell24=parseFloat(ch1.sell_24)||0, sell25=parseFloat(ch1.sell_25)||0;
const mgmt24=parseFloat(ch1.mgmt_24)||0, mgmt25=parseFloat(ch1.mgmt_25)||0;
const fin24 =parseFloat(ch1.fin_24)||0,  fin25 =parseFloat(ch1.fin_25)||0;
const rev24 =parseFloat(ch1.rev_24)||1,  rev25_v=parseFloat(ch1.rev_25)||1;
const gm24  =parseFloat(ch1.gm_rate_24)||0, gm25=parseFloat(ch1.gm_rate_25)||0;

children.push(
  h1('第一章  公司整体经营健康度'),
  p('核心问题：公司整体是在变好还是变差？账面数字和真实经营能力之间是否存在差距？'),
  sp(80),
  h2('一、规模与增长'),
  kpiTable([
    { label: `营业收入（${year}）`, value: fmt(rev, 0), sub: `较${base}增长${((rev-rev24)/rev24*100).toFixed(1)}%`, color: C.midBlue },
    { label: '综合毛利率', value: pct(gm25), sub: `变动 ${ppt(gmChg)}${gmChg<0?' ↓':' ↑'}`, color: gmChg<0?C.red:C.green },
    { label: '资产负债率', value: pct(da), sub: da<0.55?'✓ 达到战略目标':'⚠ 仍有压力', color: da<0.55?C.green:C.orange },
    { label: 'OCF/实质净利润', value: pct(ch1.ocf_quality), sub: 'FCF '+fmt(ch1.fcf), color: (parseFloat(ch1.ocf_quality)||0)>1?C.green:C.orange },
  ]),
  sp(80),
  h2('二、完整利润表与三项费用率'),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3400, 1200, 1200, 1100, 2126],
    rows: [
      new TableRow({ children: [hCell('利润层级',3400), hCell(`${base}(万)`,1200,AlignmentType.RIGHT), hCell(`${base}占收入%`,1200,AlignmentType.RIGHT), hCell(`${year}(万)`,1100,AlignmentType.RIGHT), hCell(`${year}占收入%/变动`,2126,AlignmentType.RIGHT)] }),
      new TableRow({ children: [cell('营业收入',{fill:'DEEAF1',bold:true,w:3400}), cell(fmt(rev24),{align:AlignmentType.RIGHT,w:1200,bold:true}), cell('100%',{align:AlignmentType.RIGHT,w:1200}), cell(fmt(rev),{align:AlignmentType.RIGHT,w:1100,bold:true}), cell('100%',{align:AlignmentType.RIGHT,w:2126})] }),
      new TableRow({ children: [cell('减：营业成本',{w:3400}), cell(fmt(ch1.cst_24),{align:AlignmentType.RIGHT,w:1200}), cell(pct((parseFloat(ch1.cst_24)||0)/rev24),{align:AlignmentType.RIGHT,w:1200}), cell(fmt(ch1.cst_25),{align:AlignmentType.RIGHT,w:1100}), cell(pct((parseFloat(ch1.cst_25)||0)/rev25_v)+' '+ppt((parseFloat(ch1.cst_25)||0)/rev25_v-(parseFloat(ch1.cst_24)||0)/rev24),{align:AlignmentType.RIGHT,w:2126,color:C.red})] }),
      new TableRow({ children: [cell('毛利额',{fill:'E2EFDA',bold:true,w:3400}), cell(fmt(ch1.gm_24),{align:AlignmentType.RIGHT,w:1200,color:C.green,bold:true}), cell(pct(gm24),{align:AlignmentType.RIGHT,w:1200,color:C.green}), cell(fmt(ch1.gm_25),{align:AlignmentType.RIGHT,w:1100,color:gmChg<0?C.orange:C.green,bold:true}), cell(pct(gm25)+' '+ppt(gmChg),{align:AlignmentType.RIGHT,w:2126,color:gmChg<0?C.red:C.green,bold:true})] }),
      new TableRow({ children: [cell(`减：管理费用${mgmt25<mgmt24?' ★节省'+fmt(mgmt24-mgmt25):''}`,{w:3400}), cell(fmt(mgmt24),{align:AlignmentType.RIGHT,w:1200}), cell(pct(mgmt24/rev24),{align:AlignmentType.RIGHT,w:1200}), cell(fmt(mgmt25),{align:AlignmentType.RIGHT,w:1100,color:mgmt25<mgmt24?C.green:''}), cell(pct(mgmt25/rev25_v)+' '+ppt(mgmt25/rev25_v-mgmt24/rev24),{align:AlignmentType.RIGHT,w:2126,color:mgmt25<mgmt24?C.green:''})] }),
      new TableRow({ children: [cell(`减：财务费用${fin25<fin24?' ★节省'+fmt(fin24-fin25):''}`,{w:3400}), cell(fmt(fin24),{align:AlignmentType.RIGHT,w:1200}), cell(pct(fin24/rev24),{align:AlignmentType.RIGHT,w:1200}), cell(fmt(fin25),{align:AlignmentType.RIGHT,w:1100,color:fin25<fin24?C.green:''}), cell(pct(fin25/rev25_v)+' '+ppt(fin25/rev25_v-fin24/rev24),{align:AlignmentType.RIGHT,w:2126,color:fin25<fin24?C.green:''})] }),
      new TableRow({ children: [cell('加：其他收益（含一次性处置收益）',{fill:'FFF2CC',w:3400}), cell('—',{align:AlignmentType.RIGHT,w:1200}), cell('—',{align:AlignmentType.RIGHT,w:1200}), cell(fmt(ch1.other_income_25),{align:AlignmentType.RIGHT,w:1100,color:C.orange}), cell(`含一次性${fmt(oGain)}`,{align:AlignmentType.RIGHT,w:2126,color:C.orange})] }),
      new TableRow({ children: [cell('净利润（账面）',{fill:'DEEAF1',bold:true,w:3400}), cell(fmt(ch1.net_24),{align:AlignmentType.RIGHT,w:1200,bold:true}), cell('—',{align:AlignmentType.RIGHT,w:1200}), cell(fmt(net),{align:AlignmentType.RIGHT,w:1100,bold:true}), cell(`账面增幅${((net-(parseFloat(ch1.net_24)||0))/(parseFloat(ch1.net_24)||1)*100).toFixed(1)}%`,{align:AlignmentType.RIGHT,w:2126,color:C.midBlue,bold:true})] }),
      new TableRow({ children: [cell('净利润（剔除一次性收益后）',{fill:'FFF2CC',bold:true,w:3400}), cell(fmt(ch1.net_24),{align:AlignmentType.RIGHT,w:1200,bold:true}), cell('—',{align:AlignmentType.RIGHT,w:1200}), cell(fmt(adjNet),{align:AlignmentType.RIGHT,w:1100,bold:true,color:C.orange}), cell(`实质增幅${adjG.toFixed(1)}%  ⚠`,{align:AlignmentType.RIGHT,w:2126,color:C.orange,bold:true})] }),
    ]
  }),
  sp(80),
  h2('三、财务安全性'),
  kpiTable([
    { label: '资产负债率', value: pct(da), sub: da<0.55?'✓ 达到战略目标':'⚠ 仍有压力', color: da<0.55?C.green:C.orange },
    { label: '利息保障倍数', value: `${Math.min(parseFloat(ch1.interest_cover)||0,99).toFixed(0)}x`, sub: '警戒线>3x', color: (parseFloat(ch1.interest_cover)||0)>3?C.green:C.red },
    { label: '有息负债/EBITDA', value: `${(parseFloat(ch1.debt_ebitda)||0).toFixed(2)}x`, sub: '安全线<3x', color: (parseFloat(ch1.debt_ebitda)||0)<3?C.green:C.red },
    { label: '商誉（潜在风险）', value: fmt(ch1.goodwill_25), sub: `占总资产${pct((parseFloat(ch1.goodwill_25)||0)/Math.max(parseFloat(ch1.total_asset_25)||1,1))}`, color: C.orange },
  ]),
  pb()
);

// ── 第二章 ──
const regions = ch2.region_analysis || {};
const regionsSorted = Object.entries(regions).sort((a,b)=>b[1].gm_rate_25-a[1].gm_rate_25);

const regionRows = regionsSorted.map(([rgn, v]) => {
  const gm25r = parseFloat(v.gm_rate_25)||0;
  const pptV  = parseFloat(v.gm_ppt)||0;
  const fill  = gm25r>0.20 ? 'E2EFDA' : (gm25r<0.05||pptV<-0.05 ? 'FFCCCC' : pptV<-0.02 ? 'FFF2CC' : C.white);
  const pptC  = pptV>=0 ? C.green : pptV<-0.05 ? C.red : C.orange;
  return new TableRow({ children: [
    cell(rgn,{fill,w:3200}),
    cell(fmt(v.rev_25),{align:AlignmentType.RIGHT,w:1100,fill}),
    cell(pct(v.gm_rate_24),{align:AlignmentType.RIGHT,w:1100,fill}),
    cell(pct(gm25r),{align:AlignmentType.RIGHT,w:1100,fill,bold:true}),
    cell(ppt(pptV),{align:AlignmentType.RIGHT,w:1200,fill,color:pptC,bold:true}),
    cell(fmt(v.gm_chg),{align:AlignmentType.RIGHT,w:1326,fill}),
  ]});
});

children.push(
  h1('第二章  利润结构解剖：谁在赚钱，谁在流血'),
  p('核心问题：总利润由哪些业务贡献？哪些板块是拖累？毛利率下滑的根本原因是什么？'),
  sp(80),
  h2('一、各区域综合毛利率排行'),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3200, 1100, 1100, 1100, 1200, 1326],
    rows: [
      new TableRow({ children: [hCell('区域',3200), hCell(`${year}收入(万)`,1100,AlignmentType.RIGHT), hCell(`${base}毛利率`,1100,AlignmentType.RIGHT), hCell(`${year}毛利率`,1100,AlignmentType.RIGHT), hCell('变动(ppt)',1200,AlignmentType.RIGHT), hCell('毛利额变动(万)',1326,AlignmentType.RIGHT)] }),
      ...regionRows
    ]
  }),
  sp(80),
  h2('二、项目毛利率区间分布'),
  (() => {
    const dist = ch2.proj_dist || [];
    const total = dist.reduce((s,d)=>s+(d.count||0),0);
    const distRows = dist.map(d => {
      const fills = {'亏损(<0%)':'FFCCCC','微利(0-5%)':'FCE4D6','正常(5-10%)':C.white,'良好(10-20%)':'E2EFDA','优质(>20%)':'C6E0B4'};
      const fill = fills[d.band] || C.white;
      return new TableRow({ children: [
        cell(String(d.band||''),{fill,w:2500}),
        cell(String(d.count||0),{align:AlignmentType.RIGHT,w:1000,fill,bold:true}),
        cell(`${((d.count||0)/Math.max(total,1)*100).toFixed(1)}%`,{align:AlignmentType.RIGHT,w:1000,fill}),
        cell(fmt(d.rev_sum||0),{align:AlignmentType.RIGHT,w:1200,fill}),
        cell('',{w:3326,fill}),
      ]});
    });
    return new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [2500,1000,1000,1200,3326],
      rows: [
        new TableRow({ children: [hCell('毛利率区间',2500), hCell('项目数',1000,AlignmentType.RIGHT), hCell('占比',1000,AlignmentType.RIGHT), hCell('收入合计(万)',1200,AlignmentType.RIGHT), hCell('说明',3326)] }),
        ...distRows
      ]
    });
  })(),
  pb()
);

// ── 第三章 ──
const regionDf = ch3.region_df || {};
const arTrend  = ch3.ar_trend  || {};
const rpAr25   = parseFloat(ch3.rp_ar_25)||0;
const rpAr24   = parseFloat(ch3.rp_ar_24)||0;

const arRows = Object.entries(regionDf)
  .sort((a,b)=>b[1].gap-a[1].gap)
  .map(([rgn, v]) => {
    const gap  = parseFloat(v.gap)||0;
    const days = parseFloat(v.ar_days)||0;
    const fill = gap>1e7?'FFCCCC':gap>5e6?'FCE4D6':gap>2e6?'FFF2CC':gap<0?'E2EFDA':C.white;
    const gc   = gap>5e6?C.red:gap>2e6?C.orange:gap<0?C.green:C.darkGray;
    const dc   = days>365?C.red:days>180?C.orange:days<90?C.green:C.darkGray;
    return new TableRow({ children: [
      cell(rgn,{fill,w:2800}),
      cell(fmt(v.ar_25||0),{align:AlignmentType.RIGHT,w:950,fill}),
      cell(fmt(v.ap_25||0),{align:AlignmentType.RIGHT,w:950,fill}),
      cell(fmt(v.adv_25||0),{align:AlignmentType.RIGHT,w:950,fill}),
      cell(fmt(gap),{align:AlignmentType.RIGHT,w:1000,fill,color:gc,bold:true}),
      cell(`${days.toFixed(0)}天`,{align:AlignmentType.RIGHT,w:1000,fill,color:dc,bold:true}),
    ]});
  });

children.push(
  h1('第三章  应收·应付·预收：钱在哪里，谁在垫资'),
  p('核心公式：净资金敞口 = 期末应收 − 期末应付 − 期末预收'),
  sp(80),
  h2('一、应收账款历年趋势（万元）'),
  (() => {
    const yrs = Object.entries(arTrend).sort((a,b)=>a[0].localeCompare(b[0]));
    if (yrs.length < 2) return p('暂无历年趋势数据');
    return new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: Array(yrs.length).fill(Math.floor(W/yrs.length)),
      rows: [
        new TableRow({ children: yrs.map(([yr])=>hCell(yr, Math.floor(W/yrs.length))) }),
        new TableRow({ children: yrs.map(([yr,v])=>cell(fmt(v), { align:AlignmentType.RIGHT, color: yr===yrs[yrs.length-1][0]?C.red:C.orange, bold: yr===yrs[yrs.length-1][0] })) }),
      ]
    });
  })(),
  sp(60),
  alertBox('应收账款5年从未下降——结构性危机', '每年新增欠款远超回收金额，是商业模式中长期无法向甲方完全收款的结构性缺陷。全公司净资金敞口持续扩大，是制约并购扩张的首要瓶颈。', 'red'),
  sp(80),
  h2('二、各区域净资金敞口与应收周转天数'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [2800,950,950,950,1000,1000,376],
    rows: [
      new TableRow({ children: [hCell('区域',2800), hCell('AR(万)',950,AlignmentType.RIGHT), hCell('AP(万)',950,AlignmentType.RIGHT), hCell('预收(万)',950,AlignmentType.RIGHT), hCell('净敞口(万)',1000,AlignmentType.RIGHT), hCell('AR周转天',1000,AlignmentType.RIGHT), hCell('',376)] }),
      ...(rpAr25 > 0 ? [new TableRow({ children: [
        cell('关联方（单独列示）',{fill:'EAD1F5',w:2800}),
        cell(fmt(rpAr25),{align:AlignmentType.RIGHT,w:950,fill:'EAD1F5',color:'6A0DAD'}),
        cell('—',{align:AlignmentType.RIGHT,w:950,fill:'EAD1F5'}),
        cell('—',{align:AlignmentType.RIGHT,w:950,fill:'EAD1F5'}),
        cell('—',{align:AlignmentType.RIGHT,w:1000,fill:'EAD1F5'}),
        cell('—',{align:AlignmentType.RIGHT,w:1000,fill:'EAD1F5'}),
        cell('',{w:376,fill:'EAD1F5'}),
      ]})] : []),
      ...arRows
    ]
  }),
  sp(60),
  h2('三、应收账款周转天数（分业态）'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [3500,1500,1500,1400,2126],
    rows: [
      new TableRow({ children: [hCell('业务口径',3500), hCell('平均AR(万)',1500,AlignmentType.RIGHT), hCell(`${year}收入(万)`,1500,AlignmentType.RIGHT), hCell('周转天数',1400,AlignmentType.RIGHT), hCell('参考标准',2126)] }),
      new TableRow({ children: [cell('全公司整体',{fill:'DEEAF1',bold:true,w:3500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell(`${(parseFloat(ch3.overall_days)||0).toFixed(0)}天`,{align:AlignmentType.RIGHT,w:1400,bold:true,color:(parseFloat(ch3.overall_days)||0)>150?C.orange:C.green}), cell('—',{w:2126})] }),
      new TableRow({ children: [cell('居民生活服务（住宅）',{fill:'FFCCCC',w:3500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell(`${(parseFloat(ch3.residential_days)||0).toFixed(0)}天`,{align:AlignmentType.RIGHT,w:1400,bold:true,color:(parseFloat(ch3.residential_days)||0)>180?C.red:C.orange}), cell('<180天正常',{w:2126})] }),
      new TableRow({ children: [cell('企事业后勤综合服务',{fill:'E2EFDA',w:3500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell(`${(parseFloat(ch3.commercial_days)||0).toFixed(0)}天`,{align:AlignmentType.RIGHT,w:1400,bold:true,color:(parseFloat(ch3.commercial_days)||0)<90?C.green:C.orange}), cell('<90天优秀',{w:2126})] }),
      ...(rpAr25>0?[new TableRow({ children: [cell('关联方（单独列示）',{fill:'EAD1F5',w:3500}), cell(fmt(rpAr25),{align:AlignmentType.RIGHT,w:1500,color:'6A0DAD'}), cell('—',{align:AlignmentType.RIGHT,w:1500}), cell('—',{align:AlignmentType.RIGHT,w:1400}), cell('需单独管理',{w:2126})] })]:[]),
    ]
  }),
  pb()
);

// ── 第四章 ──
const regionCf  = ch4.region_cf  || {};
const regionMkt = ch4.region_cf_mkt || {};
const corr      = ch4.region_correction || {};

const cfRows = Object.entries(regionCf).sort((a,b)=>a[1].net-b[1].net).map(([rgn,v]) => {
  const net = parseFloat(v.net)||0;
  const fill = net<-500000?'FFCCCC':net<0?'FCE4D6':net>1000000?'E2EFDA':C.white;
  const nc = net<0?C.red:C.green;
  const mkt = regionMkt[rgn] ? parseFloat(regionMkt[rgn].net)||0 : net;
  const hasCorr = !!corr[rgn];
  return new TableRow({ children: [
    cell(rgn+(hasCorr?'　⚠':''),{fill,w:3200}),
    cell(fmt(parseFloat(v.in)||0),{align:AlignmentType.RIGHT,w:1400,fill}),
    cell(fmt(parseFloat(v.out)||0),{align:AlignmentType.RIGHT,w:1500,fill}),
    cell(fmt(net),{align:AlignmentType.RIGHT,w:1300,fill,color:nc,bold:true}),
    cell(hasCorr?`纯市场化${fmt(mkt)}`:'—',{align:AlignmentType.RIGHT,w:1626,fill,color:hasCorr?C.orange:C.darkGray,size:18}),
  ]});
});

const corrEntries = Object.entries(corr);
const corrNote = corrEntries.length > 0
  ? corrEntries.map(([rgn,v]) => `${rgn}：含关联方${fmt(v.net_rp)}（占比${v.rp_pct}%），剥离后纯市场化${fmt(v.net_mkt)}`).join('\n')
  : '本期无关联方重大影响';

children.push(
  h1('第四章  现金流与偿债能力'),
  p('核心问题：利润是真实的吗？各区域谁在造血、谁在消耗？关联方一次性回款如何影响区域排名？'),
  sp(80),
  h2('一、偿债能力'),
  kpiTable([
    { label: '有息负债/EBITDA', value: `${(parseFloat(ch4.debt_ebitda)||0).toFixed(2)}x`, sub: '安全线<3x', color: (parseFloat(ch4.debt_ebitda)||0)<3?C.green:C.red },
    { label: '利息保障倍数', value: `${Math.min(parseFloat(ch4.interest_cover)||0,99).toFixed(0)}x`, sub: '警戒线>3x', color: (parseFloat(ch4.interest_cover)||0)>3?C.green:C.red },
    { label: '资产负债率', value: pct(ch4.da_ratio_25), sub: '首次低于50% ✓', color: C.green },
    { label: '现金覆盖短期借款', value: `${(parseFloat(ch4.cash_cover)||0).toFixed(1)}x`, sub: `现金${fmt(ch4.cash_25)} vs 短贷${fmt(ch4.st_loan_25)}`, color: (parseFloat(ch4.cash_cover)||0)>1.5?C.green:C.orange },
  ]),
  sp(80),
  h2('二、利润含金量'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [4500, 2000, 2526],
    rows: [
      new TableRow({ children: [hCell('指标',4500), hCell('数值',2000,AlignmentType.RIGHT), hCell('解读',2526)] }),
      new TableRow({ children: [cell('账面净利润',{w:4500}), cell(fmt(ch4.net_25),{align:AlignmentType.RIGHT,w:2000}), cell('含一次性项目',{w:2526})] }),
      new TableRow({ children: [cell('减：一次性处置收益（估算·税后）',{fill:'FFCCCC',w:4500}), cell(`-${fmt(ch4.onetime_gain)}`,{align:AlignmentType.RIGHT,w:2000,color:C.red}), cell('不可持续',{w:2526,fill:'FFCCCC'})] }),
      new TableRow({ children: [cell('实质净利润（剔除后）',{fill:'FFF2CC',bold:true,w:4500}), cell(fmt(ch4.adj_net_25),{align:AlignmentType.RIGHT,w:2000,bold:true,color:C.orange}), cell('实际增速更低',{w:2526,fill:'FFF2CC',color:C.orange})] }),
      new TableRow({ children: [cell('OCF/实质净利润',{fill:'E2EFDA',bold:true,w:4500}), cell(pct(ch4.ocf_quality),{align:AlignmentType.RIGHT,w:2000,bold:true,color:C.green}), cell((parseFloat(ch4.ocf_quality)||0)>1?'✓ 利润质量高':'⚠ 利润质量偏低',{w:2526,fill:'E2EFDA',color:C.green})] }),
      new TableRow({ children: [cell('自由现金流（FCF）',{fill:'FFF2CC',bold:true,w:4500}), cell(fmt(ch4.fcf),{align:AlignmentType.RIGHT,w:2000,bold:true,color:C.orange}), cell((parseFloat(ch4.fcf)||0)<5e7?'⚠ 难支撑大型并购':'✓ 充足',{w:2526,fill:'FFF2CC',color:C.orange})] }),
    ]
  }),
  sp(80),
  h2('三、各区域经营现金净额（含关联方修正）'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [3200,1400,1500,1300,1626],
    rows: [
      new TableRow({ children: [hCell('区域',3200), hCell('销售收款(万)',1400,AlignmentType.RIGHT), hCell('人工+采购(万)',1500,AlignmentType.RIGHT), hCell('净额(万)',1300,AlignmentType.RIGHT), hCell('关联方修正',1626,AlignmentType.RIGHT)] }),
      ...cfRows
    ]
  }),
  sp(60),
  corrEntries.length > 0 ? alertBox('⚠ 关联方一次性回款修正说明', corrNote, 'yellow') : p(''),
  pb()
);

// ── 第五章 ──
const scores = ch5.scores || {};
const issues = ch5.issues || [];
const scoreColors = {A:C.green,'B+':C.green,B:'375623','B-':C.yellow,'C+':C.yellow,C:C.orange,'D+':C.red,D:C.red};

children.push(
  h1('第五章  核心问题与优先级行动建议'),
  sp(80),
  h2('一、经营健康度综合评分'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [3000,1200,1600,3226],
    rows: [
      new TableRow({ children: [hCell('评估维度',3000), hCell('评分',1200,AlignmentType.CENTER), hCell('趋势',1600,AlignmentType.CENTER), hCell('核心判断',3226)] }),
      ...Object.entries(scores).map(([dim,sc]) =>
        new TableRow({ children: [
          cell(dim,{w:3000}),
          cell(sc,{align:AlignmentType.CENTER,w:1200,bold:true,color:scoreColors[sc]||C.darkGray}),
          cell('—',{align:AlignmentType.CENTER,w:1600}),
          cell('',{w:3226}),
        ]})
      )
    ]
  }),
  sp(80),
  h2('二、优先级行动清单'),
  ...issues.map(iss => {
    const style = iss.level==='red'?'red':iss.level==='green'?'green':'yellow';
    return [alertBox(`${iss.priority} ${iss.title}`, iss.detail+'\n来源：'+iss.chapter, style), sp(40)];
  }).flat(),
  sp(100),
  h2('三、给董事会的一句话总结'),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      width: { size: W, type: WidthType.DXA },
      shading: { fill: '1F3864', type: ShadingType.CLEAR },
      margins: { top: 200, bottom: 200, left: 240, right: 240 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 }, children: [new TextRun({ text: `公司财务结构${da<0.6?'已安全':'仍有压力'}，但经营质量需要持续提升。`, font: 'Arial', size: 22, color: 'FFFFFF' })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 }, children: [new TextRun({ text: `${year}的首要任务：`, font: 'Arial', size: 26, bold: true, color: 'FFD700' })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '压缩应收账款 → 释放现金 → 修复问题区域盈利能力 → 再以更健康的基本面推进扩张', font: 'Arial', size: 22, color: 'BDD7EE' })] }),
      ]
    })] })]
  }),
  sp(200), divider(),
  p(`报告编制日期：${rptDate}　|　数据截止日：${year.replace('年','')}年12月31日`, { color: C.midGray, italic: true, size: 18 }),
  p('本报告含敏感财务数据，仅限内部使用，禁止对外传阅。', { color: C.midGray, italic: true, size: 18 }),
);

// ── 生成文档 ──
const doc = new Document({
  numbering: {
    config: [
      { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: 'Arial' } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22, color: C.darkGray } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 36, bold: true, font: 'Arial', color: C.darkBlue }, paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: 'Arial', color: C.midBlue }, paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, font: 'Arial', color: C.darkGray }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1280, bottom: 1440, left: 1440 } } },
    headers: {
      default: new Header({ children: [
        new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue, space: 1 } }, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: `${company}  |  ${year}度综合财务分析报告  |  内部参考`, font: 'Arial', size: 16, color: C.midGray })] })
      ] })
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.lightBlue, space: 1 } }, spacing: { before: 80, after: 0 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '第 ', font: 'Arial', size: 16, color: C.midGray }), new SimpleField({ instruction: 'PAGE' }), new TextRun({ text: ' 页', font: 'Arial', size: 16, color: C.midGray })] })
      ] })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log('Done: ' + outPath);
}).catch(e => { console.error(e.message); process.exit(1); });
