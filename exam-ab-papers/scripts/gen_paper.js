#!/usr/bin/env node
// A/B 卷個人化考卷產生器（docx-js）
//   題目卷：A3 直式、雙欄、每人 paper_pages 頁（預設 2 頁＝正反兩面），表頭印班級／學號／姓名／卷別
//   答案卷：JIS B4（257×364 mm）單頁，名冊欄＋得分欄＋選擇題答案格＋計算題作答框
//   批改卷：與答案卷幾何完全相同，題號列印「小字題號＋紅色正解」，可疊在答案卷上批改
//
// 用法（在考試工作資料夾內執行，需有 exam.config.json、exam_data.js、orders.json、layout.json、students.json）：
//   node gen_paper.js paper <A|B|ALL> <out.docx> [probe]
//   node gen_paper.js sheet <A|B|ALL> <out.docx> [probe]
//   node gen_paper.js key   <A|B>     <out.docx>
//   ALL：A、B 卷混合，依 students.json 的順序（prep_roster.py 已依座位排好）
//   probe：不帶學生資料，只產生一份空白卷（plan_layout.py 量測版面用）
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, HeightRule, VerticalAlign,
  LineRuleType, SectionType, Footer, PageNumber, TabStopType, TableLayoutType, ColumnBreak,
} = require("docx");

const [KIND, WHICH, OUT, PROBE] = process.argv.slice(2);
if (!["paper", "sheet", "key"].includes(KIND) || !WHICH || !OUT) {
  console.error("usage: node gen_paper.js paper|sheet|key <A|B|ALL> <out.docx> [probe]"); process.exit(1);
}
const WD = process.cwd();
const readJSON = (f, dflt) => fs.existsSync(path.join(WD, f)) ? JSON.parse(fs.readFileSync(path.join(WD, f), "utf8")) : dflt;
const CFG = readJSON("exam.config.json", {});
const { mc, calc = [] } = require(path.join(WD, CFG.data || "exam_data.js"));
const ORDERS = readJSON("orders.json", { A: mc.map((_, i) => i), B: mc.map((_, i) => i) });
const LAYOUTS = readJSON("layout.json", {});
const STUDENTS = (PROBE || KIND === "key") ? [null]
  : readJSON("students.json", []).filter((s) => WHICH === "ALL" || s.ver === WHICH);
if (!STUDENTS.length) { console.error("students.json 沒有符合的學生"); process.exit(1); }

const TITLE = CFG.title || "期中考卷";                 // 題目卷標題，例：115年離散數學 期中考卷
const EXAM = CFG.exam_name || TITLE.replace(/卷$/, ""); // 答案卷／批改卷用，例：115年離散數學 期中考
const P1 = CFG.mc_points ?? 2, P2 = CFG.calc_points ?? 10;
const N = mc.length, M = calc.length, TOTAL = N * P1 + M * P2;
const PAGES = CFG.paper_pages || 2;
const F = Object.assign({ latin: "Times New Roman", cjk: "標楷體", math: "Cambria Math", mono: "Consolas" }, CFG.fonts || {});

const L0 = LAYOUTS.A || {};
const PT = L0.pt || CFG.pt || 14;                       // A、B 卷字級相同
const LINE = Math.round(PT * 20 * (L0.ls || CFG.ls || 1.2));
let VER, ORDER, BREAKS, GAPS, CALC_H;                   // 依每位學生的卷別切換
function useVersion(v) {
  const L = LAYOUTS[v] || {};
  VER = v; ORDER = ORDERS[v];
  BREAKS = L.breaks || [];                              // 在這些題號（本卷題號）前換欄
  GAPS = L.gaps || Array(2 * PAGES).fill(100);
  CALC_H = (L.calc_heights || CFG.sheet?.calc_heights || Array(M).fill(Math.floor(12950 / Math.max(M, 1))));
}

// ---------- 文字 ----------
const isMath = (c) => { const u = c.codePointAt(0);
  return (u >= 0x2190 && u <= 0x22ff) || (u >= 0x2070 && u <= 0x209f) || (u >= 0x1d00 && u <= 0x1dbf) || "²³¹×".includes(c); };
const fontOf = (f) => ({ ascii: f, hAnsi: f, cs: f, eastAsia: F.cjk });
// *x* 為斜體；數學符號改用數學字型；C(n, k) 逗號後改不斷行空白
function runs(text, o = {}) {
  const out = [], size = Math.round((o.pt || PT) * 2);
  text = text.replace(/C\(([^()]{1,8}?), /g, "C($1, ");
  text.split(/(\*[^*]+\*)/).forEach((seg) => {
    if (!seg) return;
    const italics = seg.length > 2 && seg.startsWith("*") && seg.endsWith("*");
    const s = italics ? seg.slice(1, -1) : seg;
    let buf = "", bm = null;
    const flush = () => { if (buf) out.push(new TextRun({ text: buf, size, bold: o.bold, italics, color: o.color,
      font: fontOf(o.font || (bm ? F.math : F.latin)) })); buf = ""; };
    for (const c of s) { const m = !o.font && isMath(c); if (bm !== null && m !== bm) flush(); bm = m; buf += c; }
    flush();
  });
  return out;
}
const exact = (line = LINE) => ({ line, lineRule: LineRuleType.EXACT });
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const allThin = { top: thin, bottom: thin, left: thin, right: thin, insideHorizontal: thin, insideVertical: thin };

function cell(text, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: o.valign || VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { left: o.ml ?? 60, right: 60, top: o.mt ?? 0, bottom: 0 },
    children: [new Paragraph({ alignment: o.align || AlignmentType.CENTER,
      spacing: { line: 240, lineRule: LineRuleType.AUTO },
      children: text ? runs(text, { pt: o.pt || 12, bold: o.bold }) : [] })],
  });
}
const paperFooter = (text) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: [new TextRun({ size: 20, font: fontOf(F.latin), children: [text, PageNumber.CURRENT, ` 頁／共 ${PAGES} 頁`] })] })] });
const sheetFooter = (text) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: runs(text, { pt: 10 }) })] });

// 名冊欄：班級／學號／姓名／卷別
function infoTable(st, totalW, labelW = 1000) {
  const rest = totalW - 4 * labelW - 1300;
  const vW = [Math.round(rest * 0.36), Math.round(rest * 0.32)]; vW.push(rest - vW[0] - vW[1]);
  const widths = [labelW, vW[0], labelW, vW[1], labelW, vW[2], labelW, 1300];
  const Lb = (t) => cell(t, labelW, { fill: "EDEDED", bold: true, pt: 12 });
  return new Table({
    width: { size: totalW, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED, borders: allThin,
    rows: [new TableRow({ height: { value: 600, rule: HeightRule.EXACT }, children: [
      Lb("班　級"), cell(st ? st.cls : "", vW[0], { pt: 14 }),
      Lb("學　號"), cell(st ? st.id : "", vW[1], { pt: 14 }),
      Lb("姓　名"), cell(st ? st.name : "", vW[2], { pt: 15, bold: true }),
      Lb("卷　別"), cell(`${VER} 卷`, 1300, { pt: 18, bold: true }),
    ] })],
  });
}

// =====================================================================
// 題目卷（A3 直式 297 × 420 mm，雙欄）
// =====================================================================
const PAGE_A3 = { width: 16838, height: 23811 };
const MAR = 794;
const TEXT_W = PAGE_A3.width - 2 * MAR;
const COL_SP = 454;
const COL_W = Math.floor((TEXT_W - COL_SP) / 2);
const HANG = 460;

function estWidth(s) {     // 以 12pt 為基準估算顯示寬度（pt）
  let w = 0;
  for (const c of s.replace(/\*/g, "")) {
    const u = c.codePointAt(0);
    if (isMath(c)) w += 9; else if (u >= 0x2e80) w += 12; else if (/[0-9]/.test(c)) w += 6;
    else if (/[A-Z]/.test(c)) w += 8.5; else if (/[a-z]/.test(c)) w += 5.5; else if (c === " ") w += 3; else w += 4;
  }
  return w * PT / 12;
}
const LET = ["A", "B", "C", "D", "E"];
function optionParas(opts, forced, colW = COL_W) {   // 依寬度自動 4／2／1 個一列
  const n = opts.length, avail = (colW - HANG) / 20;
  const labeled = opts.map((t, i) => `(${LET[i]}) ${t}`);
  const maxW = Math.max(...labeled.map(estWidth));
  const per = forced || (maxW <= avail / 4 - 8 && n === 4 ? 4 : maxW <= avail / 2 - 10 ? 2 : 1);
  const slot = Math.floor((colW - HANG) / per);
  const tabStops = [];
  for (let k = 1; k < per; k++) tabStops.push({ type: TabStopType.LEFT, position: HANG + k * slot });
  const paras = [];
  for (let i = 0; i < n; i += per) {
    const children = [];
    for (let k = 0; k < per && i + k < n; k++) {
      if (k > 0) children.push(new TextRun({ text: "\t", size: PT * 2 }));
      children.push(...runs(labeled[i + k]));
    }
    paras.push(new Paragraph({ children, tabStops, indent: { left: HANG }, spacing: exact(),
      keepNext: i + per < n, keepLines: true }));
  }
  return paras;
}
function pngSize(buf) { return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; }
// 圖片設定：exam.config.json 的 figures.<key> = { file, cm_w } 置於題幹下方置中；{ file, cm_h } 置於題幹右側
function imageRun(key) {
  const im = (CFG.figures || {})[key];
  if (!im) throw new Error(`exam.config.json 缺少 figures.${key}`);
  const data = fs.readFileSync(path.join(WD, im.file)), sz = pngSize(data), PX = 37.795;
  let wpx, hpx;
  if (im.cm_w) { wpx = im.cm_w * PX; hpx = wpx * sz.h / sz.w; } else { hpx = (im.cm_h || 4) * PX; wpx = hpx * sz.w / sz.h; }
  return { run: new ImageRun({ type: "png", data, transformation: { width: Math.round(wpx), height: Math.round(hpx) } }),
    widthDxa: Math.round(wpx / PX * 567), side: !im.cm_w };
}
function stemPara(num, text, gap) {
  return new Paragraph({
    children: [new TextRun({ text: `${num}.\t`, size: PT * 2, font: fontOf(F.latin) }), ...runs(text)],
    indent: { left: HANG, hanging: HANG }, spacing: { ...exact(), before: gap }, keepNext: true, keepLines: true,
  });
}
function mcBlock(q, num, gap) {
  const img = q.fig ? imageRun(q.fig) : null;
  if (img && img.side) {           // 左：題幹＋選項；右：附圖（無框表格）
    const rightW = img.widthDxa + 160, leftW = COL_W - rightW;
    return [new Table({
      width: { size: COL_W, type: WidthType.DXA }, columnWidths: [leftW, rightW],
      layout: TableLayoutType.FIXED, borders: noBorders,
      rows: [new TableRow({ cantSplit: true, children: [
        new TableCell({ width: { size: leftW, type: WidthType.DXA }, borders: noBorders,
          margins: { left: 0, right: 80, top: 0, bottom: 0 }, verticalAlign: VerticalAlign.CENTER,
          children: [stemPara(num, q.stem, gap), ...optionParas(q.opts, q.layout || 2, leftW)] }),
        new TableCell({ width: { size: rightW, type: WidthType.DXA }, borders: noBorders,
          margins: { left: 80, right: 0, top: 60, bottom: 0 }, verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [img.run],
            spacing: { line: 240, lineRule: LineRuleType.AUTO } })] }),   // 圖片段落不可用固定行高，否則被裁切
      ] })],
    })];
  }
  const out = [stemPara(num, q.stem, gap)];
  if (q.code) q.code.forEach((ln) => out.push(new Paragraph({
    children: [new TextRun({ text: ln, size: Math.round((PT - 1) * 2), font: fontOf(F.mono) })],
    indent: { left: HANG + 280 }, spacing: exact(Math.round((PT - 1) * 20 * 1.2)), keepNext: true, keepLines: true })));
  if (img) out.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [img.run],
    spacing: { before: 40, after: 40, line: 240, lineRule: LineRuleType.AUTO }, keepNext: true }));
  out.push(...optionParas(q.opts, q.layout));
  return out;
}
const heading = (text) => new Paragraph({ children: runs(text, { bold: true, pt: PT + 1 }),
  spacing: { line: Math.round((PT + 1) * 20 * 1.3), lineRule: LineRuleType.EXACT, before: 120, after: 40 }, keepNext: true });
const colBreak = () => new Paragraph({ children: [new ColumnBreak()], spacing: exact(20) });

function paperSections(st, first) {
  const sides = PAGES === 2 ? "（正反兩面）" : "";
  const calcNote = M ? `，計算或問答題 ${M} 題每題 ${P2} 分共 ${M * P2} 分` : "";
  const note = CFG.paper_note || `注意：本試卷共 ${PAGES} 頁${sides}，選擇題 ${N} 題每題 ${P1} 分共 ${N * P1} 分${calcNote}，滿分 ${TOTAL} 分。所有答案請寫在答案卷上，寫在題目卷上不予計分；請先核對題目卷與答案卷的卷別、姓名是否相符，考試結束後兩者一併繳回。`;
  const head = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100, line: 240, lineRule: LineRuleType.AUTO },
      children: runs(TITLE, { pt: 22, bold: true }) }),
    infoTable(st, TEXT_W),
    new Paragraph({ spacing: { before: 100, after: 60, line: 300, lineRule: LineRuleType.EXACT },
      indent: { left: 720, hanging: 720 }, children: runs(note, { pt: 11 }) }),
  ];
  const body = [heading(`一、選擇題（每題 ${P1} 分，共 ${N * P1} 分）`)];
  let col = 0;
  ORDER.forEach((qi, k) => {
    const num = k + 1;
    if (BREAKS.includes(num)) { body.push(colBreak()); col++; }
    const gap = (k === 0 || BREAKS.includes(num)) ? 40 : GAPS[col];
    body.push(...mcBlock(mc[qi], num, gap));
  });
  if (BREAKS.includes(N + 1)) { body.push(colBreak()); col++; }
  if (M) {
    body.push(heading(`二、計算或問答題（每題 ${P2} 分，共 ${M * P2} 分；請在答案卷上作答）`));
    calc.forEach((q, i) => {
      body.push(stemPara(i + 1, q.stem, GAPS[col] ?? 100));
      if (q.defs) q.defs.forEach((d) => body.push(new Paragraph({ children: runs(d, { pt: PT - 0.5 }),
        indent: { left: HANG + 280 }, spacing: exact(Math.round((PT - 0.5) * 20 * 1.25)), keepLines: true })));
    });
  }
  body.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { ...exact(), before: 240 },
    children: runs("— 試題結束 —", { pt: PT - 1 }) }));

  const tag = st ? `${st.id} ${st.name}　` : "";
  const ft = `${TITLE}　${VER} 卷　${tag}第 `;
  const page = { size: PAGE_A3, margin: { top: MAR, bottom: MAR + 120, left: MAR, right: MAR, header: 400, footer: 400 },
    pageNumbers: { start: 1 } };                           // 每位學生頁碼重新起算
  return [
    { properties: { page, type: first ? undefined : (PAGES === 2 ? SectionType.ODD_PAGE : SectionType.NEXT_PAGE) },
      footers: { default: paperFooter(ft) }, children: head },
    { properties: { page: { ...page, pageNumbers: undefined }, type: SectionType.CONTINUOUS,
        column: { count: 2, space: COL_SP, separate: true, equalWidth: true } },
      footers: { default: paperFooter(ft) }, children: body },
  ];
}

// =====================================================================
// 答案卷／批改卷（JIS B4 直式 257 × 364 mm）
// =====================================================================
const PAGE_B4 = { width: 14570, height: 20636 };
const BM = 794, B_W = PAGE_B4.width - 2 * BM;

function keyCell(n, w) {            // 批改卷題號格：小字題號＋紅色大字正解
  return new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { left: 30, right: 30, top: 0, bottom: 0 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 330, lineRule: LineRuleType.EXACT },
      children: [new TextRun({ text: `${n} `, size: 16, font: fontOf(F.latin) }),
        new TextRun({ text: mc[ORDER[n - 1]].ans, size: 28, bold: true, color: "C00000", font: fontOf(F.latin) })] })] });
}
// 批改卷只能改「格子裡的文字」，任何尺寸都必須與答案卷相同，疊放時格線才會對齊
function sheetSection(st, first, key = false) {
  const children = [];
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100, line: 240, lineRule: LineRuleType.AUTO },
    children: runs(`${EXAM}　${key ? "批改卷" : "答案卷"}`, { pt: 20, bold: true }) }));
  children.push(infoTable(key ? { cls: "", id: "", name: "批改用" } : st, B_W));
  // 得分欄
  const labels = [`選擇題（${N * P1}）`, ...calc.map((_, i) => `計算 ${i + 1}（${P2}）`), "總　分"];
  const sw = Math.floor(B_W / labels.length), sWidths = labels.map((_, i) => i < labels.length - 1 ? sw : B_W - sw * (labels.length - 1));
  children.push(new Paragraph({ spacing: { before: 0, after: 0, line: 160, lineRule: LineRuleType.EXACT }, children: [] }));
  children.push(new Table({ width: { size: B_W, type: WidthType.DXA }, columnWidths: sWidths, layout: TableLayoutType.FIXED,
    borders: allThin, rows: [
      new TableRow({ height: { value: 380, rule: HeightRule.EXACT }, children: labels.map((t, i) =>
        cell(t, sWidths[i], { fill: "EDEDED", bold: true, pt: 11 })) }),
      new TableRow({ height: { value: 720, rule: HeightRule.EXACT }, children: sWidths.map((w) => cell("", w)) }),
    ] }));
  // 大題說明：必須維持單行（固定行高）
  children.push(new Paragraph({ spacing: { before: 100, after: 0, line: 280, lineRule: LineRuleType.EXACT },
    children: runs(key ? `一、選擇題正解（${VER} 卷）：正解標示於題號列，對齊答案卷的答案格即可逐題批改，每題 ${P1} 分。`
      : `一、選擇題（每題 ${P1} 分，共 ${N * P1} 分）：請以大寫英文字母作答，字跡須清楚可辨。`, { pt: 12, bold: true }) }));
  // 選擇題答案格
  const PER = CFG.sheet?.per_row || 18, LW = 1000, gw = Math.floor((B_W - LW) / PER), gWidths = [B_W - gw * PER, ...Array(PER).fill(gw)];
  const rows = [];
  for (let s = 1; s <= N; s += PER) {
    const nums = [], blanks = [];
    for (let k = 0; k < PER; k++) {
      const n = s + k, real = n <= N;
      nums.push(key && real ? keyCell(n, gw) : cell(real ? String(n) : "", gw, { fill: real ? "EDEDED" : "BFBFBF", pt: 11 }));
      blanks.push(cell("", gw, { fill: real ? undefined : "BFBFBF" }));
    }
    rows.push(new TableRow({ height: { value: 360, rule: HeightRule.EXACT }, children: [cell(key ? "正解" : "題號", gWidths[0], { fill: "EDEDED", bold: true, pt: 11 }), ...nums] }));
    rows.push(new TableRow({ height: { value: 700, rule: HeightRule.EXACT }, children: [cell("答案", gWidths[0], { fill: "EDEDED", bold: true, pt: 11 }), ...blanks] }));
  }
  children.push(new Paragraph({ spacing: { before: 0, after: 0, line: 80, lineRule: LineRuleType.EXACT }, children: [] }));
  children.push(new Table({ width: { size: B_W, type: WidthType.DXA }, columnWidths: gWidths, layout: TableLayoutType.FIXED, borders: allThin, rows }));
  if (M) {
    children.push(new Paragraph({ spacing: { before: 160, after: 60, line: 280, lineRule: LineRuleType.EXACT },
      children: runs(`二、計算或問答題（每題 ${P2} 分，共 ${M * P2} 分）：請寫出完整計算或證明過程。`, { pt: 12, bold: true }) }));
    children.push(new Table({ width: { size: B_W, type: WidthType.DXA }, columnWidths: [B_W], layout: TableLayoutType.FIXED, borders: allThin,
      rows: calc.map((q, i) => new TableRow({ cantSplit: true, height: { value: CALC_H[i], rule: HeightRule.EXACT }, children: [
        cell(`${i + 1}.`, B_W, { align: AlignmentType.LEFT, valign: VerticalAlign.TOP, bold: true, pt: 12, ml: 100, mt: 60 })] })) }));
  }
  const tag = st ? `　${st.cls}　${st.id}　${st.name}` : "";
  return { properties: { page: { size: PAGE_B4, margin: { top: BM, bottom: BM, left: BM, right: BM, header: 400, footer: 360 } },
      type: first ? undefined : SectionType.NEXT_PAGE },
    footers: { default: sheetFooter(`${EXAM} ${key ? "批改卷" : "答案卷"}　${VER} 卷${tag}`) }, children };
}

// =====================================================================
const sections = [];
STUDENTS.forEach((st, i) => {
  useVersion(st ? st.ver : WHICH);
  if (KIND === "paper") sections.push(...paperSections(st, i === 0));
  else sections.push(sheetSection(st, i === 0, KIND === "key"));
});
const doc = new Document({
  creator: CFG.author || "", title: `${EXAM} ${{ paper: "題目卷", sheet: "答案卷", key: "批改卷" }[KIND]}${WHICH === "ALL" ? "（A、B 卷）" : ` ${WHICH}卷`}`,
  styles: { default: { document: { run: { font: fontOf(F.latin), size: PT * 2 },
    paragraph: { spacing: { line: LINE, lineRule: LineRuleType.EXACT } } } } },
  sections,
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, `${STUDENTS.length} 份`); });
