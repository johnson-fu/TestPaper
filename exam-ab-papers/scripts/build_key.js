#!/usr/bin/env node
// 教師用參考答案（A4）：A、B 卷答案速查表＋題號對照與解析＋計算題詳解
// 用法（在考試工作資料夾內）：node build_key.js <out.docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, VerticalAlign, LineRuleType, TableLayoutType,
} = require("docx");

const WD = process.cwd();
const OUT = process.argv[2] || "參考答案.docx";
const readJSON = (f, d) => fs.existsSync(path.join(WD, f)) ? JSON.parse(fs.readFileSync(path.join(WD, f), "utf8")) : d;
const CFG = readJSON("exam.config.json", {});
const { mc, calc = [] } = require(path.join(WD, CFG.data || "exam_data.js"));
const ORD = readJSON("orders.json", { A: mc.map((_, i) => i), B: mc.map((_, i) => i) });
const TITLE = CFG.title || "期中考卷";
const P1 = CFG.mc_points ?? 2, P2 = CFG.calc_points ?? 10;
const F = Object.assign({ latin: "Times New Roman", cjk: "標楷體", math: "Cambria Math" }, CFG.fonts || {});

const isMath = (c) => { const u = c.codePointAt(0);
  return (u >= 0x2190 && u <= 0x22ff) || (u >= 0x2070 && u <= 0x209f) || (u >= 0x1d00 && u <= 0x1dbf) || "²³¹×".includes(c); };
const fontOf = (f) => ({ ascii: f, hAnsi: f, cs: f, eastAsia: F.cjk });
function runs(text, pt = 10.5, bold = false) {
  const out = []; let buf = "", m0 = null;
  const flush = () => { if (buf) out.push(new TextRun({ text: buf, size: pt * 2, bold, font: fontOf(m0 ? F.math : F.latin) })); buf = ""; };
  for (const c of text.replace(/\*/g, "")) { const m = isMath(c); if (m0 !== null && m !== m0) flush(); m0 = m; buf += c; }
  flush(); return out;
}
const P = (text, o = {}) => new Paragraph({ children: runs(text, o.pt, o.bold), alignment: o.align,
  spacing: { before: o.before || 0, after: o.after || 0, line: o.line || 300, lineRule: LineRuleType.AUTO }, indent: o.indent });
const thin = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const B = { top: thin, bottom: thin, left: thin, right: thin, insideHorizontal: thin, insideVertical: thin };
const TW = 9638;                                       // A4 文字寬（邊界 2 cm）
const cell = (text, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
  shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
  margins: { left: 70, right: 70, top: 30, bottom: 30 },
  children: [P(text, { align: o.align || AlignmentType.CENTER, pt: o.pt || 10.5, bold: o.bold, line: 260 })] });

const PER = 12, LW = 900, cw = Math.floor((TW - LW) / PER), widths = [TW - cw * PER, ...Array(PER).fill(cw)];
function quickTable(ver) {
  const ord = ORD[ver], rows = [];
  for (let s = 1; s <= ord.length; s += PER) {
    const n = [], a = [];
    for (let k = 0; k < PER; k++) { const i = s + k, real = i <= ord.length;
      n.push(cell(real ? String(i) : "", cw, { fill: "EDEDED", pt: 10 }));
      a.push(cell(real ? mc[ord[i - 1]].ans : "", cw, { bold: true, pt: 11 })); }
    rows.push(new TableRow({ children: [cell("題號", widths[0], { fill: "EDEDED", bold: true }), ...n] }));
    rows.push(new TableRow({ children: [cell("答案", widths[0], { fill: "EDEDED", bold: true }), ...a] }));
  }
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED, borders: B, rows });
}
const bPos = {}; ORD.B.forEach((qi, k) => { bPos[qi] = k + 1; });
const ew = [760, 760, 700, TW - 2220];
const H = (t, w) => cell(t, w, { fill: "EDEDED", bold: true });
const expRows = [new TableRow({ tableHeader: true, children: [H("A卷題號", ew[0]), H("B卷題號", ew[1]), H("答案", ew[2]), H("解析", ew[3])] })];
ORD.A.forEach((qi, k) => { const q = mc[qi]; expRows.push(new TableRow({ cantSplit: true, children: [
  cell(String(k + 1), ew[0]), cell(String(bPos[qi]), ew[1]), cell(`(${q.ans})`, ew[2], { bold: true }),
  cell(q.exp || "", ew[3], { align: AlignmentType.LEFT, pt: 10 })] })); });

const children = [
  P(`${TITLE}　參考答案（教師用）`, { pt: 16, bold: true, align: AlignmentType.CENTER, after: 120 }),
  P(`一、選擇題（每題 ${P1} 分，共 ${mc.length * P1} 分）`, { pt: 12, bold: true, before: 60, after: 60 }),
  P("A 卷", { pt: 12, bold: true, after: 40 }), quickTable("A"),
  P("B 卷", { pt: 12, bold: true, before: 160, after: 40 }), quickTable("B"),
  P("題號對照與解析（依 A 卷題號排列）", { pt: 12, bold: true, before: 200, after: 60 }),
  new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: ew, layout: TableLayoutType.FIXED, borders: B, rows: expRows }),
];
if (calc.length) {
  children.push(P(`二、計算或問答題（每題 ${P2} 分，共 ${calc.length * P2} 分；A、B 卷題目與順序相同）`, { pt: 12, bold: true, before: 240, after: 60 }));
  calc.forEach((q, i) => {
    children.push(P(`${i + 1}. ${q.stem}`, { pt: 10.5, bold: true, before: 100, indent: { left: 300, hanging: 300 } }));
    (q.sol || []).forEach((s) => children.push(P(s, { pt: 10.5, indent: { left: 300 } })));
  });
}
const doc = new Document({ creator: CFG.author || "", title: `${TITLE} 參考答案`,
  styles: { default: { document: { run: { font: fontOf(F.latin), size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 },
    margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }] });
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("wrote", OUT); });
