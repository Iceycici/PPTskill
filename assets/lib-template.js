/* lib-template.js — PPT 设计系统模板（pptxgenjs 组件库，深蓝学术风）
 *
 * 使用方法：
 *   1. 复制本文件为 deck/lib.js，先改"【设计令牌区】"（配色/篇章色/页脚文案/截图根路径）；
 *   2. 各篇章页面模块 require 本文件，用组件函数拼装页面（见 build-template.js）；
 *   3. 组件坐标一律英寸；颜色只从令牌 C 取，页面代码禁写裸色值。
 *
 * 来源：2026-09-17 zcode 会话 69 页中文教学课件 lib.js 定稿泛化版。
 * 组件规格详见 references/02_幻灯片设计规范.md。
 */
const pptxgen = null; // 由 build.js 统一 require，本文件只产出纯函数与令牌

/* ==================== 【设计令牌区】换主题只改这里 ==================== */
const W = 13.33, H = 7.5;                 // LAYOUT_WIDE 画布（英寸）
const F = "微软雅黑";                      // 全篇唯一字体
const FOOTER = "AI提示词工程 · 从新手到高手"; // 左下页脚文案（全篇唯一形态）
const IMG = "../images";                  // 截图根路径（按项目改）

const C = {
  navy: "0E2A47", navy2: "16395C",       // 深蓝底 / 深底分层
  ink: "1C3D5A", body: "3A5570",         // 亮底标题 / 正文
  muted: "8296AB", faint: "B9C6D4",      // 次要 / 弱化
  line: "D8E1EA", tint: "F1F5F9",        // 分隔线 / 中性卡底
  paper: "FFFFFF", onDark: "FFFFFF", onDarkMuted: "AFC2D6",
  gold: "E9A13B",                        // 模板卡 label / 暗底强调
  good: "12805C", goodT: "E7F3EE",       // 正例
  bad: "B4443C", badT: "FBEAE8",         // 反例
  p1: "1D6FB8", p1t: "EAF3FA", p1d: "155587",   // 篇章1 蓝
  p2: "12805C", p2t: "E7F3EE", p2d: "0E6B4C",   // 篇章2 绿
  p3: "C05621", p3t: "FBEDE2", p3d: "A34518",   // 篇章3 橙
};

const PART = {   // 篇章主题化：篇章页与该篇 kicker 同色
  1: { color: C.p1, tint: C.p1t, dark: C.p1d, label: "第一篇 · <篇章名>" },
  2: { color: C.p2, tint: C.p2t, dark: C.p2d, label: "第二篇 · <篇章名>" },
  3: { color: C.p3, tint: C.p3t, dark: C.p3d, label: "第三篇 · <篇章名>" },
};
/* ==================== 令牌区结束 ==================== */

// 页码计数器：build.js 里 const L = { pageNo: { n: 0 }, next: () => ++L.pageNo.n }
const pageNo = { n: 0 };
const next = () => ++pageNo.n;

const bu = (color) => ({ code: "25B8", color });        // ▸ 主级列表符
const bud = (color) => ({ code: "2013", color });       // – 次级列表符

/* 内容页骨架：kicker + 标题 + 页脚 + 页码 + 备注 */
function content(pres, o) {
  const s = pres.addSlide();
  const P = PART[o.part || 1];
  s.background = { color: C.paper };
  s.addText(o.kicker || P.label, { x: 0.62, y: 0.34, w: o.kickW || 11.9, h: 0.32,
    fontSize: 12.5, bold: true, color: P.color, fontFace: F, margin: 0 });
  s.addText(o.title, { x: 0.6, y: 0.62, w: 12.1, h: 0.62,
    fontSize: o.titleSize || 28, bold: true, color: C.ink, fontFace: F, margin: 0 });
  s.addText(FOOTER, { x: 0.62, y: 7.12, w: 4.0, h: 0.26,
    fontSize: 10.5, color: C.faint, fontFace: F, margin: 0 });
  s.addText(String(next()), { x: 12.35, y: 7.12, w: 0.5, h: 0.26, align: "right",
    fontSize: 10.5, color: C.faint, fontFace: F, margin: 0 });
  if (o.notes) s.addNotes(o.notes);
  return s;
}

/* 篇章页：大号水印数字 + 篇章色 + 条目列表 */
function divider(pres, o) {
  const s = pres.addSlide();
  const P = PART[o.part || 1];
  s.background = { color: C.navy };
  s.addText(String(o.num).padStart(2, "0"), { x: 8.3, y: 1.2, w: 5.0, h: 5.4,
    fontSize: 210, bold: true, color: P.dark, fontFace: F, margin: 0, transparency: 55 });
  s.addText("PART 0" + o.num, { x: 0.9, y: 1.85, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: P.color, fontFace: F, charSpacing: 4, margin: 0 });
  s.addText(o.title, { x: 0.88, y: 2.3, w: 9, h: 0.9,
    fontSize: 40, bold: true, color: C.onDark, fontFace: F, margin: 0 });
  if (o.subtitle) s.addText(o.subtitle, { x: 0.9, y: 4.5, w: 9, h: 0.4,
    fontSize: 14, color: C.onDarkMuted, fontFace: F, margin: 0 });
  (o.items || []).forEach((t, i) => {
    s.addText("▸ " + t, { x: 0.9, y: 5.1 + i * 0.42, w: 8.5, h: 0.38,
      fontSize: 13, color: C.onDark, fontFace: F, margin: 0 });
  });
  next();
  return s;
}

/* 圆角卡片 */
function card(s, x, y, w, h, fill, shadow) {
  s.addShape("roundRect", { x, y, w, h, fill: { color: fill }, rectRadius: 0.08,
    shadow: shadow ? { type: "outer", color: "0E2A47", opacity: 0.15, blur: 6, offset: 2, angle: 90 } : undefined });
}

/* 胶囊标签 */
function chip(s, x, y, w, h, text, fill, color) {
  s.addShape("rect", { x, y, w, h, fill: { color: fill }, rectRadius: h / 2.2 });
  s.addText(text, { x, y, w, h, align: "center", valign: "middle",
    fontSize: 11.5, bold: true, color: color || "FFFFFF", fontFace: F, margin: 0 });
}

/* 列表：items 支持 "文本" 或 [加粗lead, 正文] */
function bullets(s, x, y, w, h, items, o = {}) {
  s.addText(items.map(it => Array.isArray(it)
    ? [{ text: it[0], options: { bold: true, color: o.leadColor || C.ink } },
       { text: it[1], options: { color: o.color || C.body } }]
    : [{ text: it, options: { color: o.color || C.body } }]),
    { x, y, w, h, fontSize: o.size || 12, fontFace: F, margin: 0,
      lineSpacingMultiple: 1.15, paraSpaceAfter: o.gap || 6,
      bullet: o.noBullet ? false : bu(o.buColor || C.p1) });
}

/* 纵向编号行：大数字 + 标题 + 说明 + 分隔线 */
function numRows(s, x, y, w, rows, o = {}) {
  const rh = o.rowH || 1.0;
  rows.forEach((r, i) => {
    const yy = y + i * rh;
    s.addText(String(i + 1 + (o.start0 || 0)), { x, y: yy, w: 0.6, h: rh,
      fontSize: 26, bold: true, color: o.color || C.p1, fontFace: F, margin: 0 });
    s.addText([{ text: r[0], options: { bold: true, color: C.ink } },
               { text: "\n" + r[1], options: { color: C.body, fontSize: (o.size || 12) - 0.5 } }],
      { x: x + 0.7, y: yy, w: w - 0.8, h: rh, fontSize: o.size || 12, fontFace: F,
        margin: 0, lineSpacingMultiple: 1.1 });
    if (i < rows.length - 1)
      s.addShape("line", { x: x + 0.7, y: yy + rh - 0.08, w: w - 0.8, h: 0,
        line: { color: C.line, width: 0.75 } });
  });
}

/* 横向步骤链：等宽步骤卡 + "→"；start 支持编号不从 1 起（删页重编场景） */
function flow(s, x, y, w, steps, o = {}) {
  const n = steps.length, gap = 0.32;
  const cw = (w - (n - 1) * gap) / n, ch = o.h || 0.9;
  steps.forEach((t, i) => {
    const xx = x + i * (cw + gap);
    card(s, xx, y, cw, ch, o.fill || C.tint);
    const label = (o.start ? (o.start + i) : i + 1) + "  " + t;
    s.addText(label, { x: xx + 0.1, y, w: cw - 0.2, h: ch, align: "center", valign: "middle",
      fontSize: o.size || 11.5, color: o.color || C.body, fontFace: F, margin: 0.03 });
    if (i < n - 1)
      s.addText("→", { x: xx + cw, y, w: gap, h: ch, align: "center", valign: "middle",
        fontSize: 14, bold: true, color: C.muted, fontFace: F, margin: 0 });
  });
}

/* 深蓝提示词模板卡：金色 label + 逐字模板正文（内容来自 sources/_templates.md） */
function promptCard(s, x, y, w, h, label, text, o = {}) {
  card(s, x, y, w, h, C.navy);
  s.addText(label, { x: x + 0.26, y: y + 0.14, w: w - 0.5, h: 0.3,
    fontSize: 12.5, bold: true, color: C.gold, fontFace: F, margin: 0 });
  s.addText(text, { x: x + 0.26, y: y + 0.48, w: w - 0.52, h: h - 0.62,
    fontSize: o.size || 11.5, color: C.onDark, fontFace: F, margin: 0,
    lineSpacingMultiple: 1.18 });
}

/* 双栏对比：左反例（红）右正例（绿） */
function versus(s, x, y, w, h, left, right, o = {}) {
  const half = (w - 0.3) / 2;
  card(s, x, y, half, h, o.leftFill || C.badT);
  s.addText(left.title, { x: x + 0.26, y: y + 0.14, w: half - 0.5, h: 0.3,
    fontSize: 12.5, bold: true, color: o.leftColor || C.bad, fontFace: F, margin: 0 });
  s.addText(left.items.map(t => ({ text: t, options: { breakLine: true } })),
    { x: x + 0.26, y: y + 0.5, w: half - 0.5, h: h - 0.65, fontSize: o.size || 12,
      color: C.body, fontFace: F, margin: 0, lineSpacingMultiple: 1.15, paraSpaceAfter: 4 });
  const rx = x + half + 0.3;
  card(s, rx, y, half, h, o.rightFill || C.goodT);
  s.addText(right.title, { x: rx + 0.26, y: y + 0.14, w: half - 0.5, h: 0.3,
    fontSize: 12.5, bold: true, color: o.rightColor || C.good, fontFace: F, margin: 0 });
  s.addText(right.items.map(t => ({ text: t, options: { breakLine: true } })),
    { x: rx + 0.26, y: y + 0.5, w: half - 0.5, h: h - 0.65, fontSize: o.size || 12,
      color: C.body, fontFace: F, margin: 0, lineSpacingMultiple: 1.15, paraSpaceAfter: 4 });
}

/* 深蓝金句条 */
function quote(s, x, y, w, text, o = {}) {
  card(s, x, y, w, o.h || 0.6, C.navy);
  s.addText([{ text: "❝ ", options: { color: C.gold, bold: true } },
             { text, options: { color: C.onDark } }],
    { x: x + 0.25, y, w: w - 0.5, h: o.h || 0.6, valign: "middle",
      fontSize: o.size || 12.5, fontFace: F, margin: 0 });
}

/* 截图 + 图注：按文件名首字符定位 module_N 子目录 */
function shot(s, file, x, y, w, caption, ar) {
  const mod = "module_" + String(file).trim()[0];
  s.addImage({ path: IMG + "/" + mod + "/" + file, x, y, w, h: w / ar });
  if (caption) s.addText("图：" + caption, { x, y: y + w / ar + 0.04, w, h: 0.26,
    fontSize: 10, color: C.muted, fontFace: F, margin: 0 });
}

/* 表格：表头深色 + 斑马纹；单元格可传 {text, options} 对象 */
function table(s, x, y, w, header, rows, o = {}) {
  const head = header.map(t => ({ text: t,
    options: { bold: true, color: "FFFFFF", fill: { color: o.headFill || C.navy } } }));
  const body = rows.map((r, ri) => r.map(c => {
    const cell = (typeof c === "object" && c !== null && !Array.isArray(c)) ? c : { text: c };
    return { text: cell.text,
      options: Object.assign({ color: C.body, fill: { color: ri % 2 ? "F7FAFC" : "FFFFFF" } },
                             cell.options || {}) };
  }));
  s.addTable([head, ...body], { x, y, w, colW: o.colW, rowH: o.rowH || 0.42,
    fontSize: o.size || 11, fontFace: F, valign: "middle",
    border: { type: "solid", color: C.line, pt: 0.5 }, margin: 0.06 });
}

module.exports = { W, H, F, C, PART, IMG, FOOTER, bu, bud, content, divider, card, chip,
  bullets, numRows, flow, promptCard, versus, quote, shot, table, pageNo, next };
