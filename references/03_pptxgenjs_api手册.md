# 03 · pptxgenjs API 手册（构建侧速查与常见坑）

> 来源：提炼自官方技能 document-skills:pptx 0.1.4（Z.AI，License: Proprietary）Part 2 "pptxgenjs in Depth" 的要点，并补充 2026-09-17 zcode 会话实战中踩过的坑（标注【实战】）。完整 API 以官方文档为准：https://gitbrent.github.io/PptxGenJS/

## 1. 基本骨架

```js
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.33 × 7.5 英寸，必须显式设置
pres.author = "…"; pres.title = "…";
// 按序加载篇章模块，每个模块导出 function(pres, L)
require("./p0_intro.js")(pres, L);
require("./p1_part1.js")(pres, L);
pres.writeFile({ fileName: OUT }).then(() => console.log("DONE, slides =", L.pageNo.n));
```

单位一律**英寸**（浮点）。坐标 `x/y` 为左上角，`w/h` 为宽高。

## 2. 文本

```js
s.addText("普通文本", { x, y, w, h, fontSize: 12, color: C.body, fontFace: F, margin: 0 });
s.addText([                                  // 富文本 run 数组
  { text: "加粗前导：", options: { bold: true, color: C.ink } },
  { text: "正文说明", options: { color: C.body } },
], { x, y, w, h, fontSize: 12, fontFace: F, margin: 0, lineSpacingMultiple: 1.15 });
```

要点：
- **`margin: 0` 必写**，否则文本框四周有内边距，与卡片错位；
- `valign: "middle"` 用于单行垂直居中（胶囊、表格、行）；
- `lineSpacingMultiple` 控制行距（1.1–1.2 常用）；
- 字号单位 pt；颜色一律走令牌，不写裸色值。

【实战】pptxgenjs 的富文本段落会输出重复 `<a:pPr>`，PowerPoint 打开报修复——构建后必须跑 `scripts/fix_ppr.py`。

## 3. 列表与编号

- 项目符用 `bullet: { code: "25B8" }`（▸）或字符 `–`；层级靠两个 bullet 工厂 `bu(color)`/`bud(color)` 区分主次；
- 列表项支持二元组 `[加粗lead, 正文]`，由组件内部拆 run；
- 编号行（numRows）不用自动编号，用大数字文本框自绘，才能控制颜色与字号。

## 4. 形状与卡片

```js
s.addShape("roundRect", { x, y, w, h, fill: { color: C.tint }, rectRadius: 0.08,
  shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 6, offset: 2, angle: 90 } });
s.addShape("rect", { x, y, w: 0.9, h: 0.36, fill: { color: C.p1 }, rectRadius: 0.18 }); // 胶囊 chip
```

- `rectRadius` 单位英寸；胶囊取 `h/2.2` 视觉最稳；
- 阴影克制使用：只在最上层卡片，opacity ≤0.2、blur ≤6、offset ≤2。

## 5. 图片

```js
s.addImage({ path: "module_2/2.3.1.png", x, y, w, h: w / ar });  // h 按 宽/宽高比 计算
s.addText("图：xxx", { x, y: y + h + 0.04, w, h: 0.26, fontSize: 10, color: C.muted, fontFace: F, margin: 0 });
```

【实战】路径按文件名首字符自动归入 `module_N` 子目录（`shot` 组件内置）；坐标与宽高比来自 `sources/_image_dims.json`，手写必错。

## 6. 表格

```js
s.addTable(rows, { x, y, w, colW: [2.0, 1.9, ...], rowH: 0.42, fontSize: 11, fontFace: F,
  border: { type: "solid", color: C.line, pt: 0.5 },
  fill: { color: "FFFFFF" } });
```

- 表头行单元格：`{ text: "列名", options: { bold: true, color: "FFFFFF", fill: { color: C.navy } } }`；
- 斑马纹在组件里按行号奇偶设 `F7FAFC`；
- 单元格要变色/加粗时传对象而非字符串；
- 【实战】列宽用 `colW` 数组显式指定且总和 = `w`，否则 PowerPoint 自动拉伸错位。

## 7. 图表

```js
s.addChart(pres.ChartType.bar, [{ name: "系列", labels: [...], values: [...] }],
  { x, y, w, h, barDir: "bar", chartColors: [C.p1], showValue: true, dataLabelFontSize: 10 });
```

- 只画资料里有真实数据的图；横向条形图（`barDir: "bar"`) 适合类别名较长的中文场景；
- 堆叠条形图注意 `outEnd` 数据标签在某些版本溢出的已知问题（官方 Part 2 提及），必要时改 `showValue: false` 用图例。

## 8. 母版与演讲者备注

```js
s.addNotes("来源：讲义1.4。要点：肽类案例；引用Top5仅名次无数值。");
```

- 备注在每页构建时同步写入（content 组件的 `notes` 参数）；
- 备注文案遵守中文排版规范（中英文/数字间空格），后续 polish_notes.py 会统一扫描。

## 9. 常见坑清单（实战汇总）

| 坑 | 现象 | 对策 |
|---|---|---|
| 忘设 `LAYOUT_WIDE` | 4:3 画布，重排全部 | 入口第一行显式设置 |
| `margin` 未清零 | 文字与卡片边缘错位 | 所有自绘文本 `margin: 0` |
| 重复 `<a:pPr>` | PowerPoint 打开报修复 | `scripts/fix_ppr.py` |
| 组件参数遮蔽 | 同名参数内外层覆盖（如 flow 的 `scene` 与外层 `scene`） | 组件内部参数重命名隔离 |
| 图片路径含子目录 | addImage 找不到文件报错 | `shot` 组件统一处理 `module_N` 前缀 |
| 裸色值散落 | 换主题要全文搜索 | 令牌只在 lib.js 顶部定义 |
| 页码手写 | 删插页后全错 | `L.next()` 全局递增取号 |

## 10. 构建后处理链（顺序固定）

```
node build.js                    # 生成 pptx
python scripts/fix_ppr.py xx.pptx   # 兼容修复（必做）
python scripts/check_layout.py xx.pptx   # 溢出/重叠预检
powershell ... render_all.ps1    # 渲染 → 视觉验收
```
