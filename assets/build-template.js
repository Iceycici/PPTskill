/* build-template.js — 构建入口模板
 *
 * 使用方法：
 *   1. 复制为 deck/build.js；把 lib-template.js 复制为 deck/lib.js 并完成令牌配置；
 *   2. 按篇章创建页面模块（p0_intro.js 等），每个模块导出 function(pres, L)；
 *   3. 模块内部：const s = L.content(pres, {part, kicker, title, notes}) 领内容页，
 *      L.divider(pres, {...}) 领篇章页；页码由 L.next() 全局递增，禁止手写；
 *   4. node build.js 后必须执行：
 *      python <skill>/scripts/fix_ppr.py 输出.pptx
 *      python <skill>/scripts/check_layout.py 输出.pptx
 *
 * 来源：2026-09-17 zcode 会话 deck/build.js 泛化版。
 */
const pptxgen = require("pptxgenjs");       // npm install -g pptxgenjs
const path = require("path");

const OUT = path.join(__dirname, "output.pptx");   // ← 改成你的输出文件名

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                // 13.33 × 7.5 英寸，必须显式
pres.author = "<作者>";
pres.title = "<课程标题>";

// 全局页码：模块共享同一计数器
const L = Object.assign({}, require("./lib-template.js"));
L.pageNo = { n: 0 };
L.next = () => ++L.pageNo.n;

// 按篇章顺序注册模块（模块清单按项目增减）
require("./p0_intro.js")(pres, L);
require("./p1_part1.js")(pres, L);
require("./p2a.js")(pres, L);
require("./p2b.js")(pres, L);
require("./p3_part3.js")(pres, L);
require("./p4_closing.js")(pres, L);

pres.writeFile({ fileName: OUT }).then(() => {
  console.log("DONE, slides =", L.pageNo.n);
  console.log("下一步（必做）：");
  console.log("  python <skill>/scripts/fix_ppr.py " + OUT);
  console.log("  python <skill>/scripts/check_layout.py " + OUT);
});
