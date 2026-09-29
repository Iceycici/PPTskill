# PPTskill · 安装与使用说明

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一个面向**中文教学/汇报类 PPT 全生命周期**的 Agent 技能：从讲义资料从零构建课件、在现有 pptx 上原位删改（保护用户手动修改）、序号重编与交叉引用清理、通篇文案校对、双闸门质检，直到交付说明与修订审计。

## 项目解决什么问题

| 没有 skill 时的常见痛点 | 本 skill 的解法 |
|---|---|
| 让 AI 重排/改写课件，会毁掉手动修改过的页面 | 主线 B 原位编辑：python-pptx 只改 run 文本、不动任何格式，用户手改区受保护，绝不重新生成 |
| 删掉几页后全篇编号错乱、交叉引用失效 | 单遍正则重编号 + 精确子串清单，一次替换完成，杜绝级联错替；页码文本框按位置重写为实际页序 |
| 生成的 PPT 是图片型，无法点选编辑 | 全部产出为原生 PowerPoint 对象（文字/表格/流程图/模板卡），PowerPoint 与 WPS 均可打开 |
| 文案替换后文字溢出、截断、版面被破坏 | 长度守卫（替换 ≤ 原文+2 字）+ 估算式排版预检 + 渲染 PNG 逐页视觉复验，双闸门缺一不可 |
| 校对时 AI 乱改原文引用、虚构数字 | 查实为用户资料逐字引用的一律不改（归入存疑清单）；所有数字/数据/模板逐字来自资料，不虚构 |
| 改完说不清改了什么、能不能回滚 | 副本先行（`backup_原因_日期.pptx`）+ 九节交付说明 + 按版本追加的"改动→红线→存疑→质检结论"修订审计 |

一句话定位：把"资料 → 课件 → 质检 → 交付"拆成可执行最小单元的中文 PPT 流水线。

## 主要功能

- **从零构建整套课件**（主线 A，8 步）：pptxgenjs 12 组件设计系统，设计令牌（配色/字体/篇章色/页脚）集中配置，16:9 原生可编辑输出
- **现有文件原位删改**（主线 B，6 步）：参数化删页、定点修复、用户改样范式传播，保护用户手动修改
- **序号重编与交叉引用清理**：单遍正则映射一次替换 + 无编号引用精确清单，页码按实际页序重排
- **通篇文案校对**（主线 C，5 步）：清单驱动替换，跨 run 保格式、替换次数断言、长度守卫；备注规范化/页脚统一/页码修正
- **双闸门质检**：程序化校验（零残留/页码一致/页脚统一/白名单豁免）+ 渲染 PNG 逐页视觉复验
- **交付说明与修订审计**：九节交付说明骨架（含覆盖核对表与历版修订记录），改动可追溯、备份可回滚

## 目录结构

```
PPTskill/
├── SKILL.md                  # 主文件：职责 / 触发场景 / 三条作业主线 / 输出标准 / 文件索引
├── README.md                 # 本文件
├── LICENSE                   # MIT License
├── references/               # 6 份详细规范（全部注明出处）
│   ├── 01_从零构建工作流.md   # A1–A9 全流程细则
│   ├── 02_幻灯片设计规范.md   # 深蓝学术风令牌全表 + 版式规则
│   ├── 03_pptxgenjs_api手册.md
│   ├── 04_编辑删改工作流.md   # 原位编辑铁律：fix_para 算法 / drop_rel / 单遍重编号 / 环境陷阱
│   ├── 05_中文文案与排版规范.md
│   └── 06_质检与验收标准.md   # 双闸门 + 视觉复验清单 + 修复循环
├── scripts/                  # 10 个全参数化脚本（路径走 argv，零硬编码）
│   ├── fix_ppr.py            # pptxgenjs 重复 <a:pPr> 兼容修复（构建后必跑）
│   ├── check_layout.py       # 溢出/越界/重叠估算式预检
│   ├── dump_pages.py         # 按页导出文本（表格/组合形状/备注）
│   ├── delete_slides.py      # 删页（自动高索引在前 + drop_rel）
│   ├── replace_text.py       # 清单驱动替换（跨 run 保格式 + 次数断言 + 长度守卫）
│   ├── renumber_text.py      # 单遍正则重编号 + TARGETS + 页码重排
│   ├── polish_notes.py       # 备注空格规范化 / 页脚统一 / 页码修正
│   ├── verify_deck.py        # 规则化程序校验（质检第一闸门）
│   ├── render_all.ps1        # COM 全量渲染 PNG
│   └── render_pages.ps1      # COM 抽页渲染（修复循环用）
└── assets/                   # 7 个可复用模板
    ├── lib-template.js       # 12 组件设计系统（令牌区集中配置）
    ├── build-template.js     # 分模块构建入口
    ├── deck-tokens.example.json
    ├── fixes清单.example.json
    ├── verify规则.example.json
    ├── 交付说明模板.md        # 九节交付说明骨架（含修订记录审计结构）
    └── image_map.example.json
```

## 安装方法

**方式一（个人技能目录，推荐）**：把整个 `PPTskill` 文件夹复制到 Agent 的技能目录，文件夹名可保持 `PPTskill`：

```
copy 到：C:\Users\<你>\.agents\skills\PPTskill\
```

**方式二（ZCode 插件技能目录）**：放入插件缓存对应的 skills 目录（结构同上）。

安装后重启会话，对 Agent 说"做 PPT / 改课件 / 校对全文"等即可触发（见 SKILL.md 触发场景表）。

## 依赖

| 用途 | 依赖 | 安装 |
|---|---|---|
| 原位编辑与校验 | Python 3.10+ | `pip install python-pptx` |
| 从零构建（主线 A） | Node.js | `npm install -g pptxgenjs` |
| 渲染质检（Windows） | PowerPoint 或 WPS | 已装 Office 即可（COM 调用） |

## 快速上手

**从零构建**（主线 A）：按 `SKILL.md` 主线 A 表格走 A1–A8；起步用 `assets/build-template.js` + `assets/lib-template.js`，每步细则看 `references/01`。

**修改现有文件**（主线 B）：先备份 → `dump_pages.py` 盘点 → `delete_slides.py` / `renumber_text.py` / `replace_text.py` 定点修改 → `verify_deck.py` 校验。

**通篇校对**（主线 C）：`dump_pages.py` 全文导出 → 按 `references/05` 规则产出清单 → `replace_text.py` + `polish_notes.py` 执行 → `verify_deck.py` + `render_pages.ps1` 双闸门验收。

## 输入输出示例

以排版预检脚本 `check_layout.py` 为例（构建/修改后、渲染前的第一道程序化闸门）。输入为最小演示文件：1 页 16:9 幻灯片、两个文本框（标题 + 正文），可用 python-pptx 临时构造（演示文件不入库）。

**通过**（正文单行，文本框 6″×1.0″，exit code 0）：

```text
$ python scripts/check_layout.py deck_pass.pptx
ALL PASS：无越界/溢出/重叠疑似项
```

**不通过**（同一文本框塞入约 192 个汉字，估算所需高度 1.84″ 超过实际 1.00″，exit code 2）：

```text
$ python scripts/check_layout.py deck_fail.pptx
发现 1 个疑似问题（估算式，渲染图为准）：
  S1 溢出?: need=1.84 h=1.00 text='这一段文字远超文本框容量：这一段文字远超'
```

脚本负责把疑似问题拦在渲染之前；最终判定以渲染 PNG 逐页视觉复验为准（双闸门中的第二道）。

## 内容出处

本技能由以下来源融合升级，引用处均已在文件内注明：

1. **2026-09-17 zcode 会话实战沉淀**（本技能主体）：69 页中文教学课件《AI提示词工程：从新手到高手》全流程（构建→删页重编→通篇校对→双闸门质检→交付审计）的全部脚本与规范，去除项目硬编码后参数化；
2. **document-skills:pptx 0.1.4**（Z.AI，License: Proprietary）：设计最佳实践框架与 pptxgenjs API 要点（references/02、03），仅提炼改写并注明；
3. **sparanoid/chinese-copywriting-guidelines**（MIT License）：中文排版 15 条规则（references/05 第一节）；
4. **agentskills.io 技能规范 + writing-skills**：SKILL.md frontmatter 字段、触发式 description 写法、scripts/references/assets 分工原则。

## 许可证

- 本项目以 [MIT License](LICENSE) 开源：任何人可免费使用、复制、修改及再分发（含商用），唯一义务是在副本中保留版权声明与本许可文本
- 本技能按"原样"提供，不附带任何明示或默示的担保（详见 LICENSE 全文）
- 引用来源兼容性：chinese-copywriting-guidelines 为 MIT 同协议引用；document-skills:pptx 为专有许可，本项目仅提炼改写其公开的最佳实践框架并已注明出处，未复制其代码或文本

## 版本

- v2.0.0（2026-09-17）：以官方 document-skills:pptx 为基线，融合实战工作流完成系统性升级：新增 references×6、scripts×10、assets×7，确立双闸门质检与修订审计交付标准。
