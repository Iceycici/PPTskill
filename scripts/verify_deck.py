# -*- coding: utf-8 -*-
"""verify_deck.py — 规则化程序校验器（质检第一闸门）。

规则 JSON 格式（见 assets/verify规则.example.json）：
{
  "expect_slides": 69,
  "residuals":   [ {"page": 5, "old": "复杂任务要求"}, ... ],   # 该页不得出现的旧片段
  "banned":      ["13 道", "12 个场景"],                        # 全文（含备注）禁用词
  "whitelist":   [ {"page": 44, "allow": ["头脑风暴会议"]} ],    # 豁免：合法残留
  "number_scans": [ {"regex": "场景\\\\s*(\\\\d+)", "max": 8, "name": "场景号"},
                    {"regex": "练习\\\\s*(\\\\d+)", "max": 9, "name": "练习号"} ],
  "kickers":     [ {"regex": "第二篇 · 场景 (\\d+)", "expect": [1,2,3,4,5,6,7,8]} ],
  "footer":      "AI提示词工程",                                # 页脚必须唯一为此形态
  "check_page_numbers": true,
  "must_contain": [ {"page": 35, "text": "请扮演严格编辑"} ]    # 模板卡逐字抽查（防误伤）
}

用法：
    python verify_deck.py <文件.pptx> <规则.json>

退出码：0 全过 / 2 有 FAIL（逐条列出）。
来源：2026-09-17 zcode 会话 deck/verify_v4.py + verify_final.py 规则合并泛化版。
"""
import json
import re
import sys

from pptx import Presentation

EMU = 914400.0


def iter_texts(sl):
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
                continue
            if sh.has_text_frame:
                yield sh.text_frame.text
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    yield " | ".join(c.text for c in row.cells)
    yield from walk(sl.shapes)


def main():
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    path, rule_path = sys.argv[1], sys.argv[2]
    R = json.load(open(rule_path, encoding="utf-8"))
    prs = Presentation(path)
    slides = list(prs.slides)
    errs = []

    # 1 页数
    if "expect_slides" in R and len(slides) != R["expect_slides"]:
        errs.append("页数=%d 期望=%d" % (len(slides), R["expect_slides"]))

    # 收集每页文本（正文+表格）与备注
    body = {i: "\n".join(iter_texts(sl)) for i, sl in enumerate(slides, 1)}
    notes = {i: (sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else "")
             for i, sl in enumerate(slides, 1)}

    # 2 旧片段零残留
    for r in R.get("residuals", []):
        if r["old"] in body.get(r["page"], "") or r["old"] in notes.get(r["page"], ""):
            errs.append("P%d 残留 %r" % (r["page"], r["old"]))

    # 3 禁用词 + 白名单
    wl = {}
    for w in R.get("whitelist", []):
        wl.setdefault(w["page"], set()).update(w["allow"])
    for ban in R.get("banned", []):
        for i in body:
            hit_pages = [pg for pg, txt in ((i, body[i]), (i, notes[i])) if ban in txt]
        hit = [i for i in body if ban in body[i] or ban in notes[i]]
        for i in hit:
            allows = wl.get(i, set())
            if not any(a in ban or ban in a for a in allows):
                errs.append("P%d 禁用词 %r" % (i, ban))

    # 4 编号上界扫描
    for ns in R.get("number_scans", []):
        rx = re.compile(ns["regex"])
        mx = ns.get("max")
        for i in body:
            for m in rx.finditer(body[i] + notes[i]):
                v = int(m.group(1))
                if mx is not None and v > mx:
                    errs.append("P%d %s 超界: %d (max=%d) 上下文 %r"
                                % (i, ns["name"], v, mx,
                                   (body[i] + notes[i])[max(0, m.start()-10):m.end()+10]))

    # 5 kicker 连续性
    for k in R.get("kickers", []):
        rx = re.compile(k["regex"])
        seq = sorted({int(m.group(1)) for i in body for m in [rx.search(body[i])] if m})
        if "expect" in k and seq != k["expect"]:
            errs.append("kicker 序列 %s != 期望 %s (%s)" % (seq, k["expect"], k["regex"]))

    # 6 页脚唯一形态 + 7 页码=页序
    foots = {}
    for i, sl in enumerate(slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame or sh.top is None:
                continue
            top, left = sh.top / EMU, (sh.left or 0) / EMU
            t = sh.text_frame.text.strip()
            if top > 7.0 and left < 2.0 and t and R.get("footer"):
                foots[t] = foots.get(t, 0) + 1
            if R.get("check_page_numbers") and top > 7.0 and left > 12.0 and t.isdigit():
                if int(t) != i:
                    errs.append("P%d 页码显示 %s" % (i, t))
    if R.get("footer"):
        unexpected = {k: v for k, v in foots.items() if k != R["footer"]}
        if unexpected:
            errs.append("页脚存在非目标形态: %s" % unexpected)

    # 8 模板卡逐字抽查
    for m in R.get("must_contain", []):
        if m["text"] not in body.get(m["page"], ""):
            errs.append("P%d 模板抽查缺失 %r" % (m["page"], m["text"][:20]))

    if errs:
        print("=== FAILURES (%d) ===" % len(errs))
        print("\n".join(errs))
        sys.exit(2)
    print("ALL PASS（%d 页；页脚形态 %s）" % (len(slides), list(foots) or "未校验"))


if __name__ == "__main__":
    main()
