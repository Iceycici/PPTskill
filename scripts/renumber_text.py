# -*- coding: utf-8 -*-
"""renumber_text.py — 删页后全篇重编号：单遍正则映射 + TARGETS 精确替换 + 页码重排。

配置 JSON 格式（见 assets/fixes清单.example.json 同目录约定）：
{
  "patterns": [                                  # 单遍正则重编号（一次遍历完成，防级联）
    {"regex": "场景\\s*(\\d+)", "map": {"5": 3, "6": 4, "7": 5, "8": 6, "9": 7, "12": 8},
     "template": "场景 {new}"}
  ],
  "targets": [                                   # 无编号交叉引用/计数的精确子串替换
    {"page": 3, "old": "课后练习（13 题）", "new": "课后练习（9 题）", "want": 1}
  ],
  "renumber_footers": true                       # 按位置重排右下角页码框为实际页序
}

要点：
- patterns 在同一遍扫描内完成全部映射（先 5→3 不会再被 3→1 二次替换）；
- 映射外的编号原样保留；
- 应用范围含表格与演讲备注；
- 页码框按位置定位（top>7.0" 且 left>12.0" 且纯数字），1 英寸 = 914400 EMU。

用法：
    python renumber_text.py <文件.pptx> <配置.json>

来源：2026-09-17 zcode 会话 deck/apply_renumber.py 泛化版。
"""
import json
import re
import sys

from pptx import Presentation

EMU = 914400.0


def iter_tfs(sl):
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
                continue
            if sh.has_text_frame:
                yield sh.text_frame
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        yield c.text_frame
    yield from walk(sl.shapes)


def make_replacer(spec):
    rx = re.compile(spec["regex"])
    mapping = {int(k): int(v) for k, v in spec.get("map", {}).items()}
    template = spec.get("template")

    def sub(m):
        n = int(m.group(1))
        new = mapping.get(n, n)
        return template.format(new=new) if template else str(new)
    return rx, sub


def main():
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    path, cfg_path = sys.argv[1], sys.argv[2]
    cfg = json.load(open(cfg_path, encoding="utf-8"))
    prs = Presentation(path)
    slides = list(prs.slides)
    subs = [make_replacer(p) for p in cfg.get("patterns", [])]
    log, errs = [], []

    def renum(text):
        for rx, sub in subs:
            text = rx.sub(sub, text)
        return text

    # 1) 全页正则重编（正文+表格+备注）
    touched = 0
    for i, sl in enumerate(slides, 1):
        changed = False
        for tf in iter_tfs(sl):
            for p in tf.paragraphs:
                for r in p.runs:
                    t = renum(r.text)
                    if t != r.text:
                        r.text = t
                        changed = True
        if sl.has_notes_slide:
            for p in sl.notes_slide.notes_text_frame.paragraphs:
                for r in p.runs:
                    t = renum(r.text)
                    if t != r.text:
                        r.text = t
                        changed = True
        if changed:
            touched += 1
            log.append("P%02d 正则重编命中" % i)

    # 2) TARGETS 精确替换
    from replace_text import iter_tfs as _it, fix_para  # 复用同目录算法
    for t in cfg.get("targets", []):
        n = sum(fix_para(p, t["old"], t["new"])
                for tf in _it(slides[t["page"] - 1]) for p in tf.paragraphs)
        want = t.get("want", 1)
        if n != want:
            errs.append("P%02d !!期望%d实际%d %r" % (t["page"], want, n, t["old"][:24]))
        else:
            log.append("P%02d TARGET OK %r -> %r" % (t["page"], t["old"][:20], t["new"][:20]))

    # 3) 页码框按位置重排
    if cfg.get("renumber_footers"):
        fixed = 0
        for i, sl in enumerate(slides, 1):
            for sh in sl.shapes:
                if (sh.has_text_frame and sh.top is not None and sh.left is not None
                        and sh.top / EMU > 7.0 and sh.left / EMU > 12.0
                        and sh.text_frame.text.strip().isdigit()):
                    run = sh.text_frame.paragraphs[0].runs[0]
                    if run.text != str(i):
                        run.text = str(i)
                        fixed += 1
        log.append("页码重排 %d 处" % fixed)

    if errs:
        print("\n".join(log + errs))
        print("=== 存在未命中项，未保存 ===")
        sys.exit(2)
    prs.save(path)
    print("SAVED（命中页 %d）" % touched)
    print("\n".join(log))


if __name__ == "__main__":
    main()
