# -*- coding: utf-8 -*-
"""replace_text.py — 清单驱动的精确文本替换（跨 run 保格式 + 次数断言）。

清单 JSON 格式（见 assets/fixes清单.example.json）：
{
  "fixes":  [ {"page": 5, "old": "复杂任务要求", "new": "提出复杂任务", "want": 1}, ... ],
  "notes":  [ {"page": 57, "old": "练习 1、5、6为笔记原题", "new": "练习 1、3、4 为笔记原题"} ],
  "splits": [ {"page": 31, "box_contains": "大量使用抽象词",
               "first_becomes": "大量使用抽象词",
               "new_para_text": "喜欢写宽泛总结"} ]
}

核心保证：
- 优先 run 内替换；跨 run 时按 span 拼接，命中边界之外的 run 原样保留（格式零破坏）；
- 每条断言替换次数，实际 != 期望立即报错定位；
- 长度守卫：len(new) > len(old) + 2 的条目拒绝执行（防溢出，可在 JSON 加 "allow_overflow": true 豁免）。

用法：
    python replace_text.py <文件.pptx> <清单.json>

来源：2026-09-17 zcode 会话 deck/polish_v4.py 的 fix_para 算法泛化版。
"""
import copy
import json
import sys

from pptx import Presentation

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


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


def fix_para(para, old, new):
    """run 内直接替换；跨 run 按 span 拼接，边界外 run 原样保留。返回替换次数。"""
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return 0
    for r in para.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            return 1
    s = full.index(old)
    e = s + len(old)
    pos, spans = 0, []
    for r in para.runs:
        spans.append((r, pos, pos + len(r.text)))
        pos += len(r.text)
    affected = [(r, a, b) for r, a, b in spans if b > s and a < e]
    for k, (r, a, b) in enumerate(affected):
        r.text = full[a:s] + new if k == 0 else full[max(e, a):b]
    return 1


def apply_body(slides, page, old, new, want):
    n = sum(fix_para(p, old, new)
            for tf in iter_tfs(slides[page - 1]) for p in tf.paragraphs)
    return n, want


def apply_notes(slides, page, old, new):
    sl = slides[page - 1]
    if not sl.has_notes_slide:
        return 0
    n = 0
    for para in sl.notes_slide.notes_text_frame.paragraphs:
        for r in para.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)
                n += 1
    return n


def apply_split(slides, spec):
    """段落拆条：深拷贝段落 XML 保格式，改克隆体文本后 addnext 插入。"""
    sl = slides[spec["page"] - 1]
    for tf in iter_tfs(sl):
        if spec["box_contains"] in tf.text:
            p0 = tf.paragraphs[0]
            p0.runs[0].text = spec["first_becomes"]
            newp = copy.deepcopy(p0._p)
            newp.find(A + "r").find(A + "t").text = spec["new_para_text"]
            p0._p.addnext(newp)
            return True
    return False


def main():
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    path, cfg_path = sys.argv[1], sys.argv[2]
    cfg = json.load(open(cfg_path, encoding="utf-8"))
    prs = Presentation(path)
    slides = list(prs.slides)
    errs, log = [], []

    def guard(page):
        if not (1 <= page <= len(slides)):
            errs.append("P%02d 页码越界（总页数 %d）" % (page, len(slides)))
            return False
        return True

    for f in cfg.get("fixes", []):
        if not guard(f["page"]):
            continue
        old, new = f["old"], f["new"]
        if len(new) > len(old) + 2 and not f.get("allow_overflow"):
            errs.append("P%02d 长度守卫: %r -> %r（新串超长，需拆方案或加 allow_overflow）"
                        % (f["page"], old[:20], new[:20]))
            continue
        n, want = apply_body(slides, f["page"], old, new, f.get("want", 1))
        ok = (n == want)
        if not ok:
            errs.append("P%02d !!期望%d实际%d %r" % (f["page"], want, n, old[:24]))
        log.append("P%02d %s %r -> %r" % (f["page"], "OK" if ok else "FAIL", old[:24], new[:24]))

    for f in cfg.get("notes", []):
        if not guard(f["page"]):
            continue
        n = apply_notes(slides, f["page"], f["old"], f["new"])
        if n == 0:
            errs.append("P%02d 备注未命中 %r" % (f["page"], f["old"][:24]))
        else:
            log.append("P%02d 备注修正 %r" % (f["page"], f["old"][:24]))

    for sp in cfg.get("splits", []):
        if not guard(sp["page"]):
            continue
        if not apply_split(slides, sp):
            errs.append("P%02d 拆条未命中 %r" % (sp["page"], sp["box_contains"][:20]))
        else:
            log.append("P%02d 拆条完成（首条=%r，新增=%r）"
                       % (sp["page"], sp["first_becomes"], sp["new_para_text"]))

    if errs:
        print("\n".join(log))
        print("\n=== FAILURES ===")
        print("\n".join(errs))
        print("未保存（清单需修正后重跑）")
        sys.exit(2)
    prs.save(path)
    print("SAVED")
    print("\n".join(log))


if __name__ == "__main__":
    main()
