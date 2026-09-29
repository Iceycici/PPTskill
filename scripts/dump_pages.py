# -*- coding: utf-8 -*-
"""dump_pages.py — 按页导出 pptx 文本（含表格/组合形状/演讲者备注）。

用途：校对/删改前的全文盘点（B2/C1 步），逐页打印以便人工复核与定位。

用法：
    python dump_pages.py <文件.pptx> 5 8 9      # 只导第 5/8/9 页
    python dump_pages.py <文件.pptx> --all      # 全部页
    python dump_pages.py <文件.pptx> --notes    # 附带备注（默认附带）

输出格式：
    === S12 ===
    [T] 文本框/表格内容（表格按 "a | b | c" 一行）
    [N] 演讲者备注

来源：2026-09-17 zcode 会话 deck/dump_pages.py 泛化版。
"""
import sys

from pptx import Presentation


def walk(shapes, out):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            walk(sh.shapes, out)
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            for r in sh.table.rows:
                out.append(" | ".join(c.text for c in r.cells))


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(1)
    path = args[0]
    pages, show_notes = [], True
    for a in args[1:]:
        if a == "--all":
            pages = None
        elif a == "--no-notes":
            show_notes = False
        elif a.isdigit():
            pages.append(int(a))
    prs = Presentation(path)
    slides = list(prs.slides)
    if pages is None:
        pages = range(1, len(slides) + 1)
    for i in pages:
        sl = slides[i - 1]
        out = []
        walk(sl.shapes, out)
        print("=== S%d ===" % i)
        for t in out:
            print("[T] " + t.replace("\n", "\n    "))
        if show_notes and sl.has_notes_slide:
            n = sl.notes_slide.notes_text_frame.text.strip()
            if n:
                print("[N] " + n.replace("\n", "\n    "))
        print()


if __name__ == "__main__":
    main()
