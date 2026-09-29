# -*- coding: utf-8 -*-
"""check_layout.py — 构建后程序化排版预检：文本溢出 / 越界 / 重叠。

原理（估算式预检，渲染前先拦一遍）：
- CJK 字符按 1em、ASCII 按 0.55em 估算字符宽，模拟自动折行算所需高度；
  need > h * 1.08 判溢出；
- 形状超出画布（13.33 × 7.5 英寸）判越界；
- 两个文本形状交叠超过 0.25" × 0.18" 判重叠。

已知良性误报（人工豁免）：篇章页大号半透明水印数字（有意设计）。

用法：
    python check_layout.py <文件.pptx>

来源：2026-09-17 zcode 会话 deck/check_layout.py 泛化版。
"""
import sys

from pptx import Presentation

W, H = 13.333, 7.5
EMU = 914400.0
OX, OY = 0.25, 0.18          # 重叠判定阈值（英寸）
TOL = 1.08                   # 溢出容差


def est_width(ch):
    return 1.0 if ord(ch) > 0x2E80 else 0.55


def lines_needed(text, w_in, font_pt):
    if not text or w_in <= 0 or font_pt <= 0:
        return 0
    em = font_pt / 72.0
    max_chars = max(1.0, w_in / em)
    n = 0
    for raw in text.split("\n"):
        seg = raw or " "
        width = sum(est_width(c) for c in seg)
        n += max(1, int(width / max_chars) + (1 if width % max_chars else 0))
    return n


def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def main():
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    prs = Presentation(sys.argv[1])
    problems = []
    for i, sl in enumerate(prs.slides, 1):
        boxes = []
        for sh in iter_shapes(sl.shapes):
            if sh.top is None or sh.left is None or sh.width is None or sh.height is None:
                continue
            x, y = sh.left / EMU, sh.top / EMU
            w, h = sh.width / EMU, sh.height / EMU
            if x < -0.05 or y < -0.05 or x + w > W + 0.05 or y + h > H + 0.05:
                problems.append("S%d 越界: (%.2f,%.2f,%.2f,%.2f)" % (i, x, y, w, h))
            if sh.has_text_frame and sh.text_frame.text.strip():
                t = sh.text_frame.text
                sizes = [r.font.size.pt for p in sh.text_frame.paragraphs
                         for r in p.runs if r.font.size is not None]
                pt = max(sizes) if sizes else 12
                para_txt = "\n".join("".join(r.text for r in p.runs) for p in sh.text_frame.paragraphs)
                need = lines_needed(para_txt, max(w - 0.1, 0.1), pt) * pt / 72.0 * 1.35
                if need > h * TOL + 0.06:
                    problems.append("S%d 溢出?: need=%.2f h=%.2f text=%r" % (i, need, h, t[:20]))
                boxes.append((x, y, w, h, t[:14]))
        for a in range(len(boxes)):
            for b in range(a + 1, len(boxes)):
                x1, y1, w1, h1, t1 = boxes[a]
                x2, y2, w2, h2, t2 = boxes[b]
                ox = min(x1 + w1, x2 + w2) - max(x1, x2)
                oy = min(y1 + h1, y2 + h2) - max(y1, y2)
                if ox > OX and oy > OY:
                    problems.append("S%d 重叠?: %r x %r (%.2f x %.2f)" % (i, t1, t2, ox, oy))
    if problems:
        print("发现 %d 个疑似问题（估算式，渲染图为准）：" % len(problems))
        for p in problems:
            print(" ", p)
        sys.exit(2)
    print("ALL PASS：无越界/溢出/重叠疑似项")


if __name__ == "__main__":
    main()
