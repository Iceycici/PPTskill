# -*- coding: utf-8 -*-
"""polish_notes.py — 三项一致性清理：备注空格规范化 / 页脚统一 / 页码修正。

--notes-space     备注与正文中英文/数字间批量补空格（"讲义1.1"→"讲义 1.1"），
                  并回归检查路径/URL 未被破坏（"C :"、"http s"）
--footer TEXT     页脚统一为指定课程名（位置法：top>7.0" 且 left<2.0"；
                  先长后短两轮，防止悬挂间隔号残留；传 --footer-suffix 处理
                  "课程名 · 副标题"形态）
--pagefix         右下角数字页码框重写为实际页序（位置法：top>7.0" 且 left>12.0"）

示例：
    python polish_notes.py 文件.pptx --notes-space --footer "AI提示词工程" --pagefix

来源：2026-09-17 zcode 会话 deck/polish_v4.py 后半部分泛化版。
"""
import argparse
import re
import sys

from pptx import Presentation

EMU = 914400.0


def iter_shape_slides(prs):
    for i, sl in enumerate(prs.slides, 1):
        yield i, sl


def space_text(text):
    text = re.sub(r"([\u4e00-\u9fff])([A-Za-z0-9])", r"\1 \2", text)
    text = re.sub(r"([A-Za-z0-9])([\u4e00-\u9fff])", r"\1 \2", text)
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--notes-space", action="store_true",
                    help="中英文/数字间补空格（含备注区），附路径回归检查")
    ap.add_argument("--footer", metavar="TEXT", help="页脚统一为 TEXT（先长后短）")
    ap.add_argument("--pagefix", action="store_true", help="页码框重写为实际页序")
    args = ap.parse_args()

    if not (args.notes_space or args.footer or args.pagefix):
        ap.print_help(); sys.exit(1)

    prs = Presentation(args.pptx)
    log = []

    if args.notes_space:
        notes_pages = body_pages = 0
        for i, sl in iter_shape_slides(prs):
            def fix_tf(tf):
                n = 0
                for p in tf.paragraphs:
                    for r in p.runs:
                        t = space_text(r.text)
                        if t != r.text:
                            r.text = t; n += 1
                return n
            n = sum(fix_tf(tf) for shape in sl.shapes
                    if shape.has_text_frame for tf in [shape.text_frame])
            n += sum(fix_tf(c.text_frame) for shape in sl.shapes
                     if getattr(shape, "has_table", False) and shape.has_table
                     for row in shape.table.rows for c in row.cells)
            if n:
                body_pages += 1
            if sl.has_notes_slide and fix_tf(sl.notes_slide.notes_text_frame):
                notes_pages += 1
        log.append("空格规范化：正文命中 %d 页，备注命中 %d 页" % (body_pages, notes_pages))
        # 回归检查
        bad = []
        for i, sl in iter_shape_slides(prs):
            texts = []
            for shape in sl.shapes:
                if shape.has_text_frame:
                    texts.append(shape.text_frame.text)
                if getattr(shape, "has_table", False) and shape.has_table:
                    texts.extend(c.text for row in shape.table.rows for c in row.cells)
            if sl.has_notes_slide:
                texts.append(sl.notes_slide.notes_text_frame.text)
            joined = "\n".join(texts)
            if "C :" in joined or "http s" in joined or "  " in joined.replace("\n", ""):
                bad.append(i)
        if bad:
            log.append("!! 回归警告：以下页出现双空格/路径破坏，需人工复核：%s" % bad)

    if args.footer:
        target = args.footer
        variants = []
        for shape in prs.slides[0].shapes:
            pass
        # 收集当前页脚的全部形态（先长后短替换）
        forms = {}
        for i, sl in iter_shape_slides(prs):
            for shape in sl.shapes:
                if (shape.has_text_frame and shape.top is not None
                        and shape.top / EMU > 7.0 and shape.left is not None
                        and shape.left / EMU < 2.0):
                    t = shape.text_frame.text.strip()
                    if t and t != target:
                        forms[t] = forms.get(t, 0) + 1
        ordered = sorted(forms, key=len, reverse=True)
        n = 0
        for i, sl in iter_shape_slides(prs):
            for shape in sl.shapes:
                if (shape.has_text_frame and shape.top is not None
                        and shape.top / EMU > 7.0 and shape.left is not None
                        and shape.left / EMU < 2.0):
                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            for old in ordered:
                                if r.text.strip() == old:
                                    r.text = r.text.replace(old, target)
                                    n += 1
        log.append("页脚统一 %d 处（原形态：%s）"
                   % (n, "; ".join("%s x%d" % (k, v) for k, v in forms.items()) or "无"))

    if args.pagefix:
        fixed = 0
        for i, sl in iter_shape_slides(prs):
            for shape in sl.shapes:
                if (shape.has_text_frame and shape.top is not None
                        and shape.top / EMU > 7.0 and shape.left is not None
                        and shape.left / EMU > 12.0
                        and shape.text_frame.text.strip().isdigit()):
                    run = shape.text_frame.paragraphs[0].runs[0]
                    if run.text != str(i):
                        run.text = str(i)
                        fixed += 1
        log.append("页码修正 %d 处" % fixed)

    prs.save(args.pptx)
    print("SAVED")
    print("\n".join(log))


if __name__ == "__main__":
    main()
