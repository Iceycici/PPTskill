# -*- coding: utf-8 -*-
"""delete_slides.py — 按页码删页（自动高索引在前 + drop_rel 防孤儿关系）。

用法：
    python delete_slides.py <文件.pptx> 31 32 33 72 81
    python delete_slides.py <文件.pptx> 31-36,59-63,80,81

安全提示：
    - 执行前务必备份（backup_原因_日期.pptx）；
    - 页码为 1-based 物理页序（与页码框显示可能不一致，删前用 dump_pages.py 确认）；
    - 脚本自动降序删除，传入顺序随意。

来源：2026-09-17 zcode 会话 deck/apply_delete.py 泛化版。
"""
import re
import sys

from pptx import Presentation
from pptx.oxml.ns import qn


def parse_pages(tokens):
    pages = []
    for t in tokens:
        for part in t.split(","):
            if "-" in part:
                a, b = part.split("-")
                pages.extend(range(int(a), int(b) + 1))
            elif part.strip():
                pages.append(int(part))
    return sorted(set(pages), reverse=True)   # 关键：高索引在前


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    path, tokens = sys.argv[1], sys.argv[2:]
    pages = parse_pages(tokens)
    total_before = None
    prs = Presentation(path)
    total_before = len(prs.slides._sldIdLst)
    if pages and pages[0] > total_before:
        print("!! 页码 %d 超出总页数 %d" % (pages[0], total_before)); sys.exit(1)
    for idx in pages:
        sld = list(prs.slides._sldIdLst)[idx - 1]
        prs.part.drop_rel(sld.get(qn("r:id")))
        prs.slides._sldIdLst.remove(sld)
        print("deleted slide %d" % idx)
    prs.save(path)
    print("SAVED: %d -> %d slides" % (total_before, len(prs.slides._sldIdLst)))


if __name__ == "__main__":
    main()
