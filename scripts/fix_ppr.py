# -*- coding: utf-8 -*-
"""fix_ppr.py — pptxgenjs 兼容修复：清除富文本段落中重复的 <a:pPr>。

背景：pptxgenjs 在 run 级富文本段落中会为每个 run 输出一个 <a:pPr>，
PowerPoint 打开时报"发现无法读取的内容，需要修复"。本脚本在每个
<a:p>...</a:p> 块内只保留第一个 pPr（连同其前缀），其余删除。

用法：
    python fix_ppr.py <文件.pptx> [更多.pptx ...]

来源：2026-09-17 zcode 会话 deck/fix_ppr.py 泛化版。
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

PPR = re.compile(r"<a:pPr[^>]*(?:/>|>[\s\S]*?</a:pPr>)")


def fix_para(m):
    block = m.group(0)
    fixed, count = PPR.subn("", block, count=0)
    # 保留第一个 pPr：先找出第一个匹配的跨度
    first = PPR.search(block)
    if not first:
        return block
    out = block[:first.end()]                      # 头部 + 第一个 pPr
    rest = block[first.end():]
    out += PPR.sub("", rest)                       # 其余 pPr 全删
    return out


def fix_xml(xml):
    return re.sub(r"<a:p>[\s\S]*?</a:p>", fix_para, xml)


def process(path):
    fd, tmp = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    changed = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                new = fix_xml(data.decode("utf-8"))
                if new != data.decode("utf-8"):
                    changed += 1
                data = new.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return changed


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        n = process(p)
        print("%s: fixed %d slide(s)" % (p, n))


if __name__ == "__main__":
    main()
