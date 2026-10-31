# -*- coding: utf-8 -*-
"""体检 docx：按文档顺序导出段落、表格矩阵、每个 run 的颜色，并统计颜色分布。

用法：
    python tools/inspect_docx.py "5、成交通知书.docx"
    python tools/inspect_docx.py "5、成交通知书.docx" --quiet   # 只输出非黑颜色与红字

红色系判定：R>=0x80 且 G<=0x60 且 B<=0x60，可同时覆盖 EE0000 / FF0000 / C00000。
"""
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")


def color_of(run):
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return None
    c = rpr.find(qn("w:color"))
    v = c.get(qn("w:val")) if c is not None else None
    return v.upper() if v else None


def is_red(val):
    if not val or len(val) != 6:
        return False
    try:
        r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
    except ValueError:
        return False
    return r >= 0x80 and g <= 0x60 and b <= 0x60


def iter_blocks(doc):
    """按文档顺序产出 ('p', Paragraph) 或 ('t', Table)。"""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "t", Table(child, doc)


def unique_cells(table):
    seen = set()
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            k = cell._tc.getroottree().getpath(cell._tc)
            if k in seen:
                yield ri, ci, cell, True
                continue
            seen.add(k)
            yield ri, ci, cell, False


def inspect(path, quiet=False):
    doc = Document(str(path))
    print("=" * 76)
    print("文件：", Path(path).name)
    print(f"段落 {len(doc.paragraphs)} | 表格 {len(doc.tables)}")
    print("=" * 76)

    colors = Counter()
    pidx = 0
    for kind, block in iter_blocks(doc):
        if kind == "p":
            if not block.runs:
                continue
            pidx += 1
            text = "".join(r.text for r in block.runs)
            if not text.strip():
                continue
            marks = []
            for i, r in enumerate(block.runs):
                col = color_of(r)
                if col:
                    colors[col] += 1
                if col and not quiet:
                    flag = "  ←红" if is_red(col) else ""
                    marks.append(f"      run{i} [{col}]{flag} <{r.text}>")
                elif col and quiet and is_red(col):
                    marks.append(f"      run{i} [{col}] <{r.text}>")
            if quiet and not any(is_red(color_of(r)) for r in block.runs):
                continue
            print(f"\n[段{pidx}] {text}")
            if marks:
                print("\n".join(marks))
        else:
            seen = set()
            print(f"\n[表格 {len(block.rows)} 行 × {len(block.columns)} 列]")
            for ri, row in enumerate(block.rows):
                cells = []
                for ci, cell in enumerate(row.cells):
                    k = cell._tc.getroottree().getpath(cell._tc)
                    if k in seen:
                        cells.append("<合并>")
                        continue
                    seen.add(k)
                    txt = " / ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    reds = []
                    for p in cell.paragraphs:
                        for r in p.runs:
                            col = color_of(r)
                            if col:
                                colors[col] += 1
                                if is_red(col) and r.text.strip():
                                    reds.append(r.text.strip())
                    mark = f"  «红:{'|'.join(reds)}»" if reds else ""
                    cells.append(f"{txt}{mark}")
                print(f"  R{ri}: " + " || ".join(cells))

    print("\n" + "=" * 76)
    print("颜色分布：")
    if not colors:
        print("  无非默认颜色")
    for col, n in colors.most_common():
        tag = "  ←红色系" if is_red(col) else ""
        print(f"  {col}: {n} 个 run{tag}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    inspect(sys.argv[1], quiet="--quiet" in sys.argv)
