# -*- coding: utf-8 -*-
"""填充「5、成交通知书」：乙方字段取自 晨早.xlsx，甲方与项目名在脚本内补，输出新文件不动源文件。"""
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:/git学习/Excel转指定位置word")
BASE = ROOT / "项目文件"
INFO = BASE / "红字信息" / "晨早.xlsx"
SRC = BASE / "1、工程A三方询价 价格" / "5、成交通知书.docx"
OUT = BASE / "1、工程A三方询价 价格" / "5、成交通知书_已填.docx"

PROJECT = "长沙市天心区青园小学百川楼打卡文化墙采购项目"

# 模板里的别名占位 → 红字信息里的规范占位
ALIAS = {
    "12345678": "联系电话",
}


def is_red(run):
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return False
    c = rpr.find(qn("w:color"))
    if c is None:
        return False
    val = c.get(qn("w:val"))
    if not val or len(val) != 6:
        return False
    try:
        r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
    except ValueError:
        return False
    return r >= 0x80 and g <= 0x60 and b <= 0x60


def red_blocks(para):
    out, buf = [], []
    for r in para.runs:
        if is_red(r):
            buf.append(r)
        elif buf:
            out.append(buf)
            buf = []
    if buf:
        out.append(buf)
    return out


def all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        seen = set()
        for row in t.rows:
            for cell in row.cells:
                k = cell._tc.getroottree().getpath(cell._tc)
                if k in seen:
                    continue
                seen.add(k)
                yield from cell.paragraphs
    for s in doc.sections:
        for part in (s.header, s.footer):
            yield from part.paragraphs


def fmt(val, field=""):
    if isinstance(val, (int, float)) and "金额" in str(field) and "大写" not in str(field):
        return f"{float(val):.2f}"
    if isinstance(val, float):
        return f"{val:.2f}"
    return str(val).strip()


def load_mapping(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    mapping = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 3:
            continue
        ph, field, val = row[0], row[1], row[2]
        if not ph or ph == "—" or val is None:
            continue
        mapping[str(ph).strip()] = fmt(val, field)
    wb.close()
    return mapping


merged = load_mapping(INFO)
print(f"载入 晨早.xlsx：{len(merged)} 条")

# 别名：模板占位 → 规范占位 → 值
for alias, canonical in ALIAS.items():
    if canonical in merged:
        merged[alias] = merged[canonical]

doc = Document(str(SRC))
hits, miss = [], []
for para in all_paragraphs(doc):
    for buf in red_blocks(para):
        text = "".join(r.text for r in buf).strip()
        if not text:
            continue
        if text in merged:
            buf[0].text = merged[text]
            buf[0].font.color.rgb = RGBColor(0, 0, 0)
            for r in buf[1:]:
                r.text = ""
            hits.append((text, merged[text]))
        else:
            miss.append(text)

OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(str(OUT))

print("\n" + "=" * 72)
print("已替换：")
for old, new in hits:
    print(f"  {old}  →  {new}")
if miss:
    print("\n未替换（映射表中无此项）：")
    for t in sorted(set(miss)):
        print(f"  {t}")
print("\n输出：", OUT)
