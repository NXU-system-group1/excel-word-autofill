# -*- coding: utf-8 -*-
"""填充「4、询价汇总表」：合并三份红字信息，替换红色系标记，输出 docx，不动源文件。

合并顺序：泉旭 → 左右 → 晨早。晨早最后覆盖，保证「某某公司」取中标单位。
"""
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor

sys.path.insert(0, str(Path(__file__).parent))
from doc2docx import convert as doc_to_docx  # noqa: E402

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:/git学习/Excel转指定位置word")
BASE = ROOT / "项目文件"
INFO_DIR = BASE / "红字信息"
SRC_DOC = BASE / "1、工程A三方询价 价格" / "4、询价汇总表.doc"
TMP_DOCX = ROOT / ".tmp" / "询价汇总表.docx"
OUT_DOCX = BASE / "1、工程A三方询价 价格" / "4、询价汇总表_已填.docx"

# 晨早放最后，让它的「某某公司」覆盖另外两家
ORDER = ["泉旭", "左右", "晨早"]


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
    """金额字段统一两位小数；其余原样转字符串。"""
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


# 1. 合并映射
merged = {}
for name in ORDER:
    part = load_mapping(INFO_DIR / f"{name}.xlsx")
    merged.update(part)
    print(f"载入 {name}.xlsx：{len(part)} 条")

# 汇总表把联系人写成「路人甲/乙/丙」，红字信息里是「联系人」，按顺序补齐
merged.update({"路人甲": "罗艳军", "路人乙": "张何", "路人丙": "刘念"})
print("补充 路人甲/乙/丙 → 各公司联系人")

# 2. 确保 docx 存在
if not TMP_DOCX.exists() or TMP_DOCX.stat().st_mtime < SRC_DOC.stat().st_mtime:
    print("转换 .doc → .docx ...")
    doc_to_docx(SRC_DOC, TMP_DOCX)

# 3. 替换
doc = Document(str(TMP_DOCX))
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

OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
doc.save(str(OUT_DOCX))

print("\n" + "=" * 72)
print("已替换：")
for old, new in hits:
    print(f"  {old}  →  {new}")
if miss:
    print("\n未替换（红字信息中无此项）：")
    for t in sorted(set(miss)):
        print(f"  {t}")
print("\n输出：", OUT_DOCX)
