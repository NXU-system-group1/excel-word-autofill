# -*- coding: utf-8 -*-
"""各文档的填充器。每个函数返回 (输出路径, 命中, 未命中)。"""
import copy
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.shared import RGBColor
from docx.text.run import Run

sys.path.insert(0, str(Path(__file__).parent))

from core import (  # noqa: E402
    BASE, BIDDER, COMPANIES, DOC_DIR, TMP_DIR,
    apply_settlement, load_owner_info, load_red_info, load_settlement,
    merge_maps, replace_in_doc, set_cell, uniq_cells,
)
from doc2docx import convert as doc_to_docx  # noqa: E402

DECL_DIR = DOC_DIR / "1.声明"
DECL_OUT = DECL_DIR / "已填"

# 汇总表把联系人写成「路人甲/乙/丙」
SUMMARY_ALIAS = {"路人甲": "晨早", "路人乙": "泉旭", "路人丙": "左右"}

_SETTLE = None
_BIDDER = BIDDER


def set_bidder(name):
    """切换中标方（GUI 用）。"""
    global _BIDDER
    if name not in COMPANIES:
        raise ValueError(f"未知公司：{name}")
    _BIDDER = name


def reset_settle():
    """结算清单路径变了之后调用，清掉缓存。"""
    global _SETTLE
    _SETTLE = None


def settle():
    """结算清单（明细行, 合计金额, 合计大写），懒加载一次。价格唯一来源。"""
    global _SETTLE
    if _SETTLE is None:
        _SETTLE = load_settlement()
    return _SETTLE


def _append_blank(para, text="空"):
    """在段落末尾追加文字，继承格式但强制黑色（原占位可能是红的）。"""
    if not para.runs:
        r = para.add_run(text)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return
    last = para.runs[-1]
    new_el = copy.deepcopy(last._element)
    last._element.addnext(new_el)
    r = Run(new_el, para)
    r.text = text
    r.font.color.rgb = RGBColor(0, 0, 0)


def _ensure_docx(src, tmp_name):
    """老 .doc 先转成 .docx，缓存到 .tmp。"""
    if src.suffix.lower() == ".doc":
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        tmp = TMP_DIR / tmp_name
        if not tmp.exists() or tmp.stat().st_mtime < src.stat().st_mtime:
            doc_to_docx(src, tmp)
        return tmp
    return src


def fill_declarations():
    """1.声明：三份资格声明，按文件名匹配公司。"""
    DECL_OUT.mkdir(parents=True, exist_ok=True)
    results = []
    for src in sorted(DECL_DIR.glob("*.docx")):
        if src.name.startswith("~$"):
            continue
        dst = DECL_OUT / src.name
        company = next((c for c in COMPANIES if c in src.name), None)

        if company is None:
            shutil.copy2(src, dst)
            results.append((dst, [], ["（无对应红字信息，原样复制）"]))
            continue

        by_ph, _ = load_red_info(company)
        owner_ph, _ = load_owner_info()
        _, amt, cn = settle()
        doc = Document(str(src))
        hits, miss = replace_in_doc(
            doc, apply_settlement(merge_maps(owner_ph, by_ph), amt, cn))
        doc.save(str(dst))
        results.append((dst, hits, miss))
    return results


def fill_summary():
    """4、询价汇总表：三家信息合并，晨早最后覆盖「某某公司」。"""
    src = DOC_DIR / "4、询价汇总表.doc"
    dst = DOC_DIR / "4、询价汇总表_已填.docx"
    docx = _ensure_docx(src, "询价汇总表.docx")

    maps = []
    for c in ["泉旭", "左右", _BIDDER]:      # 中标方放最后，覆盖「某某公司」
        by_ph, _ = load_red_info(c)
        maps.append(by_ph)
    owner_ph, _ = load_owner_info()
    # 汇总表要保留三家各自的报价，不能统一成结算价，所以不套 apply_settlement
    merged = merge_maps(owner_ph, *maps)

    # 路人甲/乙/丙 → 各公司联系人
    for alias, company in SUMMARY_ALIAS.items():
        _, by_field = load_red_info(company)
        name = by_field.get("乙方联系人", "")
        if name:
            merged[alias] = name

    doc = Document(str(docx))
    hits, miss = replace_in_doc(doc, merged)
    doc.save(str(dst))
    return dst, hits, miss


def fill_notice():
    """5、成交通知书：乙方字段 + 电话别名。"""
    src = DOC_DIR / "5、成交通知书.docx"
    dst = DOC_DIR / "5、成交通知书_已填.docx"

    by_ph, by_field = load_red_info(_BIDDER)
    owner_ph, _ = load_owner_info()
    _, amt, cn = settle()
    merged = apply_settlement(merge_maps(owner_ph, by_ph), amt, cn)
    tel = by_field.get("乙方电话") or by_field.get("联系电话")
    if tel:
        merged["12345678"] = tel        # 模板电话占位是 12345678

    doc = Document(str(src))
    hits, miss = replace_in_doc(doc, merged)

    # 采购单位行的「联系人 / 联系电话」是甲方信息，未知就填「空」
    # （这两格模板里也有红字，会被当乙方联系人替换掉，所以必须事后覆盖）
    for row in doc.tables[0].rows:
        cells = uniq_cells(row)
        if cells and "采购单位" in cells[0].text:
            for i, c in enumerate(cells):
                if c.text.strip() in ("联系人", "联系电话") and i + 1 < len(cells):
                    set_cell(cells[i + 1], "空")

    doc.save(str(dst))
    return dst, hits, miss


def fill_contract():
    """6、合同：段10 占位写错（法定代表人栏写成「联系人」），段11 要「姓名 电话」。"""
    src = DOC_DIR / "6、合同.doc"
    dst = DOC_DIR / "6、合同_已填.docx"
    docx = _ensure_docx(src, "合同.docx")

    by_ph, by_field = load_red_info(_BIDDER)
    owner_ph, _ = load_owner_info()
    _, amt, cn = settle()
    merged = apply_settlement(merge_maps(owner_ph, by_ph), amt, cn)

    def resolver(para_text, red_text, mapping):
        if "法定代表人" in para_text and red_text == "联系人":
            return by_field.get("乙方法定代表人")
        if "联系人及联系电话" in para_text and red_text == "联系电话":
            combo = f"{by_field.get('乙方联系人', '')} {by_field.get('乙方电话', '')}".strip()
            return combo or None
        return mapping.get(red_text)

    doc = Document(str(docx))
    hits, miss = replace_in_doc(doc, merged, resolver)

    # 甲方「负责人：」「联系人及联系电话：」是黑字空栏，模板没占位，补「空」
    for p in doc.paragraphs:
        if p.text.strip() in ("负责人：", "联系人及联系电话："):
            _append_blank(p)

    doc.save(str(dst))
    return dst, hits, miss


def fill_receipt():
    """7、签收单。"""
    src = DOC_DIR / "7、签收单.docx12.docx"
    dst = DOC_DIR / "7、签收单_已填.docx"

    by_ph, _ = load_red_info(_BIDDER)
    owner_ph, _ = load_owner_info()
    _, amt, cn = settle()
    doc = Document(str(src))
    hits, miss = replace_in_doc(
        doc, apply_settlement(merge_maps(owner_ph, by_ph), amt, cn))
    doc.save(str(dst))
    return dst, hits, miss


def fill_acceptance():
    """8、验收书。"""
    src = DOC_DIR / "8、验收书.docx"
    dst = DOC_DIR / "8、验收书_已填.docx"

    by_ph, _ = load_red_info(_BIDDER)
    owner_ph, _ = load_owner_info()
    _, amt, cn = settle()
    doc = Document(str(src))
    hits, miss = replace_in_doc(
        doc, apply_settlement(merge_maps(owner_ph, by_ph), amt, cn))
    doc.save(str(dst))
    return dst, hits, miss


def sync_quote_tables():
    """把结算清单的明细同步进 6/7/8 的表格（行数自动增删）。价格与明细同源。"""
    from sync_quote_table import sync_one

    rows, amt, cn = settle()
    targets = [
        DOC_DIR / "6、合同_已填.docx",
        DOC_DIR / "7、签收单_已填.docx",
        DOC_DIR / "8、验收书_已填.docx",
    ]
    out = []
    for t in targets:
        if not t.exists():
            out.append((t, False, "文件不存在"))
            continue
        ok = sync_one(t, rows, amt, cn)
        out.append((t, ok, f"{len(rows)} 行 / {amt:.2f}"))
    return out
