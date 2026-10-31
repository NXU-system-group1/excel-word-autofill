# -*- coding: utf-8 -*-
"""一键填充：附件 1.声明（三家）+ 附件 4~8 全部文档。

用法：
    python start.py           # 全量执行
    python start.py --dry     # 只检查前置文件，不写任何东西

规则：
  - 只替换红字信息表里存在的乙方字段，甲方一律不动
  - 6/7/8 的明细表按「中标方报价单」同步，行数随报价单自动增删
  - 全部输出到新文件，源文件一个字节都不动
"""
import sys
from pathlib import Path

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "tools"))

from core import BIDDER, COMPANIES, DOC_DIR, INFO_DIR, QUOTE_XLSX, SETTLE_XLSX  # noqa: E402
import fillers  # noqa: E402

SOURCES = [
    ("1.声明（三份）", DOC_DIR / "1.声明"),
    ("4、询价汇总表", DOC_DIR / "4、询价汇总表.doc"),
    ("5、成交通知书", DOC_DIR / "5、成交通知书.docx"),
    ("6、合同", DOC_DIR / "6、合同.doc"),
    ("7、签收单", DOC_DIR / "7、签收单.docx12.docx"),
    ("8、验收书", DOC_DIR / "8、验收书.docx"),
]


def check_inputs():
    missing = []
    for c in COMPANIES:
        p = INFO_DIR / f"{c}.xlsx"
        if not p.exists():
            missing.append(f"红字信息：{p.name}")
    if not QUOTE_XLSX.exists():
        missing.append(f"报价单：{QUOTE_XLSX.name}")
    for label, p in SOURCES:
        if not p.exists():
            missing.append(f"{label}：{p}")
        elif p.is_dir() and not list(p.glob("*.docx")):
            missing.append(f"{label}：目录里没有 docx")
    return missing


def show_hits(tag, hits, miss, indent="    "):
    if hits:
        for old, new in hits:
            print(f"{indent}{old}  →  {new}")
    if miss:
        uniq = sorted(set(miss))
        print(f"{indent}未替换（红字信息中无此项）：{'、'.join(uniq)}")


def main():
    dry = "--dry" in sys.argv

    print("=" * 72)
    print("一键填充　附件1.声明（三家）+ 附件4~8")
    print("=" * 72)

    missing = check_inputs()
    if missing:
        print("\n前置文件不全，已停止：")
        for m in missing:
            print("  ✗", m)
        return 1

    print(f"\n中标方：{BIDDER}")
    print(f"结算清单：{SETTLE_XLSX.name}　← 价格与明细均以此为准")
    print(f"参考报价：{QUOTE_XLSX.name}")
    if dry:
        print("\n--dry 模式：前置检查通过，未写入任何文件。")
        return 0

    total_hit = 0

    print("\n[1/7] 1.声明（三家）")
    for dst, hits, miss in fillers.fill_declarations():
        print(f"  {dst.name}")
        show_hits("", hits, miss, "      ")
        total_hit += len(hits)

    steps = [
        ("2/7", "4、询价汇总表", fillers.fill_summary),
        ("3/7", "5、成交通知书", fillers.fill_notice),
        ("4/7", "6、合同", fillers.fill_contract),
        ("5/7", "7、签收单", fillers.fill_receipt),
        ("6/7", "8、验收书", fillers.fill_acceptance),
    ]
    for idx, label, fn in steps:
        print(f"\n[{idx}] {label}")
        dst, hits, miss = fn()
        print(f"  {dst.name}：命中 {len(hits)} 处")
        show_hits("", hits, miss)
        total_hit += len(hits)

    print("\n[7/7] 同步报价明细到 6/7/8 的表格")
    for path, ok, info in fillers.sync_quote_tables():
        flag = "✓" if ok else "✗"
        print(f"  {flag} {path.name}：{info}")

    print("\n" + "=" * 72)
    print("完成。输出文件：")
    print(f"  1.声明/已填/  （{len(list((DOC_DIR / '1.声明' / '已填').glob('*.docx')))} 份）")
    for name in ["4、询价汇总表_已填.docx", "5、成交通知书_已填.docx",
                 "6、合同_已填.docx", "7、签收单_已填.docx", "8、验收书_已填.docx"]:
        p = DOC_DIR / name
        mark = "✓" if p.exists() else "✗"
        print(f"  {mark} {name}")
    print(f"\n累计替换红字 {total_hit} 处；源文件均未改动。")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
