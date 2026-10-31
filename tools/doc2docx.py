# -*- coding: utf-8 -*-
"""用 Word COM 把老格式 .doc / .wps 只读转换为 .docx，不修改源文件。

用法：
    python tools/doc2docx.py "6、合同.doc"
    python tools/doc2docx.py "6、合同.doc" "输出/合同.docx"

需要本机安装 Microsoft Word 和 pywin32。
"""
import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    sys.exit("缺少 pywin32，请先执行：pip install pywin32")


def convert(src, dst=None):
    src = Path(src).resolve()
    dst = Path(dst).resolve() if dst else src.with_suffix(".docx")
    if not src.is_file():
        raise FileNotFoundError(src)
    dst.parent.mkdir(parents=True, exist_ok=True)

    word = doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(str(dst), FileFormat=16)  # 16 = wdFormatXMLDocument
        print(f"OK: {dst}  ({dst.stat().st_size} bytes)")
        return dst
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        if word is not None:
            word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    convert(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
