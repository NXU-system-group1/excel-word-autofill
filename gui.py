# -*- coding: utf-8 -*-
"""图形界面：上传结算清单，一键生成 1.声明（三家）+ 4~8 全部文档。

运行：python gui.py     （或双击 run_gui.bat）
"""
import contextlib
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import traceback
from pathlib import Path


class _NullIO:
    """pythonw 没有控制台，sys.stdout/stderr 是 None。

    各模块顶部都有 sys.stdout.reconfigure(encoding="utf-8")，直接调用会炸，
    所以先塞一个空壳顶上。
    """

    def write(self, s):
        pass

    def flush(self):
        pass

    def reconfigure(self, **kwargs):
        pass

    def isatty(self):
        return False


if sys.stdout is None:
    sys.stdout = _NullIO()
if sys.stderr is None:
    sys.stderr = _NullIO()


def _ensure_tcl():
    """Windows 下 Tcl 不认中文路径，把库复制到纯 ASCII 目录再指过去。

    只在 Tk 初始化失败时动手，正常情况下什么都不做。
    """
    try:
        import tkinter as _tk
        r = _tk.Tk()
        r.destroy()
        return
    except Exception:
        pass

    base = Path(sys.base_prefix)
    src = base / "Library" / "lib"
    if not src.exists():
        return

    dst = Path(tempfile.gettempdir()) / "wbtcl"
    try:
        dst.mkdir(parents=True, exist_ok=True)
        for name in ("tcl8.6", "tk8.6"):
            s, d = src / name, dst / name
            if s.exists() and not d.exists():
                shutil.copytree(s, d)
        if (dst / "tcl8.6").exists():
            os.environ["TCL_LIBRARY"] = str(dst / "tcl8.6")
            os.environ["TK_LIBRARY"] = str(dst / "tk8.6")
    except Exception:
        pass


_ensure_tcl()

import tkinter as tk  # noqa: E402
from tkinter import filedialog, messagebox, ttk  # noqa: E402

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT / "tools"))

import core      # noqa: E402
import fillers   # noqa: E402

FONT = "微软雅黑"
BG = "#F7F8FA"
CARD = "#FFFFFF"


class QueueWriter:
    """把 print 的输出塞进队列，由主线程取出来写文本框（Tk 不能跨线程操作）。"""

    def __init__(self, q):
        self.q = q

    def write(self, s):
        self.q.put(s)

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.busy = False

        root.title("询价文件自动填充")
        root.geometry("900x800")
        root.minsize(820, 700)
        root.configure(bg=BG)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=(FONT, 10), background=BG)
        style.configure("Hint.TLabel", font=(FONT, 9), foreground="#888")
        style.configure("Title.TLabel", font=(FONT, 15, "bold"), background=BG)
        style.configure("TButton", font=(FONT, 10))
        style.configure("Go.TButton", font=(FONT, 11, "bold"))
        style.configure("TEntry", font=(FONT, 10))

        pad = {"padx": 18, "pady": 4}

        head = ttk.Frame(root, style="TFrame")
        head.configure(style="TFrame")
        head.pack(fill="x", padx=18, pady=(16, 4))
        ttk.Label(head, text="询价文件自动填充", style="Title.TLabel").pack(anchor="w")
        ttk.Label(head, text="上传结算清单 → 自动填充 1.声明（三家）与 4~8 全部文档",
                  style="Hint.TLabel").pack(anchor="w", pady=(2, 0))

        card = tk.Frame(root, bg=CARD, highlightbackground="#E3E5E8", highlightthickness=1)
        card.pack(fill="x", padx=18, pady=(10, 8))

        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="x", padx=16, pady=14)
        inner.columnconfigure(1, weight=1)

        # 结算清单
        tk.Label(inner, text="结算清单", bg=CARD, font=(FONT, 10)).grid(
            row=0, column=0, sticky="w", pady=6)
        default = core.SETTLE_XLSX
        self.settle_var = tk.StringVar(value=str(default) if default.exists() else "")
        ttk.Entry(inner, textvariable=self.settle_var).grid(
            row=0, column=1, sticky="ew", padx=10)
        ttk.Button(inner, text="浏览…", width=8,
                   command=self.pick_settle).grid(row=0, column=2)
        tk.Label(inner, text="价格与明细的唯一来源", bg=CARD,
                 font=(FONT, 8), fg="#999").grid(row=1, column=1, sticky="w", padx=10)

        # 中标方
        tk.Label(inner, text="中标方", bg=CARD, font=(FONT, 10)).grid(
            row=2, column=0, sticky="w", pady=6)
        self.bidder_var = tk.StringVar(value=core.BIDDER)
        cb = ttk.Combobox(inner, textvariable=self.bidder_var, values=core.COMPANIES,
                          state="readonly", width=12)
        cb.grid(row=2, column=1, sticky="w", padx=10)
        tk.Label(inner, text="决定 4~8 号文件填哪一家的乙方信息", bg=CARD,
                 font=(FONT, 8), fg="#999").grid(row=3, column=1, sticky="w", padx=10)

        # 甲方信息
        tk.Label(inner, text="甲方名称", bg=CARD, font=(FONT, 10)).grid(
            row=4, column=0, sticky="w", pady=6)
        self.owner_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.owner_var).grid(
            row=4, column=1, sticky="ew", padx=10)
        tk.Label(inner, text="换项目时在这里改，会覆盖「红字信息/甲方.xlsx」", bg=CARD,
                 font=(FONT, 8), fg="#999").grid(row=5, column=1, sticky="w", padx=10)

        tk.Label(inner, text="项目名称", bg=CARD, font=(FONT, 10)).grid(
            row=6, column=0, sticky="w", pady=6)
        self.project_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.project_var).grid(
            row=6, column=1, sticky="ew", padx=10)

        tk.Label(inner, text="甲方地址", bg=CARD, font=(FONT, 10)).grid(
            row=7, column=0, sticky="w", pady=6)
        self.address_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.address_var).grid(
            row=7, column=1, sticky="ew", padx=10)

        # 输出目录
        tk.Label(inner, text="输出目录", bg=CARD, font=(FONT, 10)).grid(
            row=8, column=0, sticky="w", pady=6)
        self.out_var = tk.StringVar(value=str(core.DOC_DIR))
        ttk.Entry(inner, textvariable=self.out_var, state="readonly").grid(
            row=8, column=1, sticky="ew", padx=10)
        ttk.Button(inner, text="打开", width=8,
                   command=self.open_out).grid(row=8, column=2)

        # 按钮
        bar = tk.Frame(root, bg=BG)
        bar.pack(fill="x", padx=18, pady=(0, 6))
        self.btn_start = ttk.Button(bar, text="开始生成", style="Go.TButton",
                                    command=self.start)
        self.btn_start.pack(side="left")
        self.btn_clear = ttk.Button(bar, text="清空日志", command=self.clear_log)
        self.btn_clear.pack(side="left", padx=8)

        self.status = tk.StringVar(value="就绪")
        tk.Label(bar, textvariable=self.status, bg=BG, font=(FONT, 9),
                 fg="#666").pack(side="right")

        # 日志
        logwrap = tk.Frame(root, bg=BG)
        logwrap.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        self.text = tk.Text(logwrap, bg="#1E1E1E", fg="#D4D4D4",
                            font=("Consolas", 9), wrap="word",
                            relief="flat", padx=10, pady=8, state="disabled")
        sb = ttk.Scrollbar(logwrap, command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)

        self.text.tag_configure("ok", foreground="#6A9955")
        self.text.tag_configure("err", foreground="#F48771")
        self.text.tag_configure("hi", foreground="#DCDCAA")

        self.root.after(80, self.poll)
        self._prefill_owner()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 交互 ----------

    def _prefill_owner(self):
        def clean(v):
            v = (v or "").strip()
            return "" if v == "空" else v      # 「空」是占位值，输入框里留空更直观

        try:
            by_ph, _ = core.load_owner_info()
            self.owner_var.set(clean(by_ph.get("某某甲方")))
            self.project_var.set(clean(by_ph.get("某某甲方某某项目")))
            self.address_var.set(clean(by_ph.get("某某甲方地址")))
        except Exception:
            pass

    def _guess_from_title(self, path):
        """从结算清单标题猜甲方与项目名，填进输入框（可改）。"""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            title = str(wb.active.cell(1, 1).value or "").strip()
            wb.close()
        except Exception:
            return
        if not title:
            return
        m = re.search(r"(.{2,20}?(?:小学|中学|大学|幼儿园|学院|医院|办事处|中心))", title)
        name = m.group(1) if m else ""
        if name:
            self.owner_var.set(name)
        proj = re.sub(r"(报价清单|结算清单|采购询价清单|询价清单|清单)$", "", title).strip()
        if proj:
            self.project_var.set(proj)

    def pick_settle(self):
        init = str(Path(self.settle_var.get()).parent) if self.settle_var.get() else str(core.DOC_DIR)
        p = filedialog.askopenfilename(
            title="选择结算清单",
            initialdir=init,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if p:
            self.settle_var.set(p)
            self._guess_from_title(p)

    def open_out(self):
        d = self.out_var.get()
        if os.path.isdir(d):
            os.startfile(d)

    def clear_log(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def on_close(self):
        if self.busy and not messagebox.askyesno("正在生成", "还在处理中，确定退出吗？"):
            return
        self.root.destroy()

    # ---------- 执行 ----------

    def start(self):
        if self.busy:
            return
        path = Path(self.settle_var.get().strip())
        if not path.is_file():
            messagebox.showwarning("提示", "请先选择结算清单文件")
            return
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            messagebox.showwarning("格式不支持",
                                   "只支持 .xlsx / .xlsm。老格式 .xls 请先用 Excel 另存为 .xlsx。")
            return

        core.SETTLE_XLSX = path
        fillers.reset_settle()
        try:
            rows, amt, cn = core.load_settlement()
        except Exception as e:
            messagebox.showerror(
                "读取失败",
                f"无法读取这个文件：\n\n{path.name}\n\n"
                f"{type(e).__name__}: {e}\n\n"
                "需要包含「序号」表头行和「合计」行的明细表。")
            return

        bidder = self.bidder_var.get()
        fillers.set_bidder(bidder)
        core.BIDDER = bidder

        core.reset_owner_override()
        core.set_owner_override(
            name=self.owner_var.get().strip() or None,
            project=self.project_var.get().strip() or None,
            address=self.address_var.get().strip() or None,
        )

        self.clear_log()
        self.busy = True
        self.btn_start.configure(state="disabled")
        self.status.set("正在生成…")

        threading.Thread(target=self._work, args=(path, rows, amt, cn, bidder),
                         daemon=True).start()

    def _work(self, path, rows, amt, cn, bidder):
        w = QueueWriter(self.q)
        ok = True
        try:
            with contextlib.redirect_stdout(w):
                print("=" * 62)
                print("结算清单：" + path.name)
                print(f"明细 {len(rows)} 行　合计 {amt:,.2f}　大写 {cn}")
                print("中标方：" + bidder)
                print("=" * 62)

                print("\n[1/7] 1.声明（三家）")
                for dst, hits, miss in fillers.fill_declarations():
                    print(f"  {dst.name}：命中 {len(hits)} 处")
                    for old, new in hits:
                        print(f"      {old} → {new}")

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

                print("\n[7/7] 同步结算明细到 6/7/8 的表格")
                for p, good, info in fillers.sync_quote_tables():
                    print(f"  {'✓' if good else '✗'} {p.name}：{info}")

                print("\n" + "=" * 62)
                print("完成，源文件均未改动。")
                print("=" * 62)
        except PermissionError as e:
            ok = False
            name = str(e).split(":")[-1].strip().strip("'\"")
            w.write("\n[出错] 文件被占用，无法保存：\n")
            w.write(f"    {name}\n")
            w.write("    请关闭 WPS / Word 里打开的这个文件后重试。\n")
        except Exception:
            ok = False
            w.write("\n[出错]\n" + traceback.format_exc())
        finally:
            self.root.after(0, lambda: self._done(ok))

    def _done(self, ok):
        self.busy = False
        self.btn_start.configure(state="normal")
        self.status.set("完成" if ok else "出错")
        if ok:
            messagebox.showinfo("完成", "全部文件已生成。\n\n点「打开」可查看输出目录。")
        else:
            messagebox.showerror("出错", "生成过程中出现错误，详见日志。")

    # ---------- 日志轮询 ----------

    def poll(self):
        chunks = []
        while True:
            try:
                chunks.append(self.q.get_nowait())
            except queue.Empty:
                break
        if chunks:
            self.text.configure(state="normal")
            for s in chunks:
                tag = None
                if s.startswith("✓") or "完成" in s[:6]:
                    tag = "ok"
                elif "出错" in s or s.startswith("✗"):
                    tag = "err"
                self.text.insert("end", s, tag)
            self.text.see("end")
            self.text.configure(state="disabled")
        self.root.after(80, self.poll)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
