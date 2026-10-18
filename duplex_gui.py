# -*- coding: utf-8 -*-
"""
手动双面打印助手 —— 图形界面版（单文件 exe）
===========================================
基于 D:\envs\pdfocr 环境。复用 duplex_print.py 的拆页/打印核心逻辑。

功能:
  - 置顶悬浮小窗，路径可输入 / 浏览 / 直接拖入文件
  - 首次设置打印机后写入配置，之后每次直接打到该打印机
  - 点「开始打印」: 第一轮打完弹确认框（防误操作），确认后才打第二轮
  - 支持 PDF 与 Word(.doc/.docx/.docm/.rtf，自动经本机 Word 转 PDF)
  - 「校准测试」按钮，用 2 张纸验证本机进出纸机制

核心规则（喷墨纸盒取"朝下"那面打印）:
  ⚠ 绝不要整叠翻面！出纸字面朝上，翻面会让字面朝下、第二遍印在字上，整叠报废。
    第一轮打完，整叠原样放回纸盒即可。
"""
import os
import sys
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import duplex_print as dp
from windnd import hook_dropfiles

try:
    import win32print
    HAVE_WIN32 = True
except Exception:
    HAVE_WIN32 = False

CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "duplex_print")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")


def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def list_printers():
    if not HAVE_WIN32:
        return []
    names = []
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        for p in win32print.EnumPrinters(flags, None, 1):
            names.append(p[2])
    except Exception:
        pass
    return names


class App:
    def __init__(self, root):
        self.root = root
        self.doc = None
        self.first = self.second = None
        self.busy = False
        self.printer = load_config().get("printer")
        self._build()
        hook_dropfiles(root, func=self._on_drop)
        if not self.printer:
            # 首次使用，自动弹出打印机选择
            root.after(400, self.choose_printer)

    # ---------------------------------------------------------- UI
    def _build(self):
        root = self.root
        root.title("手动双面打印助手")
        root.attributes("-topmost", True)
        root.resizable(False, False)
        root.geometry("560x300")

        # 打印机行
        top = ttk.Frame(root)
        top.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(top, text="打印机:").pack(side="left")
        self.printer_var = tk.StringVar(value=self.printer or "(未设置，请点右侧「设置」)")
        ttk.Label(top, textvariable=self.printer_var, width=42, anchor="w").pack(side="left", padx=6)
        ttk.Button(top, text="设置", width=6, command=self.choose_printer).pack(side="right")

        # 文件选择
        frm = ttk.LabelFrame(root, text="PDF / Word 文件（可直接拖入窗口）")
        frm.pack(fill="x", padx=10, pady=4)
        self.path_var = tk.StringVar()
        self.ent = ttk.Entry(frm, textvariable=self.path_var)
        self.ent.pack(side="left", fill="x", expand=True, padx=6, pady=8)
        self.browse_btn = ttk.Button(frm, text="浏览", width=7, command=self.browse)
        self.browse_btn.pack(side="right", padx=6)

        # 提示条
        tip = ttk.Label(root,
                        text="⚠ 第一轮打完整叠【原样】放回（别翻面）。Word 文件会自动转 PDF。",
                        foreground="#b00", font=("", 9))
        tip.pack(fill="x", padx=10, pady=(2, 0))

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(root, textvariable=self.status_var, anchor="w", relief="sunken").pack(
            fill="x", padx=10, pady=(8, 0))

        # 按钮
        btns = ttk.Frame(root)
        btns.pack(fill="x", padx=10, pady=12)
        self.start_btn = ttk.Button(btns, text="开始打印", command=self.on_start)
        self.start_btn.pack(side="left", padx=6)
        self.cal_btn = ttk.Button(btns, text="校准测试", command=self.on_calibrate)
        self.cal_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="退出", command=root.destroy).pack(side="right", padx=6)

    # ---------------------------------------------------------- 交互
    def set_status(self, msg):
        self.status_var.set(msg)

    def safe(self, fn, *a):
        # 后台线程切回主线程更新 UI
        self.root.after(0, lambda: fn(*a))

    def set_busy(self, b):
        self.busy = b
        state = "disabled" if b else "normal"
        for w in (self.start_btn, self.browse_btn, self.cal_btn):
            w.config(state=state)

    def browse(self):
        if self.busy:
            return
        p = filedialog.askopenfilename(
            title="选择 PDF 或 Word 文件",
            filetypes=[("可打印文件", "*.pdf;*.doc;*.docx;*.docm;*.rtf"),
                       ("PDF", "*.pdf"),
                       ("Word", "*.doc;*.docx;*.docm;*.rtf"),
                       ("所有文件", "*.*")])
        if p:
            self.path_var.set(p)

    def _on_drop(self, files):
        if self.busy or not files:
            return
        f = files[0]
        path = f.decode("gbk") if isinstance(f, bytes) else str(f)
        self.path_var.set(path)

    def choose_printer(self):
        win = tk.Toplevel(self.root)
        win.title("选择打印机")
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.geometry("440x320")
        ttk.Label(win, text="选择默认打印机（之后会记住，不用再选）:").pack(
            anchor="w", padx=12, pady=10)
        printers = list_printers()
        lb = tk.Listbox(win, height=12, exportselection=False)
        for p in printers:
            lb.insert("end", p)
        if self.printer in printers:
            lb.selection_set(printers.index(self.printer))
        elif printers:
            lb.selection_set(0)
        lb.pack(fill="both", expand=True, padx=12)

        def ok():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("提示", "请选择一个打印机", parent=win)
                return
            name = lb.get(sel[0])
            self.printer = name
            self.printer_var.set(name)
            save_config({"printer": name})
            win.destroy()

        bf = ttk.Frame(win)
        bf.pack(fill="x", padx=12, pady=12)
        ttk.Button(bf, text="确定", command=ok).pack(side="left")
        ttk.Button(bf, text="取消", command=win.destroy).pack(side="right")
        win.grab_set()

    # ---------------------------------------------------------- 打印流程
    def _prepare_doc(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in dp.WORD_EXTS:
            self.set_status("正在把 Word 转成 PDF ...")
            pdf = dp.word_to_pdf(path)
        else:
            pdf = path
        return dp.load_pages(pdf)

    def on_start(self):
        if self.busy:
            return
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择或拖入 PDF / Word 文件")
            return
        if not os.path.isfile(path):
            messagebox.showerror("文件不存在", path)
            return
        if not self.printer:
            messagebox.showwarning("提示", "请先设置打印机")
            self.choose_printer()
            return
        try:
            doc = self._prepare_doc(path)
        except Exception as e:
            messagebox.showerror("准备失败", "文件处理出错: %s" % e)
            return
        n = doc.page_count
        if n < 2:
            messagebox.showwarning("提示", "只有 %d 页，无需双面打印" % n)
            doc.close()
            return
        self.doc = doc
        self.first, self.second = dp.plan_batches(n)
        self.set_status("计划 -> 第一轮: 第 %s 页 | 第二轮: 第 %s 页 | 共 %d 张"
                        % (dp.human(self.first), dp.human(self.second), (n + 1) // 2))
        self.set_busy(True)
        threading.Thread(target=self._batch1, daemon=True).start()

    def _batch1(self):
        try:
            self.safe(self.set_status, "正在打印第一轮 ...")
            jids = dp.print_batch(self.doc, self.first, self.printer, "duplex-batch1", dp.DEFAULT_DPI)
            dp.wait_print_done(self.printer, jids)
        except Exception as e:
            self.safe(messagebox.showerror, "打印出错", str(e))
            self.safe(self.set_busy, False)
            return
        self.safe(self._confirm_second)

    def _confirm_second(self):
        ok = messagebox.askokcancel(
            "第一轮完成",
            "第一轮已打印完，纸在出纸口。\n\n"
            "请把整叠纸【原样】放回纸盒——已打印面朝上，千万别翻面！\n"
            "（翻面会让第二遍直接印在字上，整叠报废）\n\n"
            "确认无误后点「确定」开始第二轮；不确定就点「取消」先检查。")
        if ok:
            threading.Thread(target=self._batch2, daemon=True).start()
        else:
            self.set_status("已暂停：第一轮纸在出纸口，检查后可重新点「开始打印」")
            self.set_busy(False)

    def _batch2(self):
        try:
            self.safe(self.set_status, "正在打印第二轮 ...")
            jids = dp.print_batch(self.doc, self.second, self.printer, "duplex-batch2", dp.DEFAULT_DPI)
            dp.wait_print_done(self.printer, jids)
        except Exception as e:
            self.safe(messagebox.showerror, "打印出错", str(e))
            self.safe(self.set_busy, False)
            return
        self.safe(self._finished)

    def _finished(self):
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        self.doc = None
        messagebox.showinfo("完成", "全部打印完成。\n\n成品第 1 页在最上面、字朝上，零整理。")
        self.set_status("就绪")
        self.set_busy(False)

    # ---------------------------------------------------------- 校准
    def on_calibrate(self):
        if self.busy:
            return
        if not self.printer:
            messagebox.showwarning("提示", "请先设置打印机")
            self.choose_printer()
            return
        self.set_busy(True)
        threading.Thread(target=self._calibrate, daemon=True).start()

    def _calibrate(self):
        doc = dp.make_calibrate_doc()
        first, second = dp.plan_batches(doc.page_count)
        try:
            self.safe(self.set_status, "校准: 打印第一轮 ...")
            jids = dp.print_batch(doc, first, self.printer, "cal-batch1", dp.DEFAULT_DPI)
            dp.wait_print_done(self.printer, jids)
            self.safe(messagebox.showinfo, "校准-第一轮完成",
                      "校准纸第一轮已打（2 张）。\n把纸叠原样放回纸盒，点「确定」开始第二轮。")
            self.safe(self.set_status, "校准: 打印第二轮 ...")
            jids = dp.print_batch(doc, second, self.printer, "cal-batch2", dp.DEFAULT_DPI)
            dp.wait_print_done(self.printer, jids)
        except Exception as e:
            self.safe(messagebox.showerror, "校准出错", str(e))
            self.safe(self.set_busy, False)
            return
        try:
            doc.close()
        except Exception:
            pass
        self.safe(messagebox.showinfo, "校准完成",
                  "检查成品:\n1) 有没有字叠字（有说明放回时被翻面了）\n"
                  "2) 第 1 页的背面是不是第 2 页\n都对即正常，可正常使用了。")
        self.safe(self.set_status, "就绪")
        self.safe(self.set_busy, False)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
