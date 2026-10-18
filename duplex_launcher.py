# -*- coding: utf-8 -*-
"""
双面打印助手 —— 本机快速启动器（极小 exe）
=========================================
本启动器只做一件事：把同包内的 duplex_gui.py / duplex_print.py 释放到临时目录，
然后用本机 venv 的 pythonw.exe 运行它。所有重依赖（PyMuPDF / pywin32 / PIL / tkinter）
都复用 D:\envs\pdfocr 里已装好的，因此本 exe 只有几 MB、启动快。
仅限本机使用（依赖该 venv 存在）；要拷到别的机器请用「双面打印助手.exe」自包含版。
"""
import os
import sys
import subprocess


def resource_path(rel):
    # PyInstaller onefile 会把 --add-data 的文件放在 sys._MEIPASS 临时目录
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def main():
    venv = r"D:\envs\pdfocr\Scripts\pythonw.exe"
    script = resource_path("duplex_gui.py")
    py = venv if os.path.exists(venv) else "pythonw"
    try:
        subprocess.run([py, script], check=True)
    except FileNotFoundError:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "无法启动",
            "未找到 Python 运行环境：\n%s\n\n请确认 D:\\envs\\pdfocr 已安装，\n"
            "或改用「双面打印助手.exe」自包含版。" % venv)


if __name__ == "__main__":
    main()
