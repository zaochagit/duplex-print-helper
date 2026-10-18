# -*- coding: utf-8 -*-
"""
打包脚本：生成两个 exe
用法（在 pdfocr venv 下运行，且已 pip install pyinstaller windnd）:
    python build_exe.py          # 两个都打
    python build_exe.py self     # 只打自包含版（可拷到别的机器，不依赖本机环境）
    python build_exe.py local    # 只打本机 launcher 版（小、快，依赖 D:\envs\pdfocr）
"""
import os
import subprocess
import sys

PY = sys.executable  # 应在 pdfocr venv 下运行
SEP = os.pathsep     # Windows 用 ';'，Linux/macOS 用 ':'
SELF_NAME = "双面打印助手"
LOCAL_NAME = "双面打印助手-本机"


def build_self():
    subprocess.run([PY, "-m", "PyInstaller", "--onefile", "--noconsole",
                   "--name", SELF_NAME, "duplex_gui.py"], check=True)


def build_local():
    subprocess.run([PY, "-m", "PyInstaller", "--onefile", "--noconsole",
                   "--name", LOCAL_NAME,
                   "--add-data", "duplex_gui.py" + SEP + ".",
                   "--add-data", "duplex_print.py" + SEP + ".",
                   "duplex_launcher.py"], check=True)


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    if which in ("self", "all"):
        build_self()
    if which in ("local", "all"):
        build_local()
    print("完成。exe 在 dist\\ 目录。")
