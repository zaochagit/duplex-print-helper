# -*- coding: utf-8 -*-
"""
打包脚本：生成「绿色版」文件夹（onedir）
========================================
用法（在已装好依赖的 Python 环境下运行）:
    python build_exe.py

产物: dist\双面打印助手\  （一个文件夹，内含 双面打印助手.exe + 依赖库）
      -> 整个文件夹压缩后即为可分享的绿色版，解压到任意位置双击即用，无需安装。

⚠️ 体积提示:
    PyInstaller 会把当前 venv 里装的所有包都打进去。如果在装了 torch/cv2/scipy
    等重型包的环境里打包，体积会从几十 MB 暴涨到几百 MB。
    建议新建一个只装必要依赖的干净虚拟环境再打包:
        python -m venv dup_build
        dup_build\Scripts\pip install PyMuPDF pywin32 Pillow windnd pypdf pyinstaller
        dup_build\Scripts\python build_exe.py
"""
import subprocess
import sys

PY = sys.executable  # 当前 python 解释器
NAME = "双面打印助手"


def build_green():
    subprocess.run([PY, "-m", "PyInstaller", "--onedir", "--noconsole",
                   "--name", NAME, "duplex_gui.py"], check=True)


if __name__ == "__main__":
    build_green()
    print("完成。绿色版文件夹在 dist\\%s\\" % NAME)
