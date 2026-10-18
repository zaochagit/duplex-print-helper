# -*- coding: utf-8 -*-
"""
Brother DCP-T420W 手动双面打印助手
==================================
原理: 第一轮打背面页（偶数页；页数为奇数时末尾补一张空白页），
     整叠原样放回，第二轮打正面页（降序）——最后一张正好是第 1 页，
     成品第 1 页在最上面，零整理。

用法:
    双击/拖 PDF 或 Word 到 duplex.bat 上（最简单）
    或命令行: D:\\envs\\pdfocr\\Scripts\\python.exe duplex_print.py 文件.pdf
    支持: .pdf .doc .docx .docm .rtf（Word 文件自动经本机 Word 转 PDF，版式与 Word 打印一致）

第一次使用先跑校准（会用 2 张纸）:
    python duplex_print.py --calibrate
    打完检查两件事:
      1) 有没有哪一面被重复打印（字叠字）——有说明纸叠放回时被翻过面
      2) 第 1 页的背面是不是第 2 页

常用参数:
    文件.pdf                正常双面打印
    --dpi 300              渲染精度（默认 300，够用且快）
    --calibrate            打印 4 页校准测试
    --dry-run              只显示计划不实际打印
    --printer "名字"       指定打印机（默认 Brother DCP-T420W）

放回规则（关键，喷墨纸盒取"朝下"那面打印）:
    ⚠ 绝不要整叠翻面！出纸时字面朝上，翻面会让字面朝下，
      第二遍直接印在字上，整叠报废。整叠原样放回纸盒即可。

奇数页 PDF:
    不用手动加纸。第一轮末尾会多打一张空白页，让末页的纸先走一遍，
    第二轮就是完整的一轮，末张正好是第 1 页。

整理成品:
    第 1 页在最上面、字朝上，直接可用，零整理。

卡纸/错位排查:
    - 字叠字（打在已打印面） -> 放回时把纸叠翻过来了，重打
    - 背面文字上下颠倒       -> 纸叠放回时前后方向反了（应保持原方向）
    - 页码配对错乱           -> 多半是纸叠顺序被打乱，原样放回重打
"""
import os
import sys
import time
import argparse

DEFAULT_PRINTER = "Brother DCP-T420W Printer"
DEFAULT_DPI = 300


# ---------------------------------------------------------------- PDF 处理

def load_pages(path):
    """返回 fitz 文档对象"""
    import fitz
    doc = fitz.open(path)
    if doc.needs_pass:
        raise RuntimeError("PDF 已加密，请先解密")
    return doc


def make_calibrate_doc():
    """生成 4 页校准 PDF，大页码居中"""
    import fitz
    doc = fitz.open()
    for i in range(1, 5):
        page = doc.new_page()  # 默认 A4 (595x842pt)
        page.insert_text(fitz.Point(230, 420), str(i), fontsize=160)
        page.insert_text(fitz.Point(170, 480), "duplex calibrate  %d / 4" % i, fontsize=12)
    return doc


def plan_batches(n_pages):
    """生成两轮打印计划，返回 (batch1, batch2)；元素为 0-based 页号，None=空白页。

    打印机实测机制（用户确认）：后打的纸压在最上面 ——
        顺序打印 -> 最后一页在最上；逆序打印 -> 第一页在最上。
    因此第二轮用降序，让第 1 页最后打、压顶、成品零整理。

    放回规则：整叠【原样】放回（不翻面）-> 取纸顺序被倒转
    （第一轮最后打、压在最上的那张最先被取走）。

    偶数页：batch1=偶数页顺序，batch2=奇数页降序。
    奇数页：末页是单独一张（仅正面）。第一轮末尾补一张空白页让该纸先走一遍，
            batch2=[末页] + 奇数页降序，两轮张数相等，成品第 1 页仍压顶。
    """
    pages = list(range(n_pages))
    odds = pages[0::2]    # 0,2,4...（第 1,3,5...页，奇数页数时末页已含在内）
    evens = pages[1::2]   # 1,3,5...（第 2,4,6...页）
    # 偶数: batch1=偶数页顺序, batch2=奇数页降序
    # 奇数: 第一轮末尾补一张空白页(让末页那张纸先走一遍)，两轮张数仍相等
    batch1 = evens + ([None] if n_pages % 2 == 1 else [])
    batch2 = list(reversed(odds))
    return batch1, batch2


def human(page_nums):
    out = []
    for p in page_nums:
        out.append("白纸" if p is None else str(p + 1))
    return ",".join(out)


# ---------------------------------------------------------------- Word -> PDF

WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}


def word_to_pdf(path):
    """用本机 MS Word 把文档转成 PDF，返回生成的 PDF 路径
    DispatchEx 开独立进程，不影响用户正开着的 Word 窗口
    """
    import tempfile
    import win32com.client
    out = os.path.join(
        tempfile.gettempdir(),
        "duplex_" + os.path.splitext(os.path.basename(path))[0] + ".pdf",
    )
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        d.ExportAsFixedFormat(os.path.abspath(out), 17)  # 17 = wdExportFormatPDF
        d.Close(False)
    finally:
        word.Quit()
    return out


# ---------------------------------------------------------------- 打印

def render_page(doc, idx, dpi):
    """渲染一页为 PIL 图像；None 渲染为白纸"""
    from PIL import Image
    if idx is None:
        r = doc[0].rect
        return Image.new("RGB", (int(r.width * dpi / 72), int(r.height * dpi / 72)), "white")
    page = doc[idx]
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if page.rect.width > page.rect.height:
        # 横版页旋转 90 度适配纵向纸道（正反面同方向旋转，长边翻不乱）
        img = img.transpose(Image.Transpose.ROTATE_270)
    return img


def list_spool_jobs(printer_name):
    try:
        import win32print
        h = win32print.OpenPrinter(printer_name)
        try:
            jobs = win32print.EnumJobs(h, 0, 200, 1)
            return {j["JobId"]: j.get("Status", 0) for j in jobs}
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return {}


def wait_print_done(printer_name, new_job_ids, timeout=300):
    """等打印作业离开队列或标记完成，避免用户过早抽纸"""
    printed_flag = 0x80 | 0x1000  # JOB_STATUS_PRINTED | JOB_STATUS_COMPLETE
    t0 = time.time()
    while time.time() - t0 < timeout:
        cur = list_spool_jobs(printer_name)
        if not new_job_ids:
            return True
        if all(jid not in cur or (cur[jid] & printed_flag) for jid in new_job_ids):
            return True
        time.sleep(2)
    return False


def set_printer_direct(printer_name, enable):
    """把打印机切到『直连/前台』打印(enable=True)或恢复(enable=False)。
    返回原来的 Attributes；失败返回 None（调用方据此走后台兜底）。
    直连模式下 GDI 的 EndDoc 会阻塞到数据发到打印机，天然『前台执行』，
    对大扫描件 PDF 不进 spool，零等待；同时也就不需要手动勾『立即开始打印』。"""
    import win32print
    try:
        DIRECT = getattr(win32print, "PRINTER_ATTRIBUTE_DIRECT", 0x20)
        h = win32print.OpenPrinter(printer_name)
        try:
            pi2 = win32print.GetPrinter(h, 2)
            old = pi2["Attributes"]
            pi2["Attributes"] = (old | DIRECT) if enable else (old & ~DIRECT)
            win32print.SetPrinter(h, 2, pi2, 0)
            return old
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return None


def print_batch(doc, page_nums, printer_name, job_name, dpi):
    """把一批页面按给定顺序送到打印机，返回新产生的作业 id 集合"""
    import win32con
    import win32ui
    from PIL import ImageWin

    before = set(list_spool_jobs(printer_name))
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    area_w = hdc.GetDeviceCaps(win32con.HORZRES)    # 可打印区宽（设备像素）
    area_h = hdc.GetDeviceCaps(win32con.VERTRES)    # 可打印区高

    hdc.StartDoc(job_name)
    for idx in page_nums:
        img = render_page(doc, idx, dpi)
        # 等比缩放居中到可打印区
        scale = min(area_w / img.width, area_h / img.height)
        w2, h2 = int(img.width * scale), int(img.height * scale)
        x, y = (area_w - w2) // 2, (area_h - h2) // 2
        hdc.StartPage()
        ImageWin.Dib(img).draw(hdc.GetHandleOutput(), (x, y, x + w2, y + h2))
        hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()

    after = set(list_spool_jobs(printer_name))
    return after - before


# ---------------------------------------------------------------- 主流程

def run(doc, args):
    n = doc.page_count
    if n < 2:
        print("只有 %d 页，不需要双面打印" % n)
        return
    first, second = plan_batches(n)

    print("=" * 52)
    print("打印机 : %s" % args.printer)
    print("总页数 : %d" % n)
    print("纸张数 : %d" % ((n + 1) // 2))
    print("第一轮 : 第 %s 页" % human(first))
    print("第二轮 : 第 %s 页" % human(second))
    if n % 2 == 1:
        print("提示   : 奇数页数，第一轮末尾会多打一张空白页（末页的纸先走一遍），无需手动加纸")
    print("=" * 52)
    if args.dry_run:
        print("[dry-run] 未实际打印")
        return

    # 接管后台设置：强制直连(前台)打印，大扫描件不进 spool、零等待；打完恢复
    old_attr = set_printer_direct(args.printer, True)
    if old_attr is None:
        print("提示: 无法将打印机切换为直连(前台)打印，改用系统默认后台打印。")
        print("      若大文件卡顿，请在 打印机属性→高级 勾选『立即开始打印』。")
    try:
        job_ids = set()

        input("\n放好纸，按回车开始打印第一轮 -> ")
        job_ids |= print_batch(doc, first, args.printer, "duplex-batch1", args.dpi)
        if wait_print_done(args.printer, job_ids):
            print("第一轮打印完成。")
        else:
            print("警告: 打印队列超时未清空，请确认出纸完毕再继续！")

        input("\n整叠纸原样放回纸盒，什么都不动。千万别翻面（翻了就印在字上）。按回车开始第二轮 -> ")
        job_ids |= print_batch(doc, second, args.printer, "duplex-batch2", args.dpi)
        if wait_print_done(args.printer, job_ids):
            print("第二轮打印完成，成品在出纸口。")
        else:
            print("打印队列仍在处理，请等最后一张纸出来后再取。")
        print("提示: 第1页在最上面、字朝上，直接可用，零整理。")
        input("按回车退出。")
    finally:
        if old_attr is not None:
            set_printer_direct(args.printer, False)  # 恢复原始设置


def main():
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("path", nargs="?", help="PDF 文件路径")
    ap.add_argument("--printer", default=DEFAULT_PRINTER)
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    ap.add_argument("--calibrate", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if args.calibrate:
        doc = make_calibrate_doc()
    else:
        path = args.path
        if not path:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            path = filedialog.askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
            if not path:
                print("未选择文件，退出。")
                return 1
        if not os.path.isfile(path):
            print("文件不存在: %s" % path)
            return 1
        ext = os.path.splitext(path)[1].lower()
        if ext in WORD_EXTS:
            print("Word 文件，正在转换 PDF ...")
            try:
                path = word_to_pdf(path)
                print("转换完成: %s" % path)
            except Exception as e:
                print("Word 转 PDF 失败: %s" % e)
                input("按回车退出。")
                return 1
        try:
            doc = load_pages(path)
        except Exception as e:
            print("打开 PDF 失败: %s" % e)
            return 1

    try:
        run(doc, args)
    except KeyboardInterrupt:
        print("\n已取消。")
        return 1
    except Exception as e:
        print("出错: %s" % e)
        input("按回车退出。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
