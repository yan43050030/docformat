# -*- coding: utf-8 -*-
"""把文件送到打印机。

套打是"打一张、量一下、微调、再打一张"的活，来回都要经过 Word。这里
让软件自己能打：docx 先转 PDF（转换链本就有），再按页渲染送打印机。

为什么不直接调系统的"打印"动作
------------------------------
套打对**缩放**零容忍：纸上预印的栏位在固定位置，内容缩了 96% 就全废。
系统默认的打印动作会走各家阅读器自己的"适合纸张"缩放，用户还未必看得见
那个开关。这里自己渲染，把缩放钉死在 100%，并强制 A4 纵向。
"""
import logging
import os
import tempfile

logger = logging.getLogger('docformat.printing')

# A4 实际尺寸（mm），送打印机时按它定版，不随系统默认纸张走
A4_MM = (210.0, 297.0)


def available():
    """(能不能打, 说明)"""
    try:
        from PyQt5 import QtPrintSupport      # noqa: F401
    except ImportError:
        return False, '本机的 PyQt5 没装打印模块（QtPrintSupport）'
    try:
        import fitz                            # noqa: F401
    except ImportError:
        return False, '缺少 PyMuPDF，无法把文件渲染成打印页（pip install PyMuPDF）'
    return True, ''


def to_pdf(path, log=None):
    """任何输入 → PDF 路径（已经是 PDF 就原样返回）。第二个值是临时目录，用完删。"""
    if path.lower().endswith('.pdf'):
        return path, None
    from .exporter import export_pdf
    tmp = tempfile.mkdtemp(prefix='docformat_print_')
    out = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + '.pdf')
    ok, info = export_pdf(path, out, log)
    if not ok:
        raise RuntimeError('转 PDF 失败：{}'.format(info))
    return out, tmp


def print_pdf(pdf_path, printer):
    """按 100% 实际大小把 PDF 打到 printer 上，返回页数。

    每一页都按"纸多大就画多大"铺满，不做任何适应性缩放——套打差一毫米
    就废一张纸，宁可让用户自己去调打印机的进纸，也不能替他偷偷缩。
    """
    import fitz
    from PyQt5.QtCore import QRectF
    from PyQt5.QtGui import QImage, QPainter

    doc = fitz.open(pdf_path)
    painter = QPainter()
    if not painter.begin(printer):
        raise RuntimeError('打印机没能打开，可能正被占用')
    try:
        # 渲染分辨率跟着打印机走，别按屏幕的 96dpi 出图
        dpi = float(printer.resolution() or 300)
        for i, page in enumerate(doc):
            if i:
                printer.newPage()
            pm = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0),
                                 alpha=False)
            img = QImage(pm.samples, pm.width, pm.height, pm.stride,
                         QImage.Format_RGB888)
            # 画到"整张纸"上，不是可打印区：可打印区因打印机而异，
            # 按它铺会把内容整体缩进去一圈，套打就对不上了
            from PyQt5.QtPrintSupport import QPrinter
            rect = printer.paperRect(QPrinter.DevicePixel)
            painter.drawImage(QRectF(rect), img)
        return doc.page_count
    finally:
        painter.end()
        doc.close()


def make_printer():
    """一台按 A4 纵向、不缩放配好的打印机对象"""
    from PyQt5.QtPrintSupport import QPrinter
    p = QPrinter(QPrinter.HighResolution)
    p.setPageSize(QPrinter.A4)
    p.setOrientation(QPrinter.Portrait)
    p.setFullPage(True)          # 按整张纸定版，不留打印机自己的边
    try:
        p.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
    except Exception:
        pass
    return p


def print_file(path, printer=None, log=None):
    """打一个 docx/pdf。printer 为空就现配一台默认的。返回页数。"""
    ok, why = available()
    if not ok:
        raise RuntimeError(why)
    pdf, tmp = to_pdf(path, log)
    try:
        return print_pdf(pdf, printer or make_printer())
    finally:
        if tmp:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
