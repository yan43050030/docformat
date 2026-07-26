# -*- coding: utf-8 -*-
"""套头 PDF 叠加：预览渲染 + PDF 导出合并。

纯函数，封装所有 PyMuPDF 操作。套头 PDF 作为背景，排版后的内容叠加在上方。
"""
import os
import tempfile

_HEADER_OVERLAY_DPI = 150

# 全局常量
_PT_PER_INCH = 72.0
_CM_PER_INCH = 2.54
PT_TO_PX = _HEADER_OVERLAY_DPI / _PT_PER_INCH
CM_TO_PX = _HEADER_OVERLAY_DPI / _CM_PER_INCH


def available():
    """PyMuPDF 在不在，返回 (可用吗, 原因)。

    套头叠加是附加功能，目标机器（尤其麒麟 ARM）未必装得上这个二进制包。
    调用方据此把入口置灰并说明原因，而不是等着甩一个
    "No module named 'fitz'" 给用户看。
    """
    try:
        import fitz            # noqa: F401
        return True, ''
    except BaseException as e:  # noqa: BLE001
        # 带二进制扩展的包依赖坏掉时未必抛 ImportError，
        # 本项目遇到过 pyo3 panic 那种 BaseException，普通 except 拦不住
        if isinstance(e, (KeyboardInterrupt, SystemExit)):
            raise
        return False, '缺少 PyMuPDF（{}）'.format(e)


def _fitz():
    """取 fitz 模块；没装就抛出一句人话，而不是 ImportError"""
    ok, why = available()
    if not ok:
        raise RuntimeError(
            '{}。套头叠加与对位校验需要它；其余功能不受影响。'.format(why))
    import fitz
    return fitz


def page_count(path):
    """返回 PDF 页数。"""
    fitz = _fitz()
    doc = fitz.open(path)
    n = len(doc)
    doc.close()
    return n


def page_size_cm(path, page_number=0):
    """返回 (width_cm, height_cm) 的页面尺寸。"""
    fitz = _fitz()
    doc = fitz.open(path)
    rect = doc[page_number].rect
    doc.close()
    w_cm = rect.width * _CM_PER_INCH / _PT_PER_INCH
    h_cm = rect.height * _CM_PER_INCH / _PT_PER_INCH
    return (w_cm, h_cm)


def render_page_to_png(path, page_number=0):
    """渲染 PDF 某一页为临时 PNG，返回文件路径。

    调用方负责用完后清理临时文件。
    """
    fitz = _fitz()
    doc = fitz.open(path)
    page = doc[page_number]
    mat = fitz.Matrix(_HEADER_OVERLAY_DPI / _PT_PER_INCH,
                      _HEADER_OVERLAY_DPI / _PT_PER_INCH)
    pix = page.get_pixmap(matrix=mat)
    fd, out = tempfile.mkstemp(suffix='.png', prefix='docformat_hdr_')
    os.close(fd)
    pix.save(out)
    doc.close()
    return out


def page_pixmap_size(path, page_number=0):
    """返回指定页渲染后的像素 (width_px, height_px)。"""
    fitz = _fitz()
    doc = fitz.open(path)
    rect = doc[page_number].rect
    doc.close()
    w_px = int(rect.width * _HEADER_OVERLAY_DPI / _PT_PER_INCH)
    h_px = int(rect.height * _HEADER_OVERLAY_DPI / _PT_PER_INCH)
    return (w_px, h_px)


def overlay_content_on_header(header_pdf_path, content_pdf_path, output_pdf_path):
    """将排版后的内容 PDF 叠加到套头 PDF 上。

    第 1 页：套头页作底，内容页叠在上方
    第 2+ 页：若套头有多页则继续用对应页作底，否则纯内容页
    """
    fitz = _fitz()

    header_doc = fitz.open(header_pdf_path)
    content_doc = fitz.open(content_pdf_path)
    out_doc = fitz.open()

    hdr_pages = len(header_doc)
    content_pages = len(content_doc)

    for i in range(content_pages):
        cpage = content_doc[i]
        rect = cpage.rect

        out_page = out_doc.new_page(width=rect.width, height=rect.height)

        # 套头底图
        hdr_idx = min(i, hdr_pages - 1) if hdr_pages > 0 else 0
        if i < hdr_pages:
            out_page.show_pdf_page(rect, header_doc, hdr_idx)

        # 内容叠层
        out_page.show_pdf_page(rect, content_doc, i)

    out_doc.save(output_pdf_path)
    out_doc.close()
    header_doc.close()
    content_doc.close()
