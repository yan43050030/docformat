# -*- coding: utf-8 -*-
"""套头对位校验：不用打印机，也能看出套打排版跟红头纸对不对得上。

做法
----
把生成的套打文件转成 PDF，再**叠**到套头纸（红头文件纸）的 PDF 上：
套头作底、内容作面。屏幕上看到的就是纸真正印出来的样子——黑字落在哪个
格子里、日期有没有压线，一目了然，省掉一次次试打浪费的纸。

为什么走"转 PDF 再叠"而不是自己按坐标画
--------------------------------------
自己画等于把 Word 的排版引擎重写一遍，本项目里已反复吃过亏（字宽、
网格吸附、字体度量都测不准）。转 PDF 用的是本机真正的
Word/WPS/LibreOffice，排出来什么样叠出来就什么样；而且预览与导出共用
同一条 build_alignment_pdf，不会各走各的。

与 header_overlay 的分工
------------------------
所有 PDF 操作都下沉到 scripts.header_overlay（PyMuPDF 封装），
本模块只做套打这一侧的编排：填模板 → 转 PDF → 叠加，外加"依赖缺失时
如实告知"。全项目因此只有一套 PDF 依赖，不重复引入第二个 PDF 栈。
"""
import logging
import os
import tempfile

logger = logging.getLogger('docformat.overlay')

LETTERHEAD_EXTS = ('.pdf',)


def _probe():
    """PyMuPDF 在不在。

    import 失败不一定是 ImportError：这类带二进制扩展的包在依赖坏掉时
    可能抛出别的异常，甚至是 BaseException（本项目遇到过 pyo3 panic），
    普通 except Exception 拦不住，所以这里兜到 BaseException。
    """
    try:
        import fitz            # noqa: F401  PyMuPDF
        return True, ''
    except BaseException as e:  # noqa: BLE001 - 见上
        if isinstance(e, (KeyboardInterrupt, SystemExit)):
            raise
        return False, '缺少 PyMuPDF（{}）'.format(e)


def can_merge():
    """能不能做叠加导出，返回 (可以吗, 原因)"""
    return _probe()


def can_render():
    """能不能把 PDF 渲染成图给预览用，返回 (可以吗, 原因)。

    与合并同源（都靠 PyMuPDF），分成两个函数是为了界面能分别说明——
    将来若换成"合并与渲染由不同组件承担"，调用方不必改。
    """
    return _probe()


def page_size_cm(pdf_path, page=0):
    """PDF 某页的纸张尺寸（cm），取不到返回 None"""
    ok, _why = _probe()
    if not ok:
        return None
    try:
        from scripts.header_overlay import page_size_cm as _size
        return _size(pdf_path, page)
    except Exception:
        return None


def page_count(pdf_path):
    ok, _why = _probe()
    if not ok:
        return 0
    try:
        from scripts.header_overlay import page_count as _count
        return _count(pdf_path)
    except Exception:
        return 0


def merge_overlay(content_pdf, letterhead_pdf, output_pdf, log=None):
    """把 content_pdf 叠到 letterhead_pdf 上另存，返回提示列表。

    尺寸不一致时如实提示：套头 PDF 若是扫描件或被裁过，叠出来只能凑合看，
    真正的解法是让两边都用未缩放的 A4。
    """
    ok, why = _probe()
    if not ok:
        raise RuntimeError('无法合并 PDF：{}'.format(why))
    from scripts.header_overlay import overlay_content_on_header

    notes = []
    bg = page_size_cm(letterhead_pdf)
    fg = page_size_cm(content_pdf)
    if bg and fg and (abs(bg[0] - fg[0]) > 0.2 or abs(bg[1] - fg[1]) > 0.2):
        notes.append(
            '套头纸 {:.1f}×{:.1f}cm 与生成内容 {:.1f}×{:.1f}cm 尺寸不一致，'
            '叠加时会被拉伸对齐；若套头 PDF 是扫描件或被裁过，'
            '建议改用未缩放的 A4 版套头，对位才准'.format(
                bg[0], bg[1], fg[0], fg[1]))
    n_bg, n_fg = page_count(letterhead_pdf), page_count(content_pdf)
    if n_bg and n_fg > n_bg:
        notes.append('内容有 {} 页，套头只有 {} 页；超出的页没有套头底图'
                     .format(n_fg, n_bg))

    overlay_content_on_header(letterhead_pdf, content_pdf, output_pdf)
    if log:
        log('info', '已把内容叠到套头纸上，共 {} 页'.format(n_fg or 1))
    return notes


def render_page_png(pdf_path, page=0):
    """把 PDF 某页渲染成临时 PNG，返回路径；调用方用完自行删除。"""
    ok, why = _probe()
    if not ok:
        raise RuntimeError('无法渲染 PDF：{}'.format(why))
    from scripts.header_overlay import render_page_to_png
    return render_page_to_png(pdf_path, page)


def build_alignment_pdf(template_path, values, letterhead_pdf, output_pdf,
                        title_shape='trapezoid_down', title_lines=None,
                        offsets=None, log=None):
    """一条龙：填模板 → 转 PDF → 叠到套头上 → 输出。

    返回 (输出路径, 提示列表)。预览和导出都走这里，
    保证"屏幕上看到的对位"就是"导出 PDF 里的对位"。
    """
    from scripts import overprint
    from scripts.exporter import export_pdf

    tmpdir = tempfile.mkdtemp(prefix='overlay_')
    docx_path = os.path.join(tmpdir, 'content.docx')
    pdf_path = os.path.join(tmpdir, 'content.pdf')

    _n, notes = overprint.fill_form(
        template_path, values, docx_path, log=log,
        title_shape=title_shape, title_lines=title_lines, offsets=offsets)

    ok, detail = export_pdf(docx_path, pdf_path, log=log)
    if not ok:
        raise RuntimeError(
            '无法把套打文件转成 PDF：{}\n'
            '本机需要装有 Word / WPS / LibreOffice 之一。'.format(detail))

    notes = list(notes) + list(merge_overlay(pdf_path, letterhead_pdf,
                                             output_pdf, log=log))
    return output_pdf, notes
