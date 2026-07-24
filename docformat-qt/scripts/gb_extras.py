# -*- coding: utf-8 -*-
"""GB/T 9704 版头/版记附加排版：红色分隔线、版记分隔线、图表题注居中。

均为默认关闭的可选项，由预设开关 header_elements / record_elements /
format_captions 控制，通过 engine 在段落类型识别后调用。
"""
import re

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt

_CAPTION_RE = re.compile(r'^(图|表)\s*([0-9]+|[一二三四五六七八九十]+)([\-—.－、]?\d+)?')


def _set_bottom_border(para, color='FF0000', size=12):
    """给段落加下边框（红色分隔线用 color=FF0000）。size 单位 1/8 pt。"""
    pPr = para._p.get_or_add_pPr()
    borders = pPr.find(qn('w:pBdr'))
    if borders is None:
        borders = OxmlElement('w:pBdr')
        # pBdr 需在 spacing 之前，简单起见插到 pPr 末尾也可被 Word 接受
        pPr.append(borders)
    for old in borders.findall(qn('w:bottom')):
        borders.remove(old)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    borders.append(bottom)


def _set_top_border(para, color='000000', size=6):
    pPr = para._p.get_or_add_pPr()
    borders = pPr.find(qn('w:pBdr'))
    if borders is None:
        borders = OxmlElement('w:pBdr')
        pPr.append(borders)
    for old in borders.findall(qn('w:top')):
        borders.remove(old)
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), str(size))
    top.set(qn('w:space'), '1')
    top.set(qn('w:color'), color)
    borders.append(top)


def apply_header_separator(typed_entries):
    """版头红色分隔线：在最后一个版头要素（发文字号/签发人）下方加红线。"""
    header_types = ('copynum', 'urgency', 'security', 'docnum', 'signatory')
    last = None
    for para, ptype in typed_entries:
        if ptype in header_types:
            last = para
        elif ptype in ('title', 'subtitle') and last is not None:
            break
    if last is not None:
        _set_bottom_border(last, color='FF0000', size=16)
        return True
    return False


def apply_record_separators(typed_entries):
    """版记分隔线：抄送上方加线，印发行下方加线（版记区上下各一条）。"""
    firsts = [p for p, t in typed_entries if t in ('cc', 'issuer')]
    if not firsts:
        return False
    # 第一个版记要素上方加线
    _set_top_border(firsts[0], color='000000', size=6)
    # 最后一个（印发）下方加线
    _set_bottom_border(firsts[-1], color='000000', size=6)
    return True


def apply_captions(blocks, preset):
    """图/表题注：紧邻图片段落或表格的"图N…""表N…"行居中、套题注字体。

    blocks: engine 的块序列（段落与表格按文档顺序）。仅在 format_captions 开启时调用。
    """
    from .table import _iter_block_items  # noqa
    from .font import set_font
    fmt = preset.get('caption', {})
    if not fmt:
        return 0
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    n = 0
    for idx, block in enumerate(blocks):
        if not isinstance(block, Paragraph):
            continue
        text = block.text.strip()
        if not _CAPTION_RE.match(text):
            continue
        # 相邻块是否为图片段落或表格
        prev_b = blocks[idx - 1] if idx > 0 else None
        next_b = blocks[idx + 1] if idx + 1 < len(blocks) else None

        def _is_media(b):
            if isinstance(b, Table):
                return True
            if isinstance(b, Paragraph):
                return b._p.find('.//' + qn('w:drawing')) is not None or \
                       b._p.find('.//' + qn('w:pict')) is not None
            return False

        if not (_is_media(prev_b) or _is_media(next_b)):
            continue
        block.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        block.paragraph_format.first_line_indent = Pt(0)
        block.paragraph_format.left_indent = Pt(0)
        for run in block.runs:
            if run.text.strip():
                set_font(run, fmt.get('font_cn', '宋体'),
                         fmt.get('font_en', 'Times New Roman'),
                         fmt.get('size', 12), bold=fmt.get('bold', False))
        n += 1
    return n
