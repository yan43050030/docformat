# -*- coding: utf-8 -*-
"""
段落格式化 — 从 formatter.py 拆分

包含：format_paragraph、_set_paragraph_spacing_points、
      结构空行处理、deep_clean_document
"""

import re
import logging
from copy import deepcopy
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .font import set_font, _force_normal_style, _add_ppr_change

logger = logging.getLogger('docformat.paragraph')


def _set_paragraph_spacing_points(para, before_pt=0, after_pt=0):
    """Set paragraph spacing in points and clear line-based spacing leftovers."""
    pf = para.paragraph_format
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)

    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)

    for attr in ('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing'):
        spacing.attrib.pop(qn(f'w:{attr}'), None)

    spacing.set(qn('w:before'), str(int(round(before_pt * 20))))
    spacing.set(qn('w:after'), str(int(round(after_pt * 20))))


def _compact_empty_paragraph(para):
    """Clear spacing on empty paragraphs."""
    _set_paragraph_spacing_points(para, 0, 0)
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(1)


def _mark_structural_blank(para):
    """在段落 pPr 上写自定义标记，标识为结构性空行"""
    pPr = para._p.get_or_add_pPr()
    pPr.set('docfmt-structural-blank', '1')


def _is_structural_blank(para):
    """检查段落是否被标记为结构性空行"""
    pPr = para._p.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    if pPr is None:
        return False
    return pPr.get('docfmt-structural-blank') == '1'


def _format_structural_blank_paragraph(para, line_spacing_pt=28):
    """Format the intentional blank line used between document sections."""
    if not para.runs:
        para.add_run(' ')
    _set_paragraph_spacing_points(para, 0, 0)
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line_spacing_pt)
    _mark_structural_blank(para)


def _format_empty_paragraphs(doc, structural_blank_ids, line_spacing_pt=28):
    """格式化文档中所有空段落"""
    for para in doc.paragraphs:
        if para.text.strip():
            continue
        if _is_structural_blank(para):
            _format_structural_blank_paragraph(para, line_spacing_pt)
        else:
            _compact_empty_paragraph(para)


def deep_clean_document(doc):
    """深度清洗文档：移除所有段落级用户格式属性"""

    def _clean_paragraph(para):
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = None
        pf.right_indent = None
        pf.first_line_indent = None
        pf.line_spacing = None
        pf.line_spacing_rule = None
        _force_normal_style(para)
        for run in para.runs:
            run.font.color.rgb = None
            run.font.highlight_color = None
            run.font.size = None
            run.font.bold = None
            run.font.italic = None
            run.font.underline = None
            run.font.strike = None

    for para in doc.paragraphs:
        _clean_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _clean_paragraph(para)


def format_paragraph(para, fmt, para_type, line_spacing_pt=28, first_line_bold=False, revision_mode=False, bold_serial=True):
    """格式化段落

    fmt 支持的字段:
        font_cn, font_en, size, bold, align, indent,
        line_spacing  - 行距(磅), None表示使用1.5倍行距
        space_before  - 段前间距(磅), 默认0
        space_after   - 段后间距(磅), 默认0
    """
    _force_normal_style(para)

    if revision_mode:
        orig_ppr = deepcopy(para._p.pPr)
        orig_ppr_xml = para._p.xml

    pf = para.paragraph_format

    # 对齐方式
    align_map = {
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    pf.alignment = align_map.get(fmt.get('align', 'justify'), WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 段落左缩进清零
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)

    # attachment 类型走悬挂缩进
    _attachment_indent_done = False
    if para_type == 'attachment':
        font_size_pt = fmt.get('size', 16) or 16
        pf.left_indent = Pt(font_size_pt * 5)
        if '附件' in para.text:
            pf.first_line_indent = Pt(-font_size_pt * 3)
        else:
            pf.first_line_indent = Pt(0)
        try:
            pPr = para._p.get_or_add_pPr()
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                ind.attrib.pop(qn('w:firstLineChars'), None)
        except Exception:
            pass
        _attachment_indent_done = True

    # 首行缩进
    if not _attachment_indent_done:
        indent = fmt.get('indent', 0)
        if indent > 0:
            pf.first_line_indent = Pt(indent)
            size = fmt.get('size', 16) or 16
            try:
                chars_100 = int(round(indent / size * 100))
                pPr = para._p.get_or_add_pPr()
                ind = pPr.find(qn('w:ind'))
                if ind is None:
                    ind = OxmlElement('w:ind')
                    pPr.append(ind)
                ind.set(qn('w:firstLineChars'), str(chars_100))
            except Exception:
                pass
        else:
            pf.first_line_indent = Pt(0)
            try:
                pPr = para._p.get_or_add_pPr()
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    ind.attrib.pop(qn('w:firstLineChars'), None)
            except Exception:
                pass

    # 行距
    ls = fmt.get('line_spacing', line_spacing_pt)
    if ls:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(ls)
    else:
        pf.line_spacing = 1.5

    # 段前段后
    _set_paragraph_spacing_points(
        para,
        fmt.get('space_before', 0),
        fmt.get('space_after', 0)
    )

    # 字体 - 支持首句加粗
    if first_line_bold and para_type == 'body':
        full_text = para.text
        first_sentence_end = full_text.find('。')
        if first_sentence_end != -1:
            split_idx = first_sentence_end + 1
            first_part = full_text[:split_idx]
            rest_part = full_text[split_idx:]

            for run in list(para.runs):
                para._p.remove(run._r)

            run1 = para.add_run(first_part)
            set_font(run1, fmt['font_cn'], fmt['font_en'], fmt['size'], bold=True,
                      italic=fmt.get('italic', False), color=fmt.get('color'),
                      revision_mode=revision_mode)

            if rest_part:
                run2 = para.add_run(rest_part)
                set_font(run2, fmt['font_cn'], fmt['font_en'], fmt['size'], fmt.get('bold', False),
                          italic=fmt.get('italic', False), color=fmt.get('color'),
                          revision_mode=revision_mode)
        else:
            for run in para.runs:
                set_font(run, fmt['font_cn'], fmt['font_en'], fmt['size'], fmt.get('bold', False),
                          italic=fmt.get('italic', False), color=fmt.get('color'),
                          revision_mode=revision_mode)
    else:
        # 正文里的序列词加粗前缀
        if bold_serial and para_type == 'body':
            _SERIAL_PATTERNS = [
                r'^([一二三四五六七八九十]{1,3}是)([：:、]?)',
                r'^([一二三四五六七八九十]{1,3}要)([：:、]?)',
                r'^(第[一二三四五六七八九十百\d]+[点条项步])([：:、，,]?)',
                r'^([一二三四五六七八九十]{1,3}方面)([：:、]?)',
            ]
            m = None
            for _pat in _SERIAL_PATTERNS:
                m = re.match(_pat, para.text)
                if m:
                    break
            if m:
                lead = m.group(1) + (m.group(2) or '')
                rest = para.text[len(lead):]
                for run in list(para.runs):
                    para._p.remove(run._r)
                run1 = para.add_run(lead)
                set_font(run1, fmt['font_cn'], fmt['font_en'], fmt['size'], bold=True,
                      italic=fmt.get('italic', False), color=fmt.get('color'),
                      revision_mode=revision_mode)
                if rest:
                    run2 = para.add_run(rest)
                    set_font(run2, fmt['font_cn'], fmt['font_en'], fmt['size'], fmt.get('bold', False),
                          italic=fmt.get('italic', False), color=fmt.get('color'),
                          revision_mode=revision_mode)
                return

        for run in para.runs:
            set_font(run, fmt['font_cn'], fmt['font_en'], fmt['size'], fmt.get('bold', False),
                      italic=fmt.get('italic', False), color=fmt.get('color'),
                      revision_mode=revision_mode)

    # 修订模式
    if revision_mode and para._p.xml != orig_ppr_xml:
        _add_ppr_change(para, orig_ppr)
