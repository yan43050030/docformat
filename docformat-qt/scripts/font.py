# -*- coding: utf-8 -*-
"""
字体工具 — 从 formatter.py 拆分

包含：set_font、_force_normal_style、字体相关辅助函数
"""

import logging
from copy import deepcopy
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger('docformat.font')


def set_font(run, font_cn, font_en, size, bold=False, italic=False,
              color=None, revision_mode=False):
    """设置字体，同时清除原有格式

    Args:
        run: python-docx Run 对象
        font_cn: 中文字体名称
        font_en: 英文字体名称
        size: 字号（磅）
        bold: 是否加粗
        italic: 是否斜体（默认 False，公文标准不使用斜体）
        color: RGB 颜色元组 (R,G,B) 或 RGBColor，None 表示黑色
        revision_mode: 是否启用修订模式
    """
    if revision_mode:
        orig_rpr = deepcopy(run._r.rPr)
        orig_xml = run._r.xml

    run.font.name = font_en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = False
    if color is not None:
        if isinstance(color, (tuple, list)) and len(color) == 3:
            run.font.color.rgb = RGBColor(*color)
        else:
            run.font.color.rgb = color
    else:
        run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.strike = False
    run.font.double_strike = False
    run.font.subscript = False
    run.font.superscript = False

    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_en)
    rFonts.set(qn('w:hAnsi'), font_en)
    rFonts.set(qn('w:cs'), font_en)

    if revision_mode and run._r.xml != orig_xml:
        _add_rpr_change(run, orig_rpr)


def _force_normal_style(para):
    """把段落 style 重置为 Normal"""
    try:
        pPr = para._p.get_or_add_pPr()
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            pStyle = OxmlElement('w:pStyle')
            pPr.insert(0, pStyle)
        pStyle.set(qn('w:val'), 'Normal')
    except Exception:
        pass


# ===== 修订标记辅助 =====
_revision_counter = [0]


def _next_rev_id():
    _revision_counter[0] += 1
    return _revision_counter[0]


def _rev_date():
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def _add_ppr_change(para, orig_ppr):
    """将原始段落格式嵌入 <w:pPrChange>"""
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:pPrChange')):
        pPr.remove(old)

    change = OxmlElement('w:pPrChange')
    change.set(qn('w:id'), str(_next_rev_id()))
    change.set(qn('w:author'), '公文格式工具')
    change.set(qn('w:date'), _rev_date())

    if orig_ppr is not None:
        snapshot = deepcopy(orig_ppr)
        for old in snapshot.findall(qn('w:pPrChange')):
            snapshot.remove(old)
        change.append(snapshot)
    else:
        change.append(OxmlElement('w:pPr'))

    pPr.append(change)


def _add_rpr_change(run, orig_rpr):
    """将原始字符格式嵌入 <w:rPrChange>"""
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn('w:rPrChange')):
        rPr.remove(old)

    change = OxmlElement('w:rPrChange')
    change.set(qn('w:id'), str(_next_rev_id()))
    change.set(qn('w:author'), '公文格式工具')
    change.set(qn('w:date'), _rev_date())

    if orig_rpr is not None:
        snapshot = deepcopy(orig_rpr)
        for old in snapshot.findall(qn('w:rPrChange')):
            snapshot.remove(old)
        change.append(snapshot)
    else:
        change.append(OxmlElement('w:rPr'))

    rPr.append(change)


def reset_revision_counter():
    """每篇文档重置修订 ID 计数器"""
    _revision_counter[0] = 0
