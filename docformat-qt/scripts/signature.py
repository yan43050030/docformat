# -*- coding: utf-8 -*-
"""
落款对位 — 从 formatter.py 拆分

GB/T 9704 公文标准落款规则：日期右空2字、署名与日期首字错2字
"""

import logging
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .table import _text_weight

logger = logging.getLogger('docformat.signature')


def _no_snap_to_grid(para):
    """落款这两行不吸附文档网格。

    版心宽度换算成字数一般不是整数，右对齐 + 右缩进的行会被网格吸到最近
    一格上：实测「右空 2 字」印出来是 2.9 字，署名与日期的 2 字错位也变成
    2.5 字。这两行的位置是标准硬性规定的，宁可让它脱离网格也要给准。
    """
    pPr = para._p.get_or_add_pPr()
    # w:pPr 的子元素次序有约束，必须插在后续元素之前，否则 Word/WPS
    # 会把后面的设置一并忽略（套打模板上踩过同一个坑）
    _AFTER_SNAP = ('w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
                   'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
                   'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle',
                   'w:rPr', 'w:sectPr', 'w:pPrChange')
    for old in pPr.findall(qn('w:snapToGrid')):
        pPr.remove(old)
    sg = OxmlElement('w:snapToGrid')
    sg.set(qn('w:val'), '0')
    pPr.insert_element_before(sg, *_AFTER_SNAP)

    # 中西文自动间距也要关掉。开着的时候「2026年7月29日」会在每个中英
    # 交界处多撑出约 1/6 字，整行实测比按"汉字 1 字、数字半字"算出来的
    # 宽 1 个字——署名与日期的"错 2 字"于是变成错 1 字。标准是按字数说的，
    # 那就让这两行的字宽如实等于字数。
    for tag in ('w:autoSpaceDE', 'w:autoSpaceDN'):
        for old in pPr.findall(qn(tag)):
            pPr.remove(old)
        el = OxmlElement(tag)
        el.set(qn('w:val'), '0')
        pPr.insert_element_before(el, 'w:bidi', 'w:adjustRightInd',
                                  'w:snapToGrid', *_AFTER_SNAP)


def _apply_gb_signature_layout(typed_entries, preset):
    """公文落款对位（图解标准）：

    无公章（gb_seal = False / 未设置）：
      - 成文日期长于（含等于）署名：日期右空 2 字，署名首字比日期首字左移 2 字
      - 成文日期短于署名：署名右空 2 字，日期首字比署名首字右移 2 字
    加盖公章（gb_seal = True）：
      - 成文日期右空 4 字
      - 发文机关署名以成文日期为准居中编排
    长度按字符宽度计（汉字 1 字、英文数字 0.5 字）。
    """
    seal = bool(preset.get('gb_seal'))

    for idx, (para, ptype) in enumerate(typed_entries):
        if ptype != 'date':
            continue
        sigs = []
        j = idx - 1
        while j >= 0 and typed_entries[j][1] == 'signature':
            sigs.append(typed_entries[j][0])
            j -= 1
        if not sigs:
            continue

        size_d = (preset.get('date', {}) or {}).get('size', 16) or 16
        size_s = (preset.get('signature', {}) or {}).get('size', 16) or 16
        dlen = _text_weight(para.text.strip())
        slen = max(_text_weight(s.text.strip()) for s in sigs)

        if seal:
            # 加盖公章：日期右空 4 字，署名居中于日期
            d_right = 4.0
            s_right = dlen + 4.0 - slen + (slen / 2.0) - (dlen / 2.0)
        elif dlen >= slen:
            d_right = 2.0
            s_right = dlen + 4.0 - slen
        else:
            s_right = 2.0
            d_right = slen - dlen

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para.paragraph_format.right_indent = Pt(d_right * size_d)
        _no_snap_to_grid(para)
        # 孤行控制：日期行本身不被分页断开
        if preset.get('widow_control', False):
            para.paragraph_format.keep_together = True

        for s in sigs:
            s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            s.paragraph_format.right_indent = Pt(s_right * size_s)
            _no_snap_to_grid(s)
            # 孤行控制：署名与下一段（日期）绑定不分页
            if preset.get('widow_control', False):
                s.paragraph_format.keep_with_next = True
                s.paragraph_format.keep_together = True
