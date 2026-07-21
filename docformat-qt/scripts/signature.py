# -*- coding: utf-8 -*-
"""
落款对位 — 从 formatter.py 拆分

GB/T 9704 公文标准落款规则：日期右空2字、署名与日期首字错2字
"""

import logging
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .table import _text_weight

logger = logging.getLogger('docformat.signature')


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
        # 孤行控制：日期行本身不被分页断开
        para.paragraph_format.keep_lines_together = True

        for s in sigs:
            s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            s.paragraph_format.right_indent = Pt(s_right * size_s)
            # 孤行控制：署名与下一段（日期）绑定不分页
            s.paragraph_format.keep_with_next = True
            s.paragraph_format.keep_lines_together = True
