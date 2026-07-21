# -*- coding: utf-8 -*-
"""目录生成：Word 域自动目录 / 手动格式化目录页"""
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _first_body_element(doc):
    """返回文档正文第一个可插入位置的元素，找不到返回 None"""
    body = doc._body._body
    children = list(body)
    if not children:
        return None
    # 跳过 sectPr
    for child in children:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag != 'sectPr':
            return child
    return children[0]


def _insert_paragraph_before(doc, ref_element, text=''):
    """在 ref_element 之前插入新段落，返回 Paragraph 对象"""
    from docx.text.paragraph import Paragraph
    body = doc._body._body
    p = OxmlElement('w:p')
    body.insert(list(body).index(ref_element), p)
    para = Paragraph(p, body)
    if text:
        para.text = text
    return para


def insert_auto_toc(doc, levels=3, title_text='目  录'):
    """在文首插入 Word 自动目录字段"""
    ref = _first_body_element(doc)
    if ref is None:
        return

    # 目录标题
    tp = _insert_paragraph_before(doc, ref, title_text)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.runs[0].font.size = Pt(22)
    tp.runs[0].font.bold = True

    # TOC 字段
    tp2 = _insert_paragraph_before(doc, ref)
    begin_run = tp2.add_run()
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    begin_run._r.append(begin)
    instr_run = tp2.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' TOC \\o "1-{}" \\h \\z \\u '.format(levels)
    instr_run._r.append(instr)
    end_run = tp2.add_run()
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    end_run._r.append(end)

    # 提示行
    note = _insert_paragraph_before(doc, ref,
        '（↑ 此目录为 Word 自动目录域，请在 Word/WPS 中右键点击 → 更新域，即可自动生成页码）')
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in note.runs:
        r.font.size = Pt(10)
        r.font.color.rgb = None

    # 分隔空行
    _insert_paragraph_before(doc, ref, '')


def _build_heading_items(doc):
    """扫描文档，返回 [(标题文本, 级别 1-4), ...]，保持文档顺序"""
    from scripts.detector import detect_para_type, _compile_rules, _build_text_context
    from scripts.formatter import PRESETS
    all_texts, idx_map = _build_text_context(doc)
    preset = PRESETS.get('official_gbk', PRESETS['official'])
    rules = _compile_rules(preset.get('detect_rules'))
    items = []
    prev_type = None
    total = len(doc.paragraphs)
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        ptype = detect_para_type(
            text, i, total, para.paragraph_format.alignment,
            all_texts, all_texts_index=idx_map.get(i),
            prev_para_type=prev_type, rules=rules,
        )
        prev_type = ptype
        level = None
        if ptype == 'heading1':       level = 1
        elif ptype == 'heading2':     level = 2
        elif ptype == 'heading3':     level = 3
        elif ptype == 'heading4':     level = 4
        elif ptype == 'title':        level = 0
        if level is not None:
            items.append((text, level))
    return items


def build_manual_toc(doc, title_text='目  录'):
    """扫描文档标题层级，在文首生成带点引导线和页码占位符的手动目录页"""
    items = _build_heading_items(doc)
    if not items:
        return

    ref = _first_body_element(doc)
    if ref is None:
        return

    indent_map = {0: 0, 1: 0, 2: 32, 3: 64, 4: 96}   # pt 缩进
    size_map = {0: 16, 1: 16, 2: 16, 3: 14, 4: 14}     # pt 字号
    DOTS_PER_CM = 8  # 每厘米约 8 个点

    # 目录标题
    tp = _insert_paragraph_before(doc, ref, title_text)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.runs[0].font.size = Pt(22)
    tp.runs[0].font.bold = True

    # 分隔
    _insert_paragraph_before(doc, ref, '')

    # 逐条插入目录项（顺序插入 = 文档顺序）
    for text, level in items:
        indent_pt = indent_map.get(level, 0)
        font_size = size_map.get(level, 14)
        bold = (level <= 1)

        # 估算点线长度：A4 可用宽度约 14cm，减去缩进和标题文字宽度
        # 中文字宽 ≈ 字号，标题约 20 字 → 剩余空间填点线
        available_cm = 14.0 - indent_pt / 72 * 2.54 - min(len(text) * font_size / 72 * 2.54, 10)
        dot_count = max(8, int(available_cm * DOTS_PER_CM))
        full_line = text + ' ' + '. ' * (dot_count // 2) + ' ___'

        p = _insert_paragraph_before(doc, ref, full_line)
        p.paragraph_format.first_line_indent = Pt(0)
        if indent_pt > 0:
            p.paragraph_format.left_indent = Pt(indent_pt)
        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold

    # 末尾提示
    _insert_paragraph_before(doc, ref, '')
    note = _insert_paragraph_before(doc, ref,
        '（此目录为程序自动生成，标题后的 "___" 为页码占位符，请在 Word/WPS 中手动填入实际页码）')
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in note.runs:
        r.font.size = Pt(10)
        r.font.color.rgb = None


def generate_toc(input_path, output_path, mode='auto', levels=3):
    """外部调用入口"""
    doc = Document(input_path)
    if mode == 'auto':
        insert_auto_toc(doc, levels=levels)
    else:
        build_manual_toc(doc)
    doc.save(output_path)
