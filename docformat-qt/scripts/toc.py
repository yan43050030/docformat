# -*- coding: utf-8 -*-
"""目录生成：插入 Word 自动目录字段 或 手动生成目录文本页"""
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _insert_paragraph_before_first_body(doc, text='', style_name=None):
    """在文档第一个非空段落之前插入新段落"""
    body = doc._body._body
    first = body[0]
    p = OxmlElement('w:p')
    body.insert(list(body).index(first), p)
    from docx.text.paragraph import Paragraph
    para = Paragraph(p, body)
    if text:
        para.text = text
    if style_name:
        para.style = doc.styles[style_name]
    return para


def insert_auto_toc(doc, levels=3, title_text='目  录'):
    """在文首插入 Word 自动目录字段（需在 Word/WPS 中右键更新域）"""
    # 插入目录标题
    title_para = _insert_paragraph_before_first_body(doc, title_text)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.runs[0]
    title_run.font.size = Pt(22)
    title_run.font.bold = True

    # 插入空行
    blank = _insert_paragraph_before_first_body(doc)

    # 插入 TOC 字段
    toc_para = _insert_paragraph_before_first_body(doc)

    # 构造 fldChar begin
    begin_run = toc_para.add_run()
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    begin_run._r.append(begin)

    # 构造 instrText
    instr_run = toc_para.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' TOC \\o "1-{}" \\h \\z \\u '.format(levels)
    instr_run._r.append(instr)

    # 构造 fldChar end
    end_run = toc_para.add_run()
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    end_run._r.append(end)

    # 目录后空一行
    _insert_paragraph_before_first_body(doc)


def _build_heading_texts(doc):
    """扫描文档，返回 [(标题文本, 标题级别)] 用于手动目录"""
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
            text, i, total,
            para.paragraph_format.alignment,
            all_texts,
            all_texts_index=idx_map.get(i),
            prev_para_type=prev_type, rules=rules,
        )
        prev_type = ptype
        level = None
        if ptype == 'heading1':
            level = 1
        elif ptype == 'heading2':
            level = 2
        elif ptype == 'heading3':
            level = 3
        elif ptype == 'heading4':
            level = 4
        elif ptype == 'title':
            level = 0  # 主标题
        if level is not None:
            items.append((text, level))
    return items


def build_manual_toc(doc, title_text='目  录'):
    """扫描文档标题层级，在文首生成手动目录页（不含页码，打开后手动填入）"""
    items = _build_heading_texts(doc)
    if not items:
        return

    # 目录标题
    title_para = _insert_paragraph_before_first_body(doc, title_text)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.runs[0]
    title_run.font.size = Pt(22)
    title_run.font.bold = True

    # 底部空行
    blank = _insert_paragraph_before_first_body(doc)

    # 逐条插入目录项
    for text, level in reversed(items):
        p = _insert_paragraph_before_first_body(doc, text)
        run = p.runs[0]
        if level <= 1:
            run.font.size = Pt(16)
            run.font.bold = True
            p.paragraph_format.first_line_indent = Pt(0)
        elif level == 2:
            run.font.size = Pt(16)
            run.font.bold = False
            p.paragraph_format.first_line_indent = Pt(32)
        else:
            run.font.size = Pt(14)
            run.font.bold = False
            p.paragraph_format.first_line_indent = Pt(64)

    # 底部空行+提示
    note = _insert_paragraph_before_first_body(doc, '（此目录为程序自动生成，页码请在 Word/WPS 中手动填入或右键更新域）')
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in note.runs:
        r.font.size = Pt(10)
        r.font.color.rgb = None  # 浅色


def generate_toc(input_path, output_path, mode='auto', levels=3):
    """外部调用入口：读取文档 → 插入目录 → 保存"""
    doc = Document(input_path)
    if mode == 'auto':
        insert_auto_toc(doc, levels=levels)
    else:
        build_manual_toc(doc)
    doc.save(output_path)
