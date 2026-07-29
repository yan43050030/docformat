"""OOXML paragraph settings for Chinese line-breaking rules in Word/WPS."""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap",
    "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
)
_PPR_ORDER_INDEX = {name: index for index, name in enumerate(_PPR_ORDER)}


def _insert_paragraph_property(p_pr, element, local_name):
    """Insert a property in the order expected by strict OOXML consumers."""
    property_index = _PPR_ORDER_INDEX[local_name]
    for index, child in enumerate(p_pr):
        child_name = child.tag.rsplit("}", 1)[-1]
        if _PPR_ORDER_INDEX.get(child_name, -1) > property_index:
            p_pr.insert(index, element)
            return
    p_pr.append(element)


def _set_paragraph_boolean_property(paragraph, local_name, value):
    p_pr = paragraph._p.get_or_add_pPr()
    tag = qn(f"w:{local_name}")
    element = p_pr.find(tag)
    if element is None:
        element = OxmlElement(f"w:{local_name}")
        _insert_paragraph_property(p_pr, element, local_name)
    desired_value = "1" if value else "0"
    if element.get(qn("w:val")) == desired_value:
        return False
    element.set(qn("w:val"), desired_value)
    return True


def set_outline_level(paragraph, level):
    """设置段落大纲级别（0-8 表示 1-9 级；None 表示正文级）。

    公文排版会把所有段落强制成「正文」样式，Word 的自动目录域因此取不到
    任何条目。设置 outlineLvl 后，目录域（含 \\u 开关）与导航窗格都能
    正确取到标题层级，更新域即可得到真实页码。
    """
    p_pr = paragraph._p.get_or_add_pPr()
    tag = qn("w:outlineLvl")
    element = p_pr.find(tag)
    if level is None:
        if element is not None:
            p_pr.remove(element)
            return True
        return False
    if element is None:
        element = OxmlElement("w:outlineLvl")
        _insert_paragraph_property(p_pr, element, "outlineLvl")
    value = str(int(level))
    if element.get(qn("w:val")) == value:
        return False
    element.set(qn("w:val"), value)
    return True


def apply_chinese_line_break_rules_to_paragraph(paragraph):
    """Enable Chinese kinsoku and disable hanging punctuation for one paragraph."""
    if not paragraph.text.strip():
        return False
    changed = _set_paragraph_boolean_property(paragraph, "kinsoku", True)
    return _set_paragraph_boolean_property(paragraph, "overflowPunct", False) or changed


def _iter_container_paragraphs(container, seen):
    for paragraph in container.paragraphs:
        paragraph_id = id(paragraph._p)
        if paragraph_id not in seen:
            seen.add(paragraph_id)
            yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell, seen)


def apply_chinese_line_break_rules(document):
    """Apply Chinese punctuation line-break protection across document stories.

    Body, nested tables, independent headers and footers are included. Existing
    nodes are updated in place, so repeated processing is idempotent.
    """
    seen = set()
    changed_count = sum(
        apply_chinese_line_break_rules_to_paragraph(paragraph)
        for paragraph in _iter_container_paragraphs(document, seen)
    )
    story_names = (
        "header", "first_page_header", "even_page_header",
        "footer", "first_page_footer", "even_page_footer",
    )
    for section in document.sections:
        for story_name in story_names:
            story = getattr(section, story_name)
            if story.is_linked_to_previous:
                continue
            changed_count += sum(
                apply_chinese_line_break_rules_to_paragraph(paragraph)
                for paragraph in _iter_container_paragraphs(story, seen)
            )
    return changed_count


# 半角括号：正文里它落在 ASCII 区，Word 一律按"西文字体"取字形，于是
# 一句中文里冒出两个 Times New Roman 的细括号，跟前后的仿宋对不上。
# w:hint="eastAsia" 在这里帮不上忙——Word 对 U+0000–U+007F 只认 w:ascii。
# 能真正改掉的办法只有一个：把这几个字符单拆成 run，把它的西文字体
# 也写成中文字体。字符本身不动（不替换成全角），排版之外的内容零改动。
_ASCII_PUNCT_AS_CN = '()'
_CJK = tuple(zip((0x2E80, 0x3000, 0x4E00, 0xF900, 0xFF00),
                 (0x2EFF, 0x303F, 0x9FFF, 0xFAFF, 0xFFEF)))


def _has_cjk(text):
    return any(any(lo <= ord(ch) <= hi for lo, hi in _CJK) for ch in text)


def _split_run_at_punct(run):
    """把 run 按「半角括号 / 其余」切开，返回是否改动过。"""
    from copy import deepcopy

    text = run.text
    if not any(ch in text for ch in _ASCII_PUNCT_AS_CN):
        return False
    rPr = run._r.find(qn('w:rPr'))
    fonts = rPr.find(qn('w:rFonts')) if rPr is not None else None
    cn = fonts.get(qn('w:eastAsia')) if fonts is not None else None
    if not cn:
        return False

    segs = []
    for ch in text:
        want_cn = ch in _ASCII_PUNCT_AS_CN
        if segs and segs[-1][0] == want_cn:
            segs[-1][1] += ch
        else:
            segs.append([want_cn, ch])
    if len(segs) == 1 and not segs[0][0]:
        return False

    r = run._r
    parent = r.getparent()
    at = list(parent).index(r)
    for offset, (want_cn, chunk) in enumerate(segs):
        new_r = deepcopy(r)
        for t in new_r.findall(qn('w:t')):
            new_r.remove(t)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = chunk
        new_r.append(t)
        if want_cn:
            new_rPr = new_r.find(qn('w:rPr'))
            new_fonts = new_rPr.find(qn('w:rFonts')) if new_rPr is not None else None
            if new_fonts is not None:
                new_fonts.set(qn('w:ascii'), cn)
                new_fonts.set(qn('w:hAnsi'), cn)
                new_fonts.set(qn('w:cs'), cn)
        parent.insert(at + offset, new_r)
    parent.remove(r)
    return True


def apply_cn_font_to_ascii_punctuation(document):
    """让半角括号跟着中文字体走。只处理含中文的段落。"""
    seen = set()
    changed = 0
    for paragraph in _iter_container_paragraphs(document, seen):
        if not _has_cjk(paragraph.text):
            continue
        for run in list(paragraph.runs):
            if _split_run_at_punct(run):
                changed += 1
    return changed
