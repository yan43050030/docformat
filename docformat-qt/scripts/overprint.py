# -*- coding: utf-8 -*-
"""套打（套头打印）：把内容精确打到预印红头纸的空白位置上。

套打的实现机制
--------------
预印在纸上的红色内容（单位名、栏目名、表格红线）在 docx 里写成
**白色文字与白色边框**——占住完全相同的位置、保证版式分毫不差，
但白纸上打印时不显影；只有**黑色文字**才真正印到预印纸上。

因此套打模板 = 一份保留全部白色占位与几何的 docx，其中要填的位置
放 {{字段名}} 占位符。填充时只替换占位符文本、不动任何几何，
套准度就由模板本身保证。

内容自适应
----------
预留格子高度固定（trHeight 为 exact 时更严格）。正文过长会把行撑高、
把整张表挤下去，套打就错位了。故填充长文本时按可用高度估算所需行数，
不够就逐档缩小字号与行距，直到放得下（有下限，缩到底仍放不下会如实告警）。
"""
import logging
import os
import re

from docx.oxml.ns import qn
from docx.shared import Pt

logger = logging.getLogger('docformat.overprint')

PLACEHOLDER_RE = re.compile(r'\{\{\s*([^}]+?)\s*\}\}')

# 自适应下限：再小就影响阅读，宁可告警也不继续缩
MIN_FONT_PT = 9.0
MIN_LINE_SPACING_PT = 12.0
# 每档缩小的幅度
FONT_STEP = 0.5

TWIPS_PER_CM = 566.93
PT_PER_CM = 28.3465


def _iter_cells(table):
    """产出表格里每个不重复的单元格。

    不能用 id(cell._tc) 去重——lxml 代理对象被回收后 id 会复用，
    不同单元格会被误判为同一个而漏处理。用保持引用的身份比较。
    """
    seen = []
    for row in table.rows:
        for cell in row.cells:
            if any(cell._tc is x for x in seen):
                continue
            seen.append(cell._tc)
            yield cell


def _iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p, None
    for t in doc.tables:
        for cell in _iter_cells(t):
            for p in cell.paragraphs:
                yield p, cell


def scan_fields(template_path):
    """扫描模板里的 {{字段}}，按出现顺序返回去重后的字段名列表。"""
    from docx import Document
    doc = Document(template_path)
    names = []
    for p, _cell in _iter_paragraphs(doc):
        for m in PLACEHOLDER_RE.finditer(p.text):
            name = m.group(1).strip()
            if name not in names:
                names.append(name)
    return names


def _replace_in_paragraph(para, values):
    """把段落里的占位符替换为取值，保留该 run 的字体/颜色/字号。

    占位符在模板里已被合并进单个 run（模板生成时保证），
    因此逐 run 替换即可，不必跨 run 拼接。
    """
    changed = False
    for run in para.runs:
        if '{{' not in run.text:
            continue
        def _sub(m):
            return str(values.get(m.group(1).strip(), ''))
        new = PLACEHOLDER_RE.sub(_sub, run.text)
        if new != run.text:
            run.text = new
            changed = True
    return changed


def _row_of_cell(table, cell):
    for row in table.rows:
        for c in row.cells:
            if c._tc is cell._tc:
                return row
    return None


def _row_height_cm(row):
    """返回 (高度cm, 是否为固定高度)；没设高度返回 (None, False)"""
    trPr = row._tr.find(qn('w:trPr'))
    if trPr is None:
        return None, False
    he = trPr.find(qn('w:trHeight'))
    if he is None:
        return None, False
    try:
        val = int(he.get(qn('w:val')))
    except (TypeError, ValueError):
        return None, False
    rule = he.get(qn('w:hRule')) or 'atLeast'
    return val / TWIPS_PER_CM, rule == 'exact'


def _grid_widths(table):
    grid = table._tbl.find(qn('w:tblGrid'))
    out = []
    if grid is not None:
        for gc in grid.findall(qn('w:gridCol')):
            try:
                out.append(int(gc.get(qn('w:w'))))
            except (TypeError, ValueError):
                out.append(0)
    return out


def _cell_width_cm(table, cell):
    """单元格实际宽度（cm），横向合并时按 gridSpan 累加所跨列。

    只读 tcW 会把合并单元格当成首列宽度，导致宽度被严重低估、
    进而误判"放不下"而缩字号——套打里这等于自己把版式弄歪。
    """
    tcPr = cell._tc.find(qn('w:tcPr'))
    span = 1
    if tcPr is not None:
        gs = tcPr.find(qn('w:gridSpan'))
        if gs is not None:
            try:
                span = max(1, int(gs.get(qn('w:val'))))
            except (TypeError, ValueError):
                span = 1
    widths = _grid_widths(table)
    if span > 1 and widths:
        # 定位该单元格起始列，累加它跨越的列宽
        for row in table.rows:
            col = 0
            for c in row.cells:
                if c._tc is cell._tc:
                    return sum(widths[col:col + span]) / TWIPS_PER_CM
                sp = 1
                p2 = c._tc.find(qn('w:tcPr'))
                if p2 is not None:
                    g2 = p2.find(qn('w:gridSpan'))
                    if g2 is not None:
                        try:
                            sp = max(1, int(g2.get(qn('w:val'))))
                        except (TypeError, ValueError):
                            sp = 1
                col += sp
    if tcPr is not None:
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is not None and tcW.get(qn('w:type')) == 'dxa':
            try:
                w = int(tcW.get(qn('w:w')))
                if w > 0:
                    return w / TWIPS_PER_CM
            except (TypeError, ValueError):
                pass
    if widths:
        return sum(widths) / TWIPS_PER_CM
    return 16.0


def _cell_margins_cm(cell):
    """单元格左右内边距之和（cm），取不到按 0.19cm×2 估。"""
    tcPr = cell._tc.find(qn('w:tcPr'))
    if tcPr is not None:
        mar = tcPr.find(qn('w:tcMar'))
        if mar is not None:
            tot = 0
            for side in ('w:left', 'w:right'):
                el = mar.find(qn(side))
                if el is not None:
                    try:
                        tot += int(el.get(qn('w:w'))) / TWIPS_PER_CM
                    except (TypeError, ValueError):
                        pass
            if tot:
                return tot
    return 0.38


def _text_width_units(text):
    """文本宽度（以"全角字"为单位）：中日韩全角计 1，ASCII 计 0.5。

    把空格和数字也按全角算会大幅高估宽度，导致本来一行放得下的
    短字段被误判为两行而缩字号。
    """
    units = 0.0
    for ch in text:
        units += 1.0 if ord(ch) > 0x2E80 else 0.5
    return units


def estimate_lines(text, font_pt, usable_width_cm, first_indent_pt=0):
    """估算文本在给定字号/宽度下占多少行（全角字宽≈字号）。"""
    if not text:
        return 0
    char_w_cm = font_pt / PT_PER_CM
    if char_w_cm <= 0:
        return 1
    per_line = max(1.0, usable_width_cm / char_w_cm)
    lines = 0
    for i, seg in enumerate(text.split('\n')):
        # 行尾空格不会触发换行，计入会导致虚假的"多一行"
        units = _text_width_units(seg.rstrip())
        if i == 0 and first_indent_pt:
            units += first_indent_pt / font_pt
        lines += max(1, int(-(-units // per_line)))    # 向上取整
    return lines


def _para_font_pt(para, default=14.0):
    for r in para.runs:
        if r.font.size is not None:
            return r.font.size.pt
    return default


def _para_line_spacing_pt(para, font_pt):
    ls = para.paragraph_format.line_spacing
    if ls is None:
        return font_pt * 1.4
    return ls.pt if hasattr(ls, 'pt') else font_pt * float(ls)


def _set_para_size(para, font_pt, line_pt):
    from docx.enum.text import WD_LINE_SPACING
    for r in para.runs:
        r.font.size = Pt(font_pt)
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line_pt)


def autofit_cell(table, cell, warn=None):
    """让单元格内容放进预留高度：不够就逐档缩小字号与行距。

    返回 (是否缩过, 最终字号, 是否仍放不下)。
    """
    row = _row_of_cell(table, cell)
    if row is None:
        return False, None, False
    height_cm, is_exact = _row_height_cm(row)
    if not height_cm:
        return False, None, False

    paras = [p for p in cell.paragraphs if p.text.strip()]
    if not paras:
        return False, None, False

    width_cm = _cell_width_cm(table, cell) - _cell_margins_cm(cell)
    # 上下内边距按行高的一小比例估，固定扣 0.2cm 对 1.2cm 的窄行等于扣掉 17%
    avail_cm = height_cm - min(0.15, height_cm * 0.08)

    base_sizes = [_para_font_pt(p) for p in paras]
    orig = list(base_sizes)

    def total_height(scale):
        total = 0.0
        for p, base in zip(paras, base_sizes):
            fs = max(MIN_FONT_PT, base * scale)
            lsp = max(MIN_LINE_SPACING_PT, _para_line_spacing_pt(p, base) * scale)
            n = estimate_lines(p.text, fs, width_cm,
                               p.paragraph_format.first_line_indent.pt
                               if p.paragraph_format.first_line_indent else 0)
            total += n * lsp / PT_PER_CM
        return total

    # 容差按"超出的后果"区分：
    # - atLeast 行：内容超长会把行撑高、整张表往下挤 → 套打直接错位，
    #   必须收紧，略有超出就缩字号；
    # - exact 行：Word 固定行高、超出部分被裁切，几何不会变、
    #   套打对位不受影响，唯一代价是文字被切。字宽是估算值（误差几个
    #   百分点很正常），这里放宽容差，免得把本来印得好好的短字段
    #   无谓缩小、反而和预印栏位对不齐。
    tolerance = 1.35 if is_exact else 1.05
    if total_height(1.0) <= avail_cm * tolerance:
        return False, orig[0] if orig else None, False

    scale = 1.0
    while scale > 0.55:
        scale -= (FONT_STEP / max(orig)) if max(orig) else 0.05
        if total_height(scale) <= avail_cm:
            break
    fits = total_height(scale) <= avail_cm
    for p, base in zip(paras, base_sizes):
        fs = max(MIN_FONT_PT, round(base * scale * 2) / 2.0)
        lsp = max(MIN_LINE_SPACING_PT, _para_line_spacing_pt(p, base) * scale)
        _set_para_size(p, fs, lsp)
    final = max(MIN_FONT_PT, round(orig[0] * scale * 2) / 2.0) if orig else None
    if not fits and warn:
        warn('内容过长，字号已缩到 {}pt 仍可能超出预留高度，建议精简文字'.format(final))
    return True, final, not fits


def fill_form(template_path, values, output_path, autofit=True, log=None):
    """按 values 填充套打模板并另存，返回 (已填字段数, 提示列表)。"""
    from docx import Document
    doc = Document(template_path)
    notes = []

    filled = set()
    for p, _cell in _iter_paragraphs(doc):
        for m in PLACEHOLDER_RE.finditer(p.text):
            filled.add(m.group(1).strip())
        _replace_in_paragraph(p, values)

    if autofit:
        for t in doc.tables:
            for cell in _iter_cells(t):
                if not cell.text.strip():
                    continue

                def _warn(msg, _c=cell):
                    label = (_c.text.strip().splitlines() or [''])[0][:12]
                    notes.append('【{}】{}'.format(label, msg))

                shrunk, size, overflow = autofit_cell(t, cell, warn=_warn)
                if shrunk and log:
                    log('info', '套打自适应：{} 区字号调整为 {}pt{}'.format(
                        (cell.text.strip().splitlines() or [''])[0][:12],
                        size, '（仍偏长）' if overflow else ''))

    doc.save(output_path)
    used = [k for k in values if k in filled and str(values.get(k, '')).strip()]
    return len(used), notes


# ---------------- 模板发现 ----------------

def bundled_overprint_dir():
    """软件自带的套打模板目录"""
    import sys as _sys
    if getattr(_sys, 'frozen', False):
        base = getattr(_sys, '_MEIPASS', os.path.dirname(_sys.executable))
        return os.path.join(base, 'templates', '套打')
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(here, 'templates', '套打')


def user_overprint_dir():
    from app.template_common import config_dir
    d = os.path.join(config_dir(), 'overprint')
    return d


def list_templates():
    """返回 [(显示名, 路径, 是否自带)]，自带的排前面。"""
    out = []
    for d, builtin in ((bundled_overprint_dir(), True), (user_overprint_dir(), False)):
        if not os.path.isdir(d):
            continue
        for name in sorted(os.listdir(d)):
            if name.lower().endswith('.docx') and not name.startswith('~$'):
                out.append((os.path.splitext(name)[0], os.path.join(d, name), builtin))
    return out
