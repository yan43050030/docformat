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


def lock_row_heights(doc):
    """把所有已设高度的行改为固定高度（hRule=exact），返回锁定的行数。

    套打的命门是几何绝对不能变。atLeast 行一旦内容超长，Word 会把行撑高、
    下面所有内容整体下移，与预印栏位全部错位——这比文字被裁掉严重得多。
    改成 exact 后 Word 物理上无法撑高，版面永远对得住预印纸；
    配合自适应缩字号，正常内容都能放下，实在放不下会明确告警。
    """
    n = 0
    for table in doc.tables:
        for row in table.rows:
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                continue
            he = trPr.find(qn('w:trHeight'))
            if he is None:
                continue
            if he.get(qn('w:hRule')) != 'exact':
                he.set(qn('w:hRule'), 'exact')
                n += 1
    return n


def fill_form(template_path, values, output_path, autofit=True, log=None,
              lock_heights=True):
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

    if lock_heights:
        locked = lock_row_heights(doc)
        if locked and log:
            log('info', '已锁定 {} 行为固定高度，保证与预印栏位对齐'.format(locked))

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


# ---------------- 从已有 docx 提取内容 ----------------

import datetime

# 字段的标签写法（同一字段可能写成多种样子，如"经办人"/"经 办 人"）
_FIELD_LABELS = {
    '紧急程度': ['紧急程度'],
    '密级': ['密级', '密  级'],
    '标题': ['标题', '标  题', '题目'],
    '拟办意见': ['拟办意见', '拟办意见'],
    '领导批示': ['领导批示'],
    '承办部门': ['承办部门', '承办单位'],
    '经办人': ['经办人', '经 办 人'],
    '电话': ['电话', '联系电话'],
    '文字校核': ['文字校核', '文字核校'],
}

# 允许标签里夹杂空格（"经 办 人：" 这类排版用的空格）
def _label_pattern(label):
    return r'\s*'.join(re.escape(ch) for ch in label if not ch.isspace())


_DATE_PATTERNS = [
    re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'),
    re.compile(r'(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})'),
]

_CN_DIGITS = {'〇': 0, '零': 0, '一': 1, '二': 2, '三': 3, '四': 4,
              '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
# 公文成文日期常写成中文数字：二〇二六年七月二十五日
_CN_DATE_RE = re.compile(
    r'([〇零一二三四五六七八九]{4})\s*年\s*([一二三四五六七八九十]{1,3})\s*月'
    r'(?:\s*([一二三四五六七八九十]{1,3})\s*日)?')


def _cn_year(text):
    n = 0
    for ch in text:
        if ch not in _CN_DIGITS:
            return None
        n = n * 10 + _CN_DIGITS[ch]
    return n


def _cn_number(text):
    """解析 一~三十一 这类中文数字（含十/十一/二十/二十五）"""
    if not text:
        return None
    if text == '十':
        return 10
    if text.startswith('十'):                      # 十一 ~ 十九
        rest = _CN_DIGITS.get(text[1:], None) if len(text) == 2 else None
        return 10 + rest if rest is not None else None
    if '十' in text:                                # 二十 / 二十五 / 三十一
        head, _sep, tail = text.partition('十')
        h = _CN_DIGITS.get(head)
        if h is None:
            return None
        if not tail:
            return h * 10
        t = _CN_DIGITS.get(tail)
        return h * 10 + t if t is not None else None
    return _CN_DIGITS.get(text)


def parse_date(text):
    """从文本里抽出 (年, 月, 日) 字符串；抽不到返回 None。

    套打模板里年/月/日是三个独立位置，必须拆开分别落位，
    直接把"2026年6月25日"整串塞进"年"格会把后面全顶歪。
    """
    for pat in _DATE_PATTERNS:
        m = pat.search(text or '')
        if m:
            return m.group(1), str(int(m.group(2))), str(int(m.group(3)))
    m = _CN_DATE_RE.search(text or '')
    if m:
        y = _cn_year(m.group(1))
        mo = _cn_number(m.group(2))
        day = _cn_number(m.group(3)) if m.group(3) else None
        if y and mo:
            return str(y), str(mo), (str(day) if day else '')
    return None


def _blocks_of(doc):
    """产出文档里所有文本块：(文本, 是否来自表格单元格, 单元格对象或None)"""
    for p in doc.paragraphs:
        yield p.text, False, None
    for t in doc.tables:
        for cell in _iter_cells(t):
            yield cell.text, True, cell


def _strip_placeholders(text):
    return PLACEHOLDER_RE.sub('', text or '')


def extract_values(source_path, fields=None):
    """从一份已有 docx 里按标签抽取各字段内容。

    适用于"同类表单的电子版"——不要求与模板结构完全一致，
    按标签文字定位，兼容"承办部门：X"写在段落里或单元格里两种情况。
    返回 {字段: 值}，抽不到的字段不出现在结果里。
    """
    from docx import Document
    from .paragraph import sanitize_document
    doc = Document(source_path)
    sanitize_document(doc)

    wanted = list(fields) if fields else list(_FIELD_LABELS.keys()) + ['年', '月', '日']
    values = {}
    blocks = [(_strip_placeholders(txt), in_cell, cell)
              for txt, in_cell, cell in _blocks_of(doc)]

    # --- 按标签抽取 ---
    for field in wanted:
        if field in ('年', '月', '日'):
            continue
        labels = _FIELD_LABELS.get(field, [field])
        got = None
        for label in labels:
            pat = re.compile(_label_pattern(label) + r'\s*[：:]\s*(.*)', re.S)
            for txt, _in_cell, _cell in blocks:
                if not txt.strip():
                    continue
                m = pat.search(txt)
                if not m:
                    continue
                val = m.group(1).strip()
                # 截到下一个标签处，避免把同一行后面的"密级：X"一起吞掉
                cut = len(val)
                for other_labels in _FIELD_LABELS.values():
                    for ol in other_labels:
                        om = re.search(_label_pattern(ol) + r'\s*[：:]', val)
                        if om and om.start() < cut:
                            cut = om.start()
                val = val[:cut].strip()
                if val:
                    got = val
                    break
            if got:
                break
        if got:
            values[field] = got

    # --- 标题：标签抽不到时，取"标题"标签相邻单元格 ---
    if '标题' not in values:
        cells = [(txt, cell) for txt, in_cell, cell in blocks if in_cell]
        for i, (txt, _c) in enumerate(cells):
            t = txt.strip()
            if t in ('标题', '标  题', '题目') and i + 1 < len(cells):
                cand = cells[i + 1][0].strip()
                if cand:
                    values['标题'] = re.sub(r'\s*\n\s*', '', cand)
                break

    # --- 长文本字段（拟办意见/领导批示）---
    # 常见两种写法：标签与正文同块的不同段落；或标签独占一段、正文在后续段落。
    # 后者若只在本块里找，会漏掉全部正文。
    _all_labels = [l for ls in _FIELD_LABELS.values() for l in ls]

    def _starts_with_label(text):
        head = text.strip()[:12]
        for lb in _all_labels:
            if re.match(_label_pattern(lb) + r'\s*[：:]', head):
                return True
        return False

    for field in ('拟办意见', '领导批示'):
        if field not in wanted:
            continue
        if len(str(values.get(field, ''))) >= 4:
            continue
        for bi, (txt, _in_cell, _cell) in enumerate(blocks):
            if field not in txt:
                continue
            after = txt.split(field, 1)[1]
            after = re.sub(r'^\s*[：:]\s*', '', after).strip()
            if len(after) >= 4:
                values[field] = after
                break
            # 标签独占一段：往后收集，直到遇到下一个标签或日期行
            collected = []
            for nxt, _ic, _c in blocks[bi + 1:]:
                t = nxt.strip()
                if not t:
                    continue
                if _starts_with_label(t) or parse_date(t):
                    break
                collected.append(t)
            if collected:
                values[field] = '\n'.join(collected)
            break

    # --- 日期：拆成年/月/日三格，避免整串塞进一格把版面顶歪 ---
    for txt, _in_cell, _cell in blocks:
        d = parse_date(txt)
        if d:
            values['年'], values['月'], values['日'] = d
            break

    return {k: v for k, v in values.items() if k in wanted and str(v).strip()}


def fit_document(source_path, template_path, output_path,
                 overrides=None, autofit=True, log=None):
    """把一份已有 docx 的内容适配到套打模板并输出。

    overrides: 手工修正/补充的字段，优先级高于自动抽取的值。
    返回 (提取到的值, 提示列表)。
    """
    fields = scan_fields(template_path)
    values = extract_values(source_path, fields)
    if overrides:
        values.update({k: v for k, v in overrides.items() if str(v).strip()})
    if log:
        got = '、'.join('{}={}'.format(k, str(v)[:12]) for k, v in sorted(values.items()))
        log('info', '套打适配：识别到 {} 个字段（{}）'.format(len(values), got or '无'))
    _n, notes = fill_form(template_path, values, output_path,
                          autofit=autofit, log=log)
    missing = [f for f in fields if not str(values.get(f, '')).strip()]
    if missing:
        notes = list(notes) + ['未能自动识别：{}（可在对话框里手工补填）'.format('、'.join(missing))]
    return values, notes
