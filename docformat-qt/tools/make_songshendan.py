# -*- coding: utf-8 -*-
"""按实测尺寸重建「文件送审单」套打模板。

所有数字来自用户拿直尺量真实预印纸的结果（见 SPEC），不是反推的。
改尺寸只改 SPEC，重跑本脚本即可——模板不再是"改一次错一次"的手工活。

坐标约定：一律用"距纸张左边 / 距纸张上边"的厘米数；用户从下边量的值
在 SPEC 里按 H - d 换算好，注释里保留原始说法便于核对。

垂直定位靠"精确行距 + 段前距"，横向定位靠**制表位**（缇为单位、与字体
无关）。生成后用 tools/check_songshendan.py 渲染成 PDF 逐项实测复核。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

W = 21.0                     # 纸宽（实测=标称）
# 纸高用标称 A4 的 29.7：用户尺子读数 29.6 是量纸误差，而"距纸下边多少"
# 这个关系必须相对**真实纸边**成立，所以按 29.7 换算（照 29.6 算会整体高 1mm）
H = 29.7
RED = 'CC2222'               # 仅生成"红头样张"时用；模板本身是白字

SPEC = {
    # ---- 文件头 ----
    'head1_top': 2.8,        # 单位名称行 上边线距纸上边
    'head1_left': 5.3,       # 距纸左侧
    'head1_right': 5.0,      # 距纸右侧
    'head1_pt': 15.0, 'head1_spacing_cm': 0.2, 'head1_gap_cm': 0.9,
    'head2_top': 3.7,        # 「文件送审单」上边线
    'head2_bottom': 4.5,     # 下边线
    'head2_left': 7.7, 'head2_right': 7.5,
    'head2_pt': 18.0, 'head2_spacing_cm': 0.656,   # 字高实测 0.80cm 反推

    # ---- 紧急程度 / 密级 ----
    'urgent_left': 2.5,      # 「紧急程度：」左边线
    'sec_right': 5.1,        # 「密级：」右边线距纸右侧
    'rule_after_urgent': 5.7,  # 该行下面的红线距纸上边

    # ---- 表格区（红线的位置就是表格行线）----
    'rule_after_title': 8.0,
    'title_text_top': 6.6, 'title_left': 2.4,
    'lead_text_top': 8.5, 'lead_left': 2.4, 'lead_right': 4.5,
    'rule_before_opinion': H - 9.5,   # 距纸下面 9.5
    'opinion_text_top': H - 9.0,      # 距纸下面 9.0
    'rule_after_opinion': H - 4.9,    # 距纸下面 4.9
    'dept_text_top': H - 4.0,         # 承办部门 距纸下面 4.0
    'handler_text_top': H - 4.5,      # 经办人   距纸下面 4.5
    'rule_mid_right': H - 3.9,        # 经办人/文字校核之间 距纸下面 3.9
    'check_text_top': H - 3.5,        # 文字校核 距纸下面 3.5
    'rule_bottom': H - 2.9,           # 文字校核下面 距纸下面 2.9
    'vline_from_right': 12.0,         # 承办部门右侧竖线 距纸右侧
    'handler_right': 9.8,             # 「经办人：」右边线距纸右侧
    'handler_to_phone': 3.1,          # 经办人：与电话：之间
    'phone_right': 5.5,               # 「电话：」右边线距纸右侧
    'label_pt': 12.0,                 # 各栏目名字号（实测 11.9pt）
    'title_pt': 16.0,                 # 标题正文字号

    # ---- 成文日期 / 落款 ----
    'ymd_year_left': 4.4, 'ymd_month_left': 5.6, 'ymd_day_left': 7.0,
    'sign_left_from_right': 8.1,      # 落款首字左边线距纸右侧
    'sign_right_from_right': 2.2,     # 落款尾字右边线距纸右侧
    'sign_bottom_from_bottom': 2.2,   # 该行下边线距纸下面
    'sign_pt': 14.0, 'sign_text': '某地市某某单位的办公室制',

    # ---- 长红线左右 ----
    'rule_side': 2.1,
}

HEAD1_LEFT_TEXT = '中国某地市某单位'      # 8 字
HEAD1_RIGHT_TEXT = '某地市某单位'          # 6 字
HEAD2_TEXT = '文件送审单'


def _white(run):
    """预印在纸上的内容：白字占位，不显影但占准位置"""
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _set_spacing(run, cm):
    """字距（w:spacing，单位二十分之一磅）"""
    rPr = run._r.get_or_add_rPr()
    el = OxmlElement('w:spacing')
    el.set(qn('w:val'), str(int(round(cm * 28.3465 * 20))))
    rPr.append(el)


def _exact_line(para, pt, before_pt=0.0):
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(pt)
    # 段前距不能为负（OOXML 不允许），钳到 0；真要往上挪得靠上一个元素让位
    pf.space_before = Pt(max(0.0, before_pt))
    pf.space_after = Pt(0)
    # 关掉网格吸附，否则行高会被文档网格改写、纵向位置全乱。
    # 必须按 schema 顺序插入：w:pPr 的子元素次序是有约束的，直接 append
    # 会排到 w:spacing 之后，Word/LO 可能整段忽略后续设置（实测制表位失效）。
    pPr = para._p.get_or_add_pPr()
    sg = OxmlElement('w:snapToGrid')
    sg.set(qn('w:val'), '0')
    pPr.insert_element_before(
        sg, 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
        'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
        'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle',
        'w:rPr', 'w:sectPr', 'w:pPrChange')
    return para


def _tabs(para, stops):
    """stops: [(cm_from_left_margin, 'left'|'right')]"""
    ts = para.paragraph_format.tab_stops
    for pos, kind in stops:
        ts.add_tab_stop(Cm(pos), WD_TAB_ALIGNMENT.RIGHT if kind == 'right'
                        else WD_TAB_ALIGNMENT.LEFT)


def _cell_borders(cell, **sides):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:' + side)
        val = sides.get(side)
        if val:
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '8')
            el.set(qn('w:color'), 'FFFFFF')   # 白线：占位不显影
        else:
            el.set(qn('w:val'), 'none')
        b.append(el)
    tcPr.append(b)


def _row_height(row, cm):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:trHeight')
    h.set(qn('w:val'), str(int(round(cm * 566.93))))
    h.set(qn('w:hRule'), 'exact')
    trPr.append(h)


CALIB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'songshendan_calib.json')


def load_calib():
    """纵向修正量（cm）。字面高度与行盒顶端的差随字体、字号而变，
    与其猜公式，不如量出来：build → 渲染 → 实测 → 回填，迭代收敛。"""
    import json
    try:
        with open(CALIB_PATH, 'r', encoding='utf-8') as f:
            return {k: float(v) for k, v in json.load(f).items()}
    except (IOError, OSError, ValueError):
        return {}


def build(path, top_margin_cm=None, calib=None):
    S = SPEC
    C = load_calib() if calib is None else calib

    def cal(key):
        return C.get(key, 0.0)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(W), Cm(29.7)   # 打印用标称 A4
    sec.left_margin = Cm(S['rule_side'])
    sec.right_margin = Cm(S['rule_side'])
    sec.top_margin = Cm(top_margin_cm if top_margin_cm is not None
                        else S['head1_top'] + cal('head1'))
    sec.bottom_margin = Cm(0.6)
    # 去掉文档网格，纵向才由我们说了算
    sectPr = sec._sectPr
    for g in sectPr.findall(qn('w:docGrid')):
        sectPr.remove(g)

    body_left = S['rule_side']          # 版心左边界（= 长红线左端）

    def rel(x_cm, origin=None):
        """纸面绝对横坐标 → 制表位坐标。

        制表位是相对**当前容器左沿**的：普通段落是左边距，表格单元格里
        则是该单元格的左沿。用错原点会整体偏出一大截（实测差 7cm）。
        """
        return x_cm - (body_left if origin is None else origin)

    # ---------- 文件头第一行：单位名称 ----------
    p = doc.add_paragraph()
    _exact_line(p, S['head1_pt'] * 1.0)
    _tabs(p, [(rel(S['head1_left']), 'left')])
    p.add_run('\t')
    r = p.add_run(HEAD1_LEFT_TEXT)
    r.font.size = Pt(S['head1_pt'])
    _white(r); _set_spacing(r, S['head1_spacing_cm'])
    # 两组之间的间距：用一个定宽空 run 顶开（0.9cm 减掉末字多出的字距）
    gap = p.add_run(' ')
    gap.font.size = Pt(S['head1_pt'])
    _white(gap); _set_spacing(gap, S['head1_gap_cm'] - 0.5 * S['head1_pt'] / 28.3465
                            + cal('head1_gap'))
    r = p.add_run(HEAD1_RIGHT_TEXT)
    r.font.size = Pt(S['head1_pt'])
    _white(r); _set_spacing(r, S['head1_spacing_cm'])

    # ---------- 文件头第二行：文件送审单 ----------
    p = doc.add_paragraph()
    _exact_line(p, S['head2_pt'] * 1.0,
                before_pt=(S['head2_top'] - S['head1_top']
                           - S['head1_pt'] / 28.3465) * 28.3465 + cal('head2') * 28.3465)
    _tabs(p, [(rel(S['head2_left']), 'left')])
    p.add_run('\t')
    r = p.add_run(HEAD2_TEXT)
    r.font.size = Pt(S['head2_pt'])
    _white(r); _set_spacing(r, S['head2_spacing_cm'])

    # ---------- 紧急程度 / 密级 ----------
    lp = S['label_pt']
    p = doc.add_paragraph()
    _exact_line(p, lp * 1.4,
                before_pt=(S['rule_after_urgent'] - 0.55 - S['head2_bottom'])
                * 28.3465 + cal('urgent') * 28.3465)
    sec_w = 3 * lp / 28.3465            # 「密级：」三个字
    _tabs(p, [(rel(S['urgent_left']), 'left'),
              (rel(W - S['sec_right'] - sec_w), 'left')])
    p.add_run('\t')
    r = p.add_run('紧急程度：'); r.font.size = Pt(lp); _white(r)
    r = p.add_run('{{紧急程度}}'); r.font.size = Pt(lp)
    p.add_run('\t')
    r = p.add_run('密级：'); r.font.size = Pt(lp); _white(r)
    r = p.add_run('{{密级}}'); r.font.size = Pt(lp)

    # ---------- 表格 ----------
    rows = [
        ('title', S['rule_after_title'] - S['rule_after_urgent']),
        ('lead', S['rule_before_opinion'] - S['rule_after_title']),
        ('opinion', S['rule_after_opinion'] - S['rule_before_opinion']),
        ('dept1', S['rule_mid_right'] - S['rule_after_opinion']),
        ('dept2', S['rule_bottom'] - S['rule_mid_right']),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    col_l = S['vline_from_right']            # 左列宽 = 竖线位置 - 版心左
    left_w = (W - S['vline_from_right']) - body_left
    right_w = (W - S['rule_side']) - (W - S['vline_from_right'])
    grid = table._tbl.find(qn('w:tblGrid'))
    for gc, wcm in zip(grid.findall(qn('w:gridCol')), (left_w, right_w)):
        gc.set(qn('w:w'), str(int(round(wcm * 566.93))))
    # 表格整体不缩进。关键是把**表级**单元格内边距归零：Word/LO 会把
    # 表格左沿放在"左边距 − 单元格左内边距"处，默认 0.19cm，
    # 实测整张表左移到 1.91cm（应为 2.10），表内所有文字跟着差 0.19。
    tblPr = table._tbl.find(qn('w:tblPr'))
    # 表级白色框线：套打识别（cleaner.looks_like_overprint）和预览画线
    # 都读 tblBorders，只设单元格级的话它们看不到
    tb = OxmlElement('w:tblBorders')
    for side, on in (('top', True), ('left', False), ('bottom', True),
                     ('right', False), ('insideH', True), ('insideV', True)):
        e = OxmlElement('w:' + side)
        if on:
            e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '8')
            e.set(qn('w:color'), 'FFFFFF')
        else:
            e.set(qn('w:val'), 'none')
        tb.append(e)
    tblPr.append(tb)
    cm0 = OxmlElement('w:tblCellMar')
    for side in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + side)
        e.set(qn('w:w'), '0'); e.set(qn('w:type'), 'dxa')
        cm0.append(e)
    tblPr.append(cm0)
    ind = OxmlElement('w:tblInd')
    ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
    tblPr.append(ind)

    for i, (kind, h) in enumerate(rows):
        _row_height(table.rows[i], h)
        # 显式写 tcW：只改 tblGrid 的话，读宽度的代码（自适应、预览）
        # 会退回"平均分"，与真实版面不符
        for c, wcm in zip(table.rows[i].cells, (left_w, right_w)):
            c.width = Cm(wcm)
        for c in table.rows[i].cells:
            c.vertical_alignment = None
            tcPr = c._tc.get_or_add_tcPr()
            mar = OxmlElement('w:tcMar')
            for side in ('top', 'left', 'bottom', 'right'):
                e = OxmlElement('w:' + side)
                e.set(qn('w:w'), '0'); e.set(qn('w:type'), 'dxa')
                mar.append(e)
            tcPr.append(mar)

    def merge_row(i):
        c = table.rows[i].cells[0].merge(table.rows[i].cells[1])
        return c

    # 标题行：整行合并
    c = merge_row(0)
    _cell_borders(c, top=True, bottom=True)
    p = c.paragraphs[0]
    # 行距按**标题正文**字号设，不能按栏目名的 12pt：标题回成两行时
    # 12pt 的行距装不下 16pt 的字，两行会叠在一起还压出栏外
    _exact_line(p, S['title_pt'] * 1.0,
                before_pt=(S['title_text_top'] - S['rule_after_urgent'])
                * 28.3465 + cal('title') * 28.3465)
    # 悬挂缩进：第一行让出"标  题"这个栏目名的位置，回行的第二行
    # 缩进到标题正文的起点，不然会退回格子左沿、跑到栏目名底下
    # 3.28 是「标  题  」在渲染里的实测宽度（全角单位）——里面的空格
    # 比半个汉字窄，按字数算会多缩进 1cm 以上
    _hang = (S['title_left'] - body_left) + 3.28 * lp / 28.3465
    p.paragraph_format.left_indent = Cm(_hang)
    p.paragraph_format.first_line_indent = Cm(-_hang)
    _tabs(p, [(rel(S['title_left']), 'left')])
    p.add_run('\t')
    r = p.add_run('标  题'); r.font.size = Pt(lp); _white(r)
    r = p.add_run('  '); r.font.size = Pt(lp); _white(r)
    r = p.add_run('{{标题}}'); r.font.size = Pt(S['title_pt'])

    # 领导批示行
    c = merge_row(1)
    _cell_borders(c, bottom=True)
    p = c.paragraphs[0]
    _exact_line(p, lp * 1.0,
                before_pt=(S['lead_text_top'] - S['rule_after_title']) * 28.3465 + cal('lead') * 28.3465)
    _tabs(p, [(rel(S['lead_left']), 'left')])
    p.add_run('\t')
    r = p.add_run('领导批示：'); r.font.size = Pt(lp); _white(r)

    # 拟办意见行
    c = merge_row(2)
    _cell_borders(c, bottom=True)
    p = c.paragraphs[0]
    _exact_line(p, lp * 1.0,
                before_pt=(S['opinion_text_top'] - S['rule_before_opinion'])
                * 28.3465 + cal('opinion') * 28.3465)
    _tabs(p, [(rel(S['lead_left']), 'left')])
    p.add_run('\t')
    r = p.add_run('拟办意见：'); r.font.size = Pt(lp); _white(r)
    p2 = c.add_paragraph()
    _exact_line(p2, 14 * 1.4)
    # 正文首行缩进两个字：公文行文惯例，拟办意见是成段的话
    p2.paragraph_format.first_line_indent = Cm(2 * 14 / 28.3465)
    r = p2.add_run('{{拟办意见}}'); r.font.size = Pt(14)

    # 承办部门（纵向合并两行）/ 经办人 / 文字校核
    left_cell = table.rows[3].cells[0].merge(table.rows[4].cells[0])
    _cell_borders(left_cell, bottom=True, right=True)
    p = left_cell.paragraphs[0]
    _exact_line(p, lp * 1.0,
                before_pt=(S['dept_text_top'] - S['rule_after_opinion'])
                * 28.3465 + cal('dept') * 28.3465)
    _tabs(p, [(rel(S['lead_left']), 'left')])
    p.add_run('\t')
    r = p.add_run('承办部门：'); r.font.size = Pt(lp); _white(r)
    r = p.add_run('{{承办部门}}'); r.font.size = Pt(lp)

    hl = W - S['handler_right'] - 4 * lp / 28.3465      # 「经 办 人：」左沿
    ph = W - S['phone_right'] - 3 * lp / 28.3465        # 「电话：」左沿
    for ri, (label, key, top_cm, prev_rule) in (
            (3, ('经 办 人：', '经办人', S['handler_text_top'], S['rule_after_opinion'])),
            (4, ('文字校核：', '文字校核', S['check_text_top'], S['rule_mid_right']))):
        c = table.rows[ri].cells[1]
        _cell_borders(c, bottom=True)
        p = c.paragraphs[0]
        _exact_line(p, lp * 1.0,
                    before_pt=(top_cm - prev_rule + cal('handler')) * 28.3465)
        cell_left = W - S['vline_from_right']      # 右列单元格左沿
        # 用户量的是「经办人：」「电话：」的**右沿**（含冒号）。栏目名里
        # 有没有空格、多宽，靠字数猜不准，改由实测偏差校准左制表位。
        ckey = 'handler_x' if ri == 3 else 'check_x'   # 校准键，别和字段名 key 撞
        stops = [(rel(hl + cal(ckey), cell_left), 'left')]
        if ri == 3:
            stops.append((rel(ph + cal('phone_x'), cell_left), 'left'))
        _tabs(p, stops)
        p.add_run('\t')
        r = p.add_run(label); r.font.size = Pt(lp); _white(r)
        r = p.add_run('{{%s}}' % key); r.font.size = Pt(lp)
        if ri == 3:
            p.add_run('\t')
            r = p.add_run('电话：'); r.font.size = Pt(lp); _white(r)
            r = p.add_run('{{电话}}'); r.font.size = Pt(lp)

    # ---------- 成文日期 + 落款 ----------
    p = doc.add_paragraph()
    _exact_line(p, S['sign_pt'] * 1.0,
                before_pt=((H - S['sign_bottom_from_bottom'] - S['sign_pt'] / 28.3465)
                           - S['rule_bottom'] + cal('sign')) * 28.3465)
    sign_left = W - S['sign_left_from_right']
    # 数字用**右对齐**制表位顶到预印「年/月/日」的左沿，预印字紧随其后：
    # 这样一位数两位数都自动贴齐，不必按位数补空格
    ymd = [S['ymd_year_left'], S['ymd_month_left'], S['ymd_day_left']]
    stops = []
    for x in ymd:
        stops.append((rel(x), 'right'))        # 数字右沿顶到预印字左沿
        stops.append((rel(x + 0.01), 'left'))  # 预印字本身；错开 0.01cm 免得并成一个
    stops.append((rel(sign_left), 'left'))
    _tabs(p, stops)
    for key, ch in (('年', '年'), ('月', '月'), ('日', '日')):
        p.add_run('\t')
        r = p.add_run('{{%s}}' % key); r.font.size = Pt(S['sign_pt'])
        p.add_run('\t')
        r = p.add_run(ch); r.font.size = Pt(S['sign_pt']); _white(r)
    p.add_run('\t')
    r = p.add_run(S['sign_text']); r.font.size = Pt(S['sign_pt']); _white(r)

    doc.save(path)
    return path


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     'templates', '套打', '文件送审单.docx')
    build(out)
    print('已生成', out)
