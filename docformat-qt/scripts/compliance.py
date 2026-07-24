# -*- coding: utf-8 -*-
"""公文合规性检查：对照"当前预设"核对文档版式，报告偏差，并可对认可的偏差自动修正。

关键设计：检查标准来自用户选中的预设，而非死国标——用户改了预设，
检查标准自动跟着变，适用范围更广。每类检查可通过 options 开关启停。

交互式修正：check_compliance 会给"可自动修正"的偏差打上 fix_key，UI 让用户
勾选认可哪些，再调 apply_compliance_fixes 只对认可项动手，其余保持原样。
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

# 可配置的检查项（键: 说明），供 UI 生成勾选面板
CHECK_ITEMS = [
    ('margins', '页边距是否符合预设'),
    ('paper', '纸张大小是否符合预设'),
    ('body_font', '正文字体字号是否符合预设'),
    ('title_center', '标题是否居中'),
    ('structure', '结构完整性（标题/主送机关/成文日期是否齐全）'),
    ('line_spacing', '正文行距是否符合预设'),
    ('punctuation', '是否残留英文标点'),
]

DEFAULT_OPTIONS = {k: True for k, _ in CHECK_ITEMS}

# 各 fix_key 的一句话说明，供修正对话框和结果提示复用
FIX_LABELS = {
    'margins': '把页边距改为预设值',
    'paper': '把纸张改为 A4',
    'body_font': '把正文字体字号改为预设值',
    'title_center': '把标题设为居中',
    'line_spacing': '把正文行距改为预设固定值',
    'punctuation': '修复英文/不规范标点',
}


def _cm(v):
    return round(v, 2) if v is not None else None


def _detect_types(doc, preset):
    """返回 [(paragraph, ptype)]，仅含非空段，供检查与修正共用。"""
    from .detector import detect_para_type, _compile_rules, _build_text_context
    rules = _compile_rules(preset.get('detect_rules'))
    all_texts, idx_map = _build_text_context(doc)
    result = []
    prev = None
    total = len(doc.paragraphs)
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        ai = idx_map.get(i)
        ptype = detect_para_type(t, i, total, p.paragraph_format.alignment,
                                 all_texts, all_texts_index=ai, prev_para_type=prev,
                                 rules=rules)
        result.append((p, ptype))
        prev = ptype
    return result


def check_compliance(doc, preset, options=None, detect_types=None):
    """返回 findings 列表：
    [{'level': 'warn'/'info'/'ok', 'item': 名称, 'detail': 说明, 'fix_key': 可选}]

    fix_key 存在表示该偏差可被 apply_compliance_fixes 自动修正。
    detect_types: 兼容旧参数，当前未使用（内部自行识别）。
    """
    opts = dict(DEFAULT_OPTIONS)
    if options:
        opts.update(options)
    findings = []

    def add(level, item, detail, fix_key=None):
        f = {'level': level, 'item': item, 'detail': detail}
        if fix_key:
            f['fix_key'] = fix_key
        findings.append(f)

    sec = doc.sections[0]
    page = preset.get('page', {})

    # 页边距
    if opts.get('margins'):
        exp = (page.get('top'), page.get('bottom'), page.get('left'), page.get('right'))
        got = (_cm(sec.top_margin.cm), _cm(sec.bottom_margin.cm),
               _cm(sec.left_margin.cm), _cm(sec.right_margin.cm))
        bad = [n for n, e, g in zip(('上', '下', '左', '右'), exp, got)
               if e is not None and abs((g or 0) - e) > 0.05]
        if bad:
            add('warn', '页边距',
                '实际 上{}/下{}/左{}/右{} cm，预设要求 上{}/下{}/左{}/右{} cm，'
                '不符：{}'.format(got[0], got[1], got[2], got[3],
                                exp[0], exp[1], exp[2], exp[3], '、'.join(bad)),
                fix_key='margins')
        else:
            add('ok', '页边距', '符合预设')

    # 纸张大小
    if opts.get('paper'):
        w, h = _cm(sec.page_width.cm), _cm(sec.page_height.cm)
        want = preset.get('page_size', 'A4')
        is_a4 = abs((w or 0) - 21.0) < 0.2 and abs((h or 0) - 29.7) < 0.2
        if want == 'A4' and not is_a4:
            add('warn', '纸张', '当前 {}×{} cm 非 A4（预设要求 A4）'.format(w, h),
                fix_key='paper')
        else:
            add('ok', '纸张', '{}×{} cm'.format(w, h))

    # 正文字体字号（抽样首个正文段）
    if opts.get('body_font') or opts.get('line_spacing'):
        body = preset.get('body', {})
        body_para = None
        for p in doc.paragraphs:
            t = p.text.strip()
            if len(t) > 15 and p.runs:
                body_para = p
                break
        if body_para is not None:
            if opts.get('body_font'):
                run = body_para.runs[0]
                rpr = run._element.rPr
                ea = rpr.rFonts.get(qn('w:eastAsia')) if rpr is not None and rpr.rFonts is not None else None
                size = run.font.size.pt if run.font.size else None
                exp_font = body.get('font_cn'); exp_size = body.get('size')
                if ea and exp_font and ea != exp_font:
                    add('warn', '正文字体', '实际「{}」，预设要求「{}」'.format(ea, exp_font),
                        fix_key='body_font')
                elif size and exp_size and abs(size - exp_size) > 0.3:
                    add('warn', '正文字号', '实际 {}pt，预设要求 {}pt'.format(size, exp_size),
                        fix_key='body_font')
                else:
                    add('ok', '正文字体字号', '符合预设')
            if opts.get('line_spacing'):
                exp_ls = body.get('line_spacing')
                ls = body_para.paragraph_format.line_spacing
                if exp_ls:
                    got_pt = ls.pt if hasattr(ls, 'pt') else None
                    if got_pt and abs(got_pt - exp_ls) > 1:
                        add('warn', '正文行距', '实际约 {}pt，预设要求固定 {}pt'.format(round(got_pt, 1), exp_ls),
                            fix_key='line_spacing')
                    else:
                        add('ok', '正文行距', '符合预设')

    # 结构完整性 + 标题居中
    if opts.get('structure') or opts.get('title_center'):
        typed = _detect_types(doc, preset)
        types = {}
        title_para = None
        for p, ptype in typed:
            types[ptype] = types.get(ptype, 0) + 1
            if ptype == 'title' and title_para is None:
                title_para = p
        if opts.get('structure'):
            missing = []
            if not types.get('title'):
                missing.append('标题')
            if not types.get('recipient'):
                missing.append('主送机关')
            if not types.get('date'):
                missing.append('成文日期')
            if missing:
                # 结构缺失无法凭空补出内容，不给 fix_key（不可自动修正）
                add('warn', '结构完整性', '未识别到：{}（如确有请在预览里核对/指定）'.format('、'.join(missing)))
            else:
                add('ok', '结构完整性', '标题/主送机关/成文日期齐全')
        if opts.get('title_center') and title_para is not None:
            al = title_para.paragraph_format.alignment
            if al != WD_ALIGN_PARAGRAPH.CENTER:
                add('warn', '标题居中', '标题当前非居中', fix_key='title_center')
            else:
                add('ok', '标题居中', '标题居中')

    # 残留英文标点
    if opts.get('punctuation'):
        from . import analyzer
        issues = analyzer.analyze_punctuation(doc)
        if issues:
            add('warn', '标点规范', '发现 {} 处英文/不规范标点'.format(len(issues)),
                fix_key='punctuation')
        else:
            add('ok', '标点规范', '未发现英文标点')

    return findings


# ---------------- 自动修正 ----------------

def _fix_margins(doc, preset):
    page = preset.get('page', {})
    sec = doc.sections[0]
    done = []
    for attr, key, label in (('top_margin', 'top', '上'), ('bottom_margin', 'bottom', '下'),
                             ('left_margin', 'left', '左'), ('right_margin', 'right', '右')):
        v = page.get(key)
        if v is not None:
            setattr(sec, attr, Cm(v))
            done.append(label)
    return '页边距→上{}/下{}/左{}/右{} cm'.format(
        page.get('top', '-'), page.get('bottom', '-'),
        page.get('left', '-'), page.get('right', '-')) if done else ''


def _fix_paper(doc, preset):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    return '纸张→A4（21×29.7 cm）'


def _fix_title_center(doc, preset, typed=None):
    typed = typed if typed is not None else _detect_types(doc, preset)
    for p, ptype in typed:
        if ptype == 'title':
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            return '标题→居中'
    return ''


def _fix_line_spacing(doc, preset, typed=None):
    from .paragraph import paragraph_has_media
    exp_ls = preset.get('body', {}).get('line_spacing')
    if not exp_ls:
        return ''
    typed = typed if typed is not None else _detect_types(doc, preset)
    n = 0
    for p, ptype in typed:
        if ptype != 'body':
            continue
        if paragraph_has_media(p):   # 图片段落不动，固定行距会裁切图片
            continue
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(exp_ls)
        n += 1
    return '正文行距→固定 {}pt（{} 段）'.format(exp_ls, n) if n else ''


def _fix_body_font(doc, preset, typed=None):
    from .font import set_font
    body = preset.get('body', {})
    font_cn = body.get('font_cn'); font_en = body.get('font_en', font_cn)
    size = body.get('size')
    if not font_cn or not size:
        return ''
    typed = typed if typed is not None else _detect_types(doc, preset)
    n = 0
    for p, ptype in typed:
        if ptype != 'body':
            continue
        for run in p.runs:
            if not run.text.strip():
                continue
            set_font(run, font_cn, font_en, size, bold=body.get('bold', False))
        n += 1
    return '正文字体→{} {}pt（{} 段）'.format(font_cn, size, n) if n else ''


def _fix_punctuation(doc, preset):
    from . import punctuation
    quote_state = {'dq': 0, 'sq': 0}
    n = 0
    for p in doc.paragraphs:
        if punctuation.process_paragraph(p, quote_state=quote_state):
            n += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if punctuation.process_paragraph(p, quote_state=quote_state):
                        n += 1
    return '标点修复（{} 段有改动）'.format(n) if n else '标点修复（无需改动）'


# 需要段落类型识别的修正，接收共享的 typed 以免重复识别
_FIXERS = {
    'margins': (_fix_margins, False),
    'paper': (_fix_paper, False),
    'title_center': (_fix_title_center, True),
    'line_spacing': (_fix_line_spacing, True),
    'body_font': (_fix_body_font, True),
    'punctuation': (_fix_punctuation, False),
}


def apply_compliance_fixes(input_path, output_path, preset, fix_keys):
    """打开文档，仅对 fix_keys 指定的偏差自动修正，另存为 output_path。

    返回已执行修正的说明列表（每项一句话），供 UI 反馈。
    """
    from docx import Document
    from .paragraph import sanitize_document
    doc = Document(input_path)
    sanitize_document(doc)

    keys = [k for k in fix_keys if k in _FIXERS]
    needs_typed = any(_FIXERS[k][1] for k in keys)
    typed = _detect_types(doc, preset) if needs_typed else None

    applied = []
    for key in keys:
        fn, uses_typed = _FIXERS[key]
        try:
            desc = fn(doc, preset, typed) if uses_typed else fn(doc, preset)
        except Exception as e:
            applied.append('{}：修正失败（{}）'.format(FIX_LABELS.get(key, key), e))
            continue
        if desc:
            applied.append(desc)

    doc.save(output_path)
    return applied


def format_compliance_report(filename, findings, preset_name=''):
    warns = [f for f in findings if f['level'] == 'warn']
    lines = ['◆ 公文合规性检查：{}'.format(filename)]
    if preset_name:
        lines.append('  对照预设：{}'.format(preset_name))
    lines.append('  {}'.format('存在 {} 项偏差'.format(len(warns)) if warns else '未发现偏差 ✓'))
    lines.append('')
    for f in findings:
        mark = {'warn': '✗', 'ok': '✓', 'info': '·'}.get(f['level'], '·')
        lines.append('  {} 【{}】{}'.format(mark, f['item'], f['detail']))
    return '\n'.join(lines)
