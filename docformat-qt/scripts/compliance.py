# -*- coding: utf-8 -*-
"""公文合规性检查：对照"当前预设"核对文档版式，报告偏差。

关键设计：检查标准来自用户选中的预设，而非死国标——用户改了预设，
检查标准自动跟着变，适用范围更广。每类检查可通过 options 开关启停。
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH

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


def _cm(v):
    return round(v, 2) if v is not None else None


def check_compliance(doc, preset, options=None, detect_types=None):
    """返回 findings 列表：[{'level': 'warn'/'info'/'ok', 'item': 名称, 'detail': 说明}]

    detect_types: 可选 {非空段序号: 类型} 或 None，用于结构完整性判断；
                  None 时内部用检测器识别。
    """
    opts = dict(DEFAULT_OPTIONS)
    if options:
        opts.update(options)
    findings = []

    def add(level, item, detail):
        findings.append({'level': level, 'item': item, 'detail': detail})

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
                                exp[0], exp[1], exp[2], exp[3], '、'.join(bad)))
        else:
            add('ok', '页边距', '符合预设')

    # 纸张大小
    if opts.get('paper'):
        w, h = _cm(sec.page_width.cm), _cm(sec.page_height.cm)
        want = preset.get('page_size', 'A4')
        is_a4 = abs((w or 0) - 21.0) < 0.2 and abs((h or 0) - 29.7) < 0.2
        if want == 'A4' and not is_a4:
            add('warn', '纸张', '当前 {}×{} cm 非 A4（预设要求 A4）'.format(w, h))
        else:
            add('ok', '纸张', '{}×{} cm'.format(w, h))

    # 正文字体字号（抽样首个正文段）
    if opts.get('body_font') or opts.get('line_spacing'):
        body = preset.get('body', {})
        from docx.oxml.ns import qn
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
                    add('warn', '正文字体', '实际「{}」，预设要求「{}」'.format(ea, exp_font))
                elif size and exp_size and abs(size - exp_size) > 0.3:
                    add('warn', '正文字号', '实际 {}pt，预设要求 {}pt'.format(size, exp_size))
                else:
                    add('ok', '正文字体字号', '符合预设')
            if opts.get('line_spacing'):
                exp_ls = body.get('line_spacing')
                ls = body_para.paragraph_format.line_spacing
                if exp_ls:
                    got_pt = ls.pt if hasattr(ls, 'pt') else None
                    if got_pt and abs(got_pt - exp_ls) > 1:
                        add('warn', '正文行距', '实际约 {}pt，预设要求固定 {}pt'.format(round(got_pt, 1), exp_ls))
                    else:
                        add('ok', '正文行距', '符合预设')

    # 结构完整性 + 标题居中
    if opts.get('structure') or opts.get('title_center'):
        from .detector import detect_para_type, _compile_rules, _build_text_context
        rules = _compile_rules(preset.get('detect_rules'))
        all_texts, idx_map = _build_text_context(doc)
        types = {}
        prev = None
        title_para = None
        for i, p in enumerate(doc.paragraphs):
            t = p.text.strip()
            if not t:
                continue
            ai = idx_map.get(i)
            ptype = detect_para_type(t, i, len(doc.paragraphs), p.paragraph_format.alignment,
                                     all_texts, all_texts_index=ai, prev_para_type=prev, rules=rules)
            types[ptype] = types.get(ptype, 0) + 1
            if ptype == 'title' and title_para is None:
                title_para = p
            prev = ptype
        if opts.get('structure'):
            missing = []
            if not types.get('title'):
                missing.append('标题')
            if not types.get('recipient'):
                missing.append('主送机关')
            if not types.get('date'):
                missing.append('成文日期')
            if missing:
                add('warn', '结构完整性', '未识别到：{}（如确有请在预览里核对/指定）'.format('、'.join(missing)))
            else:
                add('ok', '结构完整性', '标题/主送机关/成文日期齐全')
        if opts.get('title_center') and title_para is not None:
            al = title_para.paragraph_format.alignment
            if al != WD_ALIGN_PARAGRAPH.CENTER:
                add('warn', '标题居中', '标题当前非居中')
            else:
                add('ok', '标题居中', '标题居中')

    # 残留英文标点
    if opts.get('punctuation'):
        from . import analyzer
        issues = analyzer.analyze_punctuation(doc)
        if issues:
            add('warn', '标点规范', '发现 {} 处英文/不规范标点，建议先用「标点修复」'.format(len(issues)))
        else:
            add('ok', '标点规范', '未发现英文标点')

    return findings


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
