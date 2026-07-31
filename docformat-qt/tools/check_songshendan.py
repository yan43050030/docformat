# -*- coding: utf-8 -*-
"""把送审单模板渲染成 PDF，逐项实测与 SPEC 对照。

模板是不是"按尺寸做对了"，不能靠看，要量。本脚本填一份样例、转 PDF、
用 PyMuPDF 读出每段文字的真实坐标，和用户尺子量的数值逐条比。
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

PT2CM = 2.54 / 72.0
TOL = 0.15          # 允许误差（cm）：尺子读数本身就有 ~1mm 量级

# 口径存疑的几项：单独列出来，不计入"超差"，但数字照printed，等在纸上复核。
# 它们对不上不是模板算错了，是"量的到底是哪条线"这件事本身没定死——
# 硬把模板迁就过去，反而会把已经对准的地方弄歪。
UNSETTLED = {
    '文件送审单 下边线距上':
        '量的是字面底边，随字体而变；本机没装方正大标宋，渲染用的是替代字体',
}


def spans(pdf_path, page=0):
    import fitz
    out = []
    for b in fitz.open(pdf_path)[page].get_text('dict')['blocks']:
        for l in b.get('lines', []):
            for s in l['spans']:
                t = s['text'].strip()
                if t:
                    x0, y0, x1, y1 = s['bbox']
                    out.append({'x0': x0 * PT2CM, 'x1': x1 * PT2CM,
                                'y0': y0 * PT2CM, 'y1': y1 * PT2CM, 't': t})
    return out


def find(sps, needle, nth=0):
    hits = [s for s in sps if needle in s['t']]
    return hits[nth] if len(hits) > nth else None


def find_line(sps, chars, ytol=0.3):
    """按"同一行上的这些字"合并成一个整体范围。

    字距拉得大时，渲染器会把一串字拆成多个片段（「文 件 送 审 单」就是
    五段），用 in 匹配整串会一个都找不到。这里按首字定位那一行，
    再把该行的片段并起来量左右上下沿。
    """
    seed = [x for x in sps if chars[0] in x['t']]
    if not seed:
        return None
    best = None
    for sd in seed:
        row = [x for x in sps if abs(x['y0'] - sd['y0']) <= ytol]
        if all(any(c in x['t'] for x in row) for c in chars):
            cand = {'x0': min(x['x0'] for x in row),
                    'x1': max(x['x1'] for x in row),
                    'y0': min(x['y0'] for x in row),
                    'y1': max(x['y1'] for x in row),
                    't': ''.join(x['t'] for x in row)}
            if best is None or cand['y0'] < best['y0']:
                best = cand
    return best


def check(pdf_path, spec, W=21.0, H=29.6):
    sps = spans(pdf_path)
    rows = []

    def cmp(name, got, want, unit='cm'):
        ok = got is not None and abs(got - want) <= TOL
        rows.append((name, want, got, ok))

    s = find(sps, '中国某地市某单位')
    cmp('单位名称行 上边线距上', s['y0'] if s else None, spec['head1_top'])
    cmp('单位名称行 左边线距左', s['x0'] if s else None, spec['head1_left'])
    s6 = find(sps, '某地市某单位', 1) or s
    cmp('单位名称行 右边线距右', (W - s6['x1']) if s6 else None, spec['head1_right'])

    s = find_line(sps, '文件送审单') or find(sps, '文件送审单')
    cmp('文件送审单 上边线距上', s['y0'] if s else None, spec['head2_top'])
    cmp('文件送审单 下边线距上', s['y1'] if s else None, spec['head2_bottom'])
    cmp('文件送审单 左边线距左', s['x0'] if s else None, spec['head2_left'])
    cmp('文件送审单 右边线距右', (W - s['x1']) if s else None, spec['head2_right'])

    s = find(sps, '紧急程度')
    cmp('紧急程度：左边线距左', s['x0'] if s else None, spec['urgent_left'])
    s = find(sps, '密级')
    cmp('密级：右边线距右', (W - s['x1']) if s else None, spec['sec_right'])

    s = find(sps, '标  题') or find(sps, '标')
    cmp('标题 上边线距上', s['y0'] if s else None, spec['title_text_top'])
    cmp('标题 左边线距左', s['x0'] if s else None, spec['title_left'])

    s = find(sps, '领导批示')
    cmp('领导批示 上边线距上', s['y0'] if s else None, spec['lead_text_top'])
    cmp('领导批示 左边线距左', s['x0'] if s else None, spec['lead_left'])
    cmp('领导批示 右边线距左', s['x1'] if s else None, spec['lead_right'])

    s = find(sps, '拟办意见')
    cmp('拟办意见 上边线距下', (H - s['y0']) if s else None,
        H - spec['opinion_text_top'])
    cmp('拟办意见 左边线距左', s['x0'] if s else None, spec['lead_left'])

    s = find(sps, '承办部门')
    cmp('承办部门 上边线距下', (H - s['y0']) if s else None,
        H - spec['dept_text_top'])
    cmp('承办部门 左边线距左', s['x0'] if s else None, spec['lead_left'])

    s = find(sps, '经办人') or find(sps, '经 办 人')
    cmp('经办人 上边线距下', (H - s['y0']) if s else None,
        H - spec['handler_text_top'])
    cmp('经办人 右边线距右', (W - s['x1']) if s else None, spec['handler_right'])
    ph = find(sps, '电话')
    if s and ph:
        cmp('经办人→电话 间距', ph['x0'] - s['x1'], spec['handler_to_phone'])
    cmp('电话 右边线距右', (W - ph['x1']) if ph else None, spec['phone_right'])

    s = find(sps, '文字校核')
    cmp('文字校核 上边线距下', (H - s['y0']) if s else None,
        H - spec['check_text_top'])
    cmp('文字校核 右边线距右', (W - s['x1']) if s else None, spec['handler_right'])

    for ch, key in (('年', 'ymd_year_left'), ('月', 'ymd_month_left'),
                    ('日', 'ymd_day_left')):
        hit = [x for x in sps if x['t'] == ch]
        cmp('{} 左边线距左'.format(ch), hit[-1]['x0'] if hit else None, spec[key])

    # 落款要按整串找：它的前几个字与文件头单位名重复，截取匹配会抓错行
    cands = [x for x in sps if spec['sign_text'][:6] in x['t'] and x['y0'] > 20]
    s = cands[-1] if cands else None
    cmp('落款 首字左边线距右', (W - s['x0']) if s else None,
        spec['sign_left_from_right'])
    e = s
    cmp('落款 尾字右边线距右', (W - e['x1']) if e else None,
        spec['sign_right_from_right'])
    cmp('落款 下边线距下', (H - e['y1']) if e else None,
        spec['sign_bottom_from_bottom'])
    return rows


def main(tpl):
    from scripts import overprint as op
    from scripts.exporter import export_pdf
    from tools.make_songshendan import SPEC, W, H
    tmp = tempfile.mkdtemp()
    dx = os.path.join(tmp, 'f.docx')
    pdf = os.path.join(tmp, 'f.pdf')
    # 白字在 PDF 里照样有坐标，直接填空值即可量位置
    op.fill_form(tpl, {}, dx, one_page=False)
    ok, info = export_pdf(dx, pdf)
    if not ok:
        print('转 PDF 失败：', info)
        return 1
    rows = check(pdf, SPEC, W, H)
    bad = 0
    print('%-24s %8s %8s %8s' % ('项目', '实测(尺)', 'PDF', '差'))
    print('-' * 54)
    for name, want, got, good in rows:
        note = name in UNSETTLED
        if not good and not note:
            bad += 1
        print('%-24s %8.2f %8s %8s  %s'
              % (name, want,
                 '%.2f' % got if got is not None else '  --',
                 '%+.2f' % (got - want) if got is not None else '  --',
                 '' if good else ('※' if note else '✗')))
    print('-' * 54)
    print('%d 项，%d 项超出 ±%.2fcm' % (len(rows), bad, TOL))
    flagged = [(n, w, g) for n, w, g, ok in rows if not ok and n in UNSETTLED]
    if flagged:
        print('\n※ 以下几项口径存疑，请拿真纸复核后再定（不计入超差）：')
        for n, w, g in flagged:
            print('  · %s：尺子 %.2f，模板 %.2f\n    %s' % (n, w, g, UNSETTLED[n]))
    return 0 if bad == 0 else 2


if __name__ == '__main__':
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else
                  os.path.join(os.path.dirname(os.path.dirname(
                      os.path.abspath(__file__))),
                      'templates', '套打', '文件送审单.docx')))
