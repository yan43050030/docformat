# -*- coding: utf-8 -*-
"""拿一张空白套头扫描件，自动量出模板整体偏了多少。

原理很直白：套头纸上那几条红线是最好认的基准——横贯整幅、位置固定。
两边都用同一个探测器把线的位置找出来，一比就知道差多少：

    扫描件的线   ──┬──────────────┬──   实测在纸上的真实位置
    模板的线     ──┴─┐            ┴─┐   我们排出来的位置
                      差这一点，就是要补的偏移

模板那边的框线本来是白的（套打靠白线占位、不显影），量之前先把它临时刷黑，
渲染成 PDF 再探测。两边走同一套逻辑，探测器本身的系统误差正好抵消。

只解一个**整体平移**（dx, dy）和一个缩放比。扫描件的歪斜、纸张形变、
打印机自身的走纸误差都不在这个模型里——所以结果只当"建议值"，
真正定稿还得打一张出来看。

不引入 numpy：整页按 100dpi 渲染成灰度也就 80 万个字节，用
bytes.translate + count 做逐行统计，纯 Python 也够快（实测 <0.3 秒）。
"""
import logging
import os
import shutil
import tempfile

from docx.oxml.ns import qn

logger = logging.getLogger('docformat.scan_align')

PT_PER_CM = 28.3465
DPI = 100.0
# 灰度低于这个值算"有墨"。红线（约 204,34,34）的灰度约 85，稳稳落在里面
INK_MAX = 190
# 认线看的是"**连续**有墨的最长一段有多长"，不是这一行总共多少个墨点：
# 一行正文的墨点也能很多，但它是断的；线是连的。用连续长度一刀就切开了。
MIN_H_CM = 8.0     # 横线：套头上的长红线有 16.8cm，正文行连不出 8cm
MIN_V_CM = 1.5     # 竖线：表格里的分栏线只有一两厘米高，门槛得低
# 两边的线相距多远还算"同一条"
MATCH_TOL_CM = 1.2


def available():
    """(能不能用, 说明)"""
    try:
        import fitz  # noqa: F401
    except Exception as exc:
        return False, '缺少 PyMuPDF，无法读取扫描件（{}）'.format(
            exc.__class__.__name__)
    return True, 'PyMuPDF'


_RUN_RE = None


def _ink_profile(page, rotate):
    """把一页渲染成灰度，逐行求"连续有墨的最长一段"有多少像素。

    rotate=90 时整页转 90°，行也就变成了列——省得在 Python 里做转置。
    逐行用正则找最长连续段，是 C 里跑的；整页 100dpi 也就 80 万字节。
    """
    import re
    import fitz
    global _RUN_RE
    if _RUN_RE is None:
        _RUN_RE = re.compile(b'\x01+')
    zoom = DPI / 72.0
    mat = fitz.Matrix(zoom, zoom)
    if rotate:
        mat = mat * fitz.Matrix(0, 1, -1, 0, 0, 0)   # 顺时针 90°
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY, alpha=False)
    table = bytes(1 if i <= INK_MAX else 0 for i in range(256))
    data = pix.samples.translate(table)
    w, h = pix.width, pix.height
    runs = []
    for y in range(h):
        row = data[y * w:(y + 1) * w]
        best = 0
        for m in _RUN_RE.finditer(row):
            n = m.end() - m.start()
            if n > best:
                best = n
        runs.append(best)
    return runs, h, w


def _lines_from_profile(runs, need_px, px_per_cm):
    """把连续够长的那几行并成一条线，返回线中心的厘米坐标列表。"""
    out = []
    run_start = None
    for i, c in enumerate(runs):
        if c >= need_px:
            if run_start is None:
                run_start = i
        elif run_start is not None:
            out.append((run_start + i - 1) / 2.0 / px_per_cm)
            run_start = None
    if run_start is not None:
        out.append((run_start + len(runs) - 1) / 2.0 / px_per_cm)
    return out


def rules_of_pdf(path, page_index=0):
    """探测一页里的横线与竖线，返回 {'h': [cm], 'v': [cm], 'page_cm': (w, h)}"""
    import fitz
    doc = fitz.open(path)
    try:
        if page_index >= doc.page_count:
            raise ValueError('这份 PDF 只有 {} 页'.format(doc.page_count))
        page = doc[page_index]
        px_per_cm = DPI / 2.54
        runs, _h, _w = _ink_profile(page, rotate=False)
        rows = _lines_from_profile(runs, MIN_H_CM * px_per_cm, px_per_cm)
        runs_t, _ht, _wt = _ink_profile(page, rotate=True)
        cols = _lines_from_profile(runs_t, MIN_V_CM * px_per_cm, px_per_cm)
        return {'h': rows, 'v': cols,
                'page_cm': (page.rect.width / PT_PER_CM,
                            page.rect.height / PT_PER_CM)}
    finally:
        doc.close()


def _blacken_borders(doc):
    """把模板里的白色框线临时刷黑，好让探测器看得见。"""
    n = 0
    for el in doc.element.body.iter():
        tag = el.tag.split('}')[-1]
        if tag not in ('tblBorders', 'tcBorders'):
            continue
        for side in el:
            if side.get(qn('w:color')) in ('FFFFFF', 'ffffff'):
                side.set(qn('w:color'), '000000')
                if (side.get(qn('w:val')) or 'none') == 'none':
                    continue
                n += 1
    return n


def rules_of_template(template_path):
    """模板自己那几条线在哪儿（把白线刷黑后渲染再探测）。"""
    from docx import Document
    from .exporter import export_pdf
    tmp = tempfile.mkdtemp(prefix='align_')
    try:
        dx = os.path.join(tmp, 'tpl.docx')
        pdf = os.path.join(tmp, 'tpl.pdf')
        doc = Document(template_path)
        if not _blacken_borders(doc):
            raise ValueError('这个模板里没有框线，没法拿线来对位')
        doc.save(dx)
        ok, info = export_pdf(dx, pdf)
        if not ok:
            raise RuntimeError('模板转 PDF 失败：{}'.format(info))
        return rules_of_pdf(pdf)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _fit(template_lines, scan_lines):
    """把两组线配上对，解出平移量。

    只取"互为最近"的那些对，避免一条线被抢着配。差值取中位数：
    扫描件上多出来的线（装订孔、折痕）只会污染个别配对，中位数不受影响。
    """
    pairs = []
    for t in template_lines:
        if not scan_lines:
            break
        s = min(scan_lines, key=lambda x: abs(x - t))
        if abs(s - t) > MATCH_TOL_CM:
            continue
        # 反过来也得是最近的，否则两条模板线会挤到同一条扫描线上
        back = min(template_lines, key=lambda x: abs(x - s))
        if abs(back - t) > 1e-9:
            continue
        pairs.append((t, s))
    if not pairs:
        return None, [], 1.0
    diffs = sorted(s - t for t, s in pairs)
    shift = diffs[len(diffs) // 2] if len(diffs) % 2 else \
        (diffs[len(diffs) // 2 - 1] + diffs[len(diffs) // 2]) / 2.0
    scale = 1.0
    if len(pairs) >= 2:
        t_span = pairs[-1][0] - pairs[0][0]
        s_span = pairs[-1][1] - pairs[0][1]
        if abs(t_span) > 3.0:
            scale = s_span / t_span
    return shift, pairs, scale


def align(scan_path, template_path):
    """量出模板相对扫描件整体偏了多少。

    返回 dict：dx/dy（cm，正数=模板要往右/往下挪）、匹配到的线、缩放比、
    以及给人看的说明。量不出来时 dx/dy 为 None，reason 说明原因。
    """
    t = rules_of_template(template_path)
    s = rules_of_pdf(scan_path)
    out = {'template_rules': t, 'scan_rules': s, 'warnings': []}

    tw, th = t['page_cm']
    sw, sh = s['page_cm']
    if abs(tw - sw) > 0.5 or abs(th - sh) > 0.5:
        out['warnings'].append(
            '扫描件页面 {:.1f}×{:.1f}cm，模板 {:.1f}×{:.1f}cm，'
            '两者纸张大小不一致——多半是扫描时缩放了，量出来的偏移不可信。'
            .format(sw, sh, tw, th))

    dy, hp, ys = _fit(t['h'], s['h'])
    dx, vp, xs = _fit(t['v'], s['v'])
    out.update({'dx': dx, 'dy': dy, 'h_pairs': hp, 'v_pairs': vp,
                'scale_y': ys, 'scale_x': xs})
    if dy is None and dx is None:
        out['reason'] = ('在扫描件上没找到能与模板对上的线。'
                         '请确认扫描的是**空白套头纸**、按原尺寸（100%）扫描、'
                         '且四边没有被裁掉。')
    for name, sc in (('横向', xs), ('纵向', ys)):
        if abs(sc - 1.0) > 0.01:
            out['warnings'].append(
                '{}上扫描件比模板{}了 {:.1%}，说明扫描或打印有缩放；'
                '整体平移补不了缩放，请把扫描/打印都设成 100% 再来一次。'
                .format(name, '大' if sc > 1 else '小', abs(sc - 1.0)))
    return out


def describe(result):
    """把 align 的结果说成人话。"""
    if result.get('dx') is None and result.get('dy') is None:
        return result.get('reason', '量不出偏移')
    lines = []
    dx, dy = result.get('dx'), result.get('dy')
    if dx is not None:
        lines.append('横向：模板要往{}挪 {:.2f}cm（配上 {} 条竖线）'
                     .format('右' if dx > 0 else '左', abs(dx),
                             len(result.get('v_pairs') or [])))
    else:
        lines.append('横向：没找到能对上的竖线，保持不动')
    if dy is not None:
        lines.append('纵向：模板要往{}挪 {:.2f}cm（配上 {} 条横线）'
                     .format('下' if dy > 0 else '上', abs(dy),
                             len(result.get('h_pairs') or [])))
    else:
        lines.append('纵向：没找到能对上的横线，保持不动')
    lines += result.get('warnings') or []
    return '\n'.join(lines)
