# -*- coding: utf-8 -*-
"""纵向自动校准：build → 渲染 → 实测 → 回填修正量，迭代到收敛。

字面顶端与行盒顶端差多少，取决于字体、字号、渲染引擎，猜公式屡试屡错。
不如量：把每个元素的实测偏差直接加回它的段前距，重来一次，两三轮就贴合。
收敛结果写进 songshendan_calib.json，生成模板时自动带上。
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools.check_songshendan import check, TOL          # noqa: E402
from tools.make_songshendan import CALIB_PATH, SPEC, W, H, build   # noqa: E402

# 实测项 → 影响它的那个校准键（纵向的才需要）
OWNER = {
    '单位名称行 上边线距上': 'head1',
    '文件送审单 上边线距上': 'head2',
    '标题 上边线距上': 'title',
    '领导批示 上边线距上': 'lead',
    '拟办意见 上边线距下': 'opinion',
    '承办部门 上边线距下': 'dept',
    '经办人 上边线距下': 'handler',
    '落款 下边线距下': 'sign',
}
# 横向已经不需要校准了：栏目名全是全角字，宽度 = 字数 × 字号，
# 由此算出的左制表位是绝对位置，与字体无关（原来的
# handler_x / check_x / phone_x / head1_gap 四个修正量随之作废）。
OWNER_X = {}
FROM_RIGHT = set()
# 从纸下边量的项，误差方向与从上边量的相反
FROM_BOTTOM = {'拟办意见 上边线距下', '承办部门 上边线距下',
               '经办人 上边线距下', '落款 下边线距下'}


def measure(tpl):
    from scripts import overprint as op
    from scripts.exporter import export_pdf
    tmp = tempfile.mkdtemp()
    dx = os.path.join(tmp, 'f.docx')
    pdf = os.path.join(tmp, 'f.pdf')
    op.fill_form(tpl, {}, dx, one_page=False)
    ok, info = export_pdf(dx, pdf)
    if not ok:
        raise RuntimeError('转 PDF 失败：%s' % info)
    return check(pdf, SPEC, W, H)


def main(rounds=6):
    calib = {}
    tmp = os.path.join(tempfile.mkdtemp(), 'cal.docx')
    for it in range(rounds):
        build(tmp, calib=calib)
        rows = measure(tmp)
        errs = {n: (g - w) for n, w, g, _ok in rows if g is not None}
        worst = max((abs(v) for n, v in errs.items()
                     if n in OWNER or n in OWNER_X), default=0)
        print('第 %d 轮：纵向最大偏差 %.3f cm' % (it + 1, worst))
        if worst <= 0.03:
            break
        for name, key in OWNER.items():
            if name not in errs:
                continue
            d = errs[name]
            # 从下边量的：实测值大 = 元素偏高 = 该往下挪
            # 钳在 ±2cm：修正量本该是毫米级，跑出这个范围说明该元素被行高
            # 裁掉、测到的已不是它本身，继续迭代只会越跑越偏
            nv = calib.get(key, 0.0) + (d if name in FROM_BOTTOM else -d)
            calib[key] = round(max(-2.0, min(2.0, nv)), 3)
        for name, key in OWNER_X.items():
            if name not in errs:
                continue
            nv = calib.get(key, 0.0) + errs[name]   # 距右沿：实测大就往右挪
            calib[key] = round(max(-3.0, min(3.0, nv)), 3)
    with open(CALIB_PATH, 'w', encoding='utf-8') as f:
        json.dump(calib, f, ensure_ascii=False, indent=2, sort_keys=True)
    print('校准量已写入', CALIB_PATH)
    print(json.dumps(calib, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == '__main__':
    sys.exit(main())
