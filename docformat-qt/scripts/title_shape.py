# -*- coding: utf-8 -*-
"""标题梯形回行：长标题分行排成正梯形（上长下短）或倒梯形（上短下长）。

GB/T 9704：标题回行时词意完整、排列对称、长短适宜。中文无词库无法 100%
分词，这里用启发式尽力而为：按版心宽度估算每行字数，在"安全断点"处折行，
避免拆坏词、避免单字甩行。用户可在预览里手动微调。
"""
import re

# 不宜在其后立即断行的字（断了会把词拆开）
_NO_BREAK_AFTER = set('的和与及为对在向从把被将之其此该等或')
# 不宜出现在行首的字（介词/助词等）
_NO_BREAK_BEFORE = set('的了着过之')


def _char_width(ch):
    """全角/汉字算 1，半角字母数字算 0.5"""
    return 0.5 if ord(ch) < 128 else 1.0


def _text_width(s):
    return sum(_char_width(c) for c in s)


def _find_break(text, target, hard_max):
    """在 target 附近找一个安全断点，返回断点下标（该下标前成一行）。"""
    n = len(text)
    if target >= n:
        return n
    # 在 [target-3, target+3] 窗口里找不拆词的位置，优先接近 target
    candidates = []
    lo = max(1, int(target) - 3)
    hi = min(n - 1, int(target) + 3)
    for pos in range(lo, hi + 1):
        prev = text[pos - 1]
        cur = text[pos]
        if prev in _NO_BREAK_AFTER or cur in _NO_BREAK_BEFORE:
            continue
        candidates.append(pos)
    if candidates:
        return min(candidates, key=lambda p: abs(p - target))
    # 找不到安全点就在 target 硬断（不超过 hard_max）
    return min(int(target) if target >= 1 else 1, hard_max)


def split_title_lines(text, chars_per_line, shape='trapezoid_down'):
    """把标题文本拆成多行，形成梯形。返回行列表（≥1 行）。

    shape: 'trapezoid_down' 正梯形(上长下短) / 'trapezoid_up' 倒梯形(上短下长)
    chars_per_line: 版心可容纳的全角字符数（每行上限）
    """
    text = text.strip()
    if not text:
        return [text]
    total = _text_width(text)
    # 能一行放下就不折
    if total <= chars_per_line:
        return [text]

    # 目标行数：尽量 2 行，过长才 3 行
    lines_n = 2 if total <= chars_per_line * 2 - 2 else 3
    lines_n = min(lines_n, 3)

    # 目标各行宽度：正梯形递减、倒梯形递增，和为 total
    base = total / lines_n
    if shape == 'trapezoid_up':
        weights = [1.0 - 0.18 * (lines_n - 1 - i) for i in range(lines_n)]  # 递增
    else:
        weights = [1.0 - 0.18 * i for i in range(lines_n)]                   # 递减
    ws = sum(weights)
    targets = [total * w / ws for w in weights]

    lines = []
    rest = text
    for i in range(lines_n - 1):
        if not rest:
            break
        # 目标宽度换算成字符下标（近似：按累计宽度找）
        tw = targets[i]
        acc = 0.0
        cut = len(rest)
        for j, ch in enumerate(rest):
            acc += _char_width(ch)
            if acc >= tw:
                cut = j + 1
                break
        cut = _find_break(rest, cut, len(rest))
        cut = max(1, min(cut, len(rest) - 1))
        lines.append(rest[:cut])
        rest = rest[cut:]
    if rest:
        lines.append(rest)
    return [ln for ln in lines if ln]


def apply_title_shape(paragraph, chars_per_line, shape='trapezoid_down'):
    """在标题段落内插入换行符，形成梯形。返回是否改动。"""
    if shape not in ('trapezoid_down', 'trapezoid_up'):
        return False
    text = paragraph.text.strip()
    lines = split_title_lines(text, chars_per_line, shape)
    if len(lines) <= 1:
        return False

    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    # 保留首个 run 的字体属性，重建为多行（用 <w:br/> 分隔）
    runs = paragraph.runs
    if not runs:
        return False
    template_rpr = runs[0]._r.find(qn('w:rPr'))
    # 清空现有 runs 文本
    for r in list(runs):
        paragraph._p.remove(r._r)

    for i, line in enumerate(lines):
        run = paragraph.add_run(line)
        if template_rpr is not None:
            # 复制字体属性到新 run
            import copy
            new_rpr = copy.deepcopy(template_rpr)
            old = run._r.find(qn('w:rPr'))
            if old is not None:
                run._r.remove(old)
            run._r.insert(0, new_rpr)
        if i < len(lines) - 1:
            br = OxmlElement('w:br')
            run._r.append(br)
    return True
