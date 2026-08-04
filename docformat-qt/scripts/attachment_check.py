# -*- coding: utf-8 -*-
"""正文说的附件，和实际附上的附件，对不对得上。

收文时最常见的一类差错：正文写着"附件：1.实施方案 2.任务分工表"，翻到后面
只有一个附件；或者附件标识写着"附件2"，正文却只提了一个。这类错人眼容易
放过（要来回翻页数数），机器判得准。

两侧各是什么
------------
  · **附件说明**：正文末尾那几行「附件：1.xxx 2.yyy」，告诉收文人有什么；
  · **附件标识**：附件正文首页左上角顶格的「附件1」，标明这一页是第几个附件。

只有两侧都存在时才谈得上"对不上"。附件常常是单独的文件，正文有说明、
本文档里没有标识是完全正常的——那种情况只作提示，不算差错，
否则十份文件有九份要被冤枉。
"""
import logging
import re

logger = logging.getLogger('docformat.attachment')

# 「附件：」「附件1：」「附件一」——说明行的起头
_DESC_HEAD = re.compile(r'^附件\s*([0-9０-９一二三四五六七八九十]*)\s*[：:．.]?\s*')
# 说明行里的条目编号：1. / 1、 / （1）
_ITEM = re.compile(r'(?:^|\s)([0-9０-９]{1,2})\s*[.、．]\s*')
# 附件标识：独占一行的「附件」「附件1」「附件一」
_LABEL = re.compile(r'^附件\s*([0-9０-９一二三四五六七八九十]*)\s*$')

_CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
           '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}

GROUPS = [
    ('att_count', '附件个数（正文说的与实际附上的对不对得上）'),
    ('att_number', '附件序号（是否连续、有无重复；只有一个附件时不编号）'),
]
GROUP_KEYS = [k for k, _ in GROUPS]


def _num(text):
    """把「1」「１」「一」都读成整数；读不出返回 None"""
    t = (text or '').strip()
    if not t:
        return None
    t = t.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    if t.isdigit():
        return int(t)
    if t == '十':
        return 10
    if len(t) == 1:
        return _CN_NUM.get(t)
    if t.startswith('十'):
        return 10 + _CN_NUM.get(t[1:], 0)
    return None


def parse_desc(texts):
    """从附件说明的若干行里读出条目：[(序号 or None, 名称)]。

    说明的写法有好几种，都要认：
        附件：1.实施方案  2.任务分工表        （一行里带编号）
        附件：                                 （编号在后续各行）
          1.实施方案
          2.任务分工表
        附件：实施方案                         （只有一个，不编号）
        附件1：实施方案                        （每个附件各占一行）
    """
    items = []
    for raw in texts:
        line = (raw or '').strip()
        if not line:
            continue
        head = _DESC_HEAD.match(line)
        head_no = None
        if head:
            head_no = _num(head.group(1))
            line = line[head.end():].strip()
            if not line:
                continue
        marks = list(_ITEM.finditer(line))
        if marks:
            for i, m in enumerate(marks):
                end = marks[i + 1].start() if i + 1 < len(marks) else len(line)
                name = line[m.end():end].strip(' 　;；,，')
                items.append((_num(m.group(1)), name))
        else:
            # 没有编号：整行就是一个附件名（「附件1：实施方案」的序号在头上）
            items.append((head_no, line.strip(' 　;；,，')))
    return items


def parse_labels(texts):
    """附件标识行 → [序号 or None]"""
    out = []
    for raw in texts:
        m = _LABEL.match((raw or '').strip())
        if m:
            out.append(_num(m.group(1)))
    return out


def check(desc_texts, label_texts, groups=None):
    """两侧对照，返回 findings（字段与 compliance 一致）。

    都不给 fix_key——附件对不上是内容问题，只有人知道到底该有几个附件，
    机器凭空补不出来，也不该替人删。
    """
    on = {k: True for k in GROUP_KEYS}
    on.update(groups or {})
    items = parse_desc(desc_texts)
    labels = parse_labels(label_texts)
    out = []

    if not items and not labels:
        return out

    if on['att_count']:
        if items and labels and len(items) != len(labels):
            out.append({
                'level': 'warn', 'item': '附件·个数对不上',
                'detail': '正文说有 {} 个附件（{}），实际附上了 {} 个附件标识。'
                          '请核对是漏附了，还是说明写多了'
                          .format(len(items),
                                  '、'.join(n or '未命名' for _i, n in items[:4]),
                                  len(labels)),
            })
        elif labels and not items:
            out.append({
                'level': 'warn', 'item': '附件·正文未说明',
                'detail': '文中有 {} 个附件标识，但正文末尾没有"附件："说明行。'
                          '收文人不知道该有哪些附件'.format(len(labels)),
            })
        elif items and not labels:
            # 附件多半是单独的文件，这里只作提示，不算差错
            out.append({
                'level': 'info', 'item': '附件·说明',
                'detail': '正文说有 {} 个附件，本文档内没有附件正文'
                          '（附件若是单独的文件，属正常）'.format(len(items)),
            })
        else:
            out.append({'level': 'ok', 'item': '附件·个数',
                        'detail': '正文说明与附件标识都是 {} 个'.format(len(items))})

    if not on['att_number']:
        return out

    for name, seq in (('正文的附件说明', [n for n, _t in items]),
                      ('附件标识', labels)):
        nums = [n for n in seq if n is not None]
        if len(nums) < 2:
            continue
        if len(set(nums)) != len(nums):
            dup = sorted({n for n in nums if nums.count(n) > 1})
            out.append({'level': 'warn', 'item': '附件·序号重复',
                        'detail': '{}里序号 {} 出现了不止一次'.format(
                            name, '、'.join(str(x) for x in dup))})
        elif sorted(nums) != list(range(1, len(nums) + 1)):
            out.append({'level': 'warn', 'item': '附件·序号不连续',
                        'detail': '{}的序号是 {}，应当从 1 起连续编号'.format(
                            name, '、'.join(str(x) for x in nums))})

    # GB/T 9704：附件只有一个时不编号；两个以上才用阿拉伯数字标顺序号。
    # 个数本身已经对不上时不再提这一条——那时候连"到底有几个附件"都还没定，
    # 再追着编号说事只是往报告里添噪音
    mismatch = bool(items and labels and len(items) != len(labels))
    if not mismatch:
        sides = []
        if len(items) == 1 and items[0][0] is not None:
            sides.append('正文说明里写作「{}.」'.format(items[0][0]))
        if len(labels) == 1 and labels[0] is not None:
            sides.append('附件标识写作「附件{}」'.format(labels[0]))
        if sides:
            out.append({'level': 'warn', 'item': '附件·单个不编号',
                        'detail': '只有一个附件却编了序号（{}）。GB/T 9704 规定'
                                  '附件只有一个时不编号'.format('，'.join(sides))})
    return out


def check_doc(typed, groups=None):
    """从已识别好类型的段落里取两侧文字，再对照。

    typed 是 [(段序, paragraph, 类型)]——与 compliance 的段落识别共用一次
    结果，不重复跑一遍识别。
    """
    desc, labels = [], []
    for _i, para, ptype in typed:
        t = para.text.strip()
        if not t:
            continue
        if ptype == 'attachment':
            desc.append(t)
        elif ptype == 'attachment_label':
            labels.append(t)
    return check(desc, labels, groups=groups)
