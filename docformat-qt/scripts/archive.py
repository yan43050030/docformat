# -*- coding: utf-8 -*-
"""归档命名与登记表：把处理好的公文按规矩改名、并记一行台账。

排版完之后还有一段没人愿意干的活：按"〔年〕号-文种-事由"把文件改名，再把
文号、标题、成文日期、密级一条条敲进登记台账。这些字软件其实全都认得——
排版时就识别过标题、发文字号、成文日期，密级检查时又读过密级和份号。
让人再抄一遍，纯属浪费。

两条底线
--------
· **不动原文件**：默认复制到归档目录，原件留在原地。要移动得明说。
· **不覆盖**：目标重名就自动加 (2)(3)，宁可多一份也不能悄悄盖掉别人的。

台账里存的是标题、文号这些**明文**——那正是台账的用处，但也意味着它自己
就是一份需要按密级管理的东西。存到哪儿由用户自己定，软件不替他挑地方。
"""
import csv
import io
import logging
import os
import re
import shutil

logger = logging.getLogger('docformat.archive')

# 命名式里可以用的占位符 → (说明, 取不到时的替代)
FIELDS = [
    ('文号', '发文字号，如 某发〔2026〕12号', ''),
    ('标题', '公文标题', '无标题'),
    ('文种', '通知/请示/报告…', ''),
    ('成文日期', '20260717 形式', ''),
    ('年', '成文年份，如 2026', ''),
    ('月', '成文月份，两位', ''),
    ('日', '成文日，两位', ''),
    ('密级', '秘密★1年 / 空', ''),
    ('密级词', '秘密 / 机密 / 绝密 / 空', ''),
    ('份号', '涉密件的份号', ''),
    ('发文机关', '落款署名', ''),
    ('主送机关', '第一个主送机关', ''),
    ('原文件名', '不含扩展名', ''),
]
FIELD_KEYS = [k for k, _d, _f in FIELDS]

DEFAULT_PATTERN = '{成文日期}-{文号}-{标题}'

# 台账的列。顺序即列序，改了会跟已有台账对不上，别轻易动
LEDGER_COLUMNS = ['归档时间', '文号', '标题', '文种', '成文日期', '密级',
                  '份号', '发文机关', '主送机关', '归档文件名', '原文件名']

# Windows/Linux 文件名里不能出现的字符
_BAD_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')
# 文件名长度上限：路径整体还有目录要占，留足余量
MAX_STEM = 120


def safe_stem(text, fallback='未命名'):
    """把一段文字弄成能当文件名的样子"""
    s = _BAD_CHARS.sub('', (text or '')).strip(' .　')
    s = re.sub(r'\s+', ' ', s)
    # 连续的分隔符：字段取不到时留下的空档，收干净
    s = re.sub(r'[-_—]{2,}', '-', s).strip('-_ ')
    return (s or fallback)[:MAX_STEM]


def extract_meta(path, preset=None):
    """从一份 docx 里读出归档要用的那些字段。

    全部复用已有的识别能力：段落类型识别给标题/发文字号/成文日期/署名/
    主送机关，密级检查给密级和份号，用语模块给文种。这里只做汇总，
    不另起一套识别——两套识别迟早会各说各话。
    """
    from docx import Document
    from .paragraph import sanitize_document
    doc = Document(path)
    sanitize_document(doc)

    meta = {k: '' for k in FIELD_KEYS}
    meta['原文件名'] = os.path.splitext(os.path.basename(path))[0]

    try:
        from .compliance import _detect_types
        from .data_model import PRESETS
        typed = _detect_types(doc, preset or PRESETS['official_gbk'])
    except Exception as exc:
        logger.warning('归档：段落识别失败（%s）', exc)
        typed = []

    title_parts = []
    for _i, para, ptype in typed:
        t = para.text.strip()
        if not t:
            continue
        if ptype == 'title':
            title_parts.append(t)
        elif ptype == 'docnum' and not meta['文号']:
            meta['文号'] = t
        elif ptype == 'date' and not meta['成文日期']:
            meta.update(_split_date(t))
        elif ptype == 'signature' and not meta['发文机关']:
            meta['发文机关'] = t
        elif ptype == 'recipient' and not meta['主送机关']:
            meta['主送机关'] = t.rstrip('：: ').split('、')[0]
    meta['标题'] = ''.join(title_parts)

    try:
        from .security_mark import scan as _sec_scan
        f = _sec_scan(doc)
        if f['sec']:
            meta['密级'] = f['sec'][2]
            m = re.match(r'^(绝密|机密|秘密)', f['sec'][2])
            meta['密级词'] = m.group(1) if m else ''
        if f['copynum']:
            meta['份号'] = f['copynum'][2]
    except Exception as exc:
        logger.warning('归档：密级读取失败（%s）', exc)

    try:
        from .wording import doc_kind
        meta['文种'] = doc_kind([meta['标题']] or [])
    except Exception as exc:
        logger.warning('归档：文种识别失败（%s）', exc)

    return meta


def _split_date(text):
    """成文日期 → {成文日期: 20260717, 年/月/日}"""
    out = {'成文日期': '', '年': '', '月': '', '日': ''}
    try:
        from .overprint import parse_date
        y, mth, d = parse_date(text)
    except Exception:
        y = mth = d = ''
    if not y:
        m = re.search(r'(\d{4})\D{1,3}(\d{1,2})\D{1,3}(\d{1,2})', text or '')
        if m:
            y, mth, d = m.group(1), m.group(2), m.group(3)
    if not y:
        return out
    out['年'] = str(y)
    out['月'] = str(mth).zfill(2) if mth else ''
    out['日'] = str(d).zfill(2) if d else ''
    out['成文日期'] = '{}{}{}'.format(out['年'], out['月'], out['日'])
    return out


def render_name(pattern, meta, ext='.docx'):
    """按命名式生成文件名（含扩展名）。

    取不到的字段留空，随后把因此产生的连续分隔符收干净——
    「20260717--关于…」这种带着空档的名字，比不改名还难看。
    """
    def sub(m):
        return str(meta.get(m.group(1), '') or '')
    name = re.sub(r'\{([^{}]+)\}', sub, pattern or DEFAULT_PATTERN)
    return safe_stem(name, fallback=meta.get('原文件名') or '未命名') + ext


def unique_path(directory, filename):
    """重名就加 (2)(3)——宁可多一份，也不能悄悄盖掉别人的文件"""
    base, ext = os.path.splitext(filename)
    cand = os.path.join(directory, filename)
    n = 2
    while os.path.exists(cand):
        cand = os.path.join(directory, '{}({}){}'.format(base, n, ext))
        n += 1
    return cand


def plan(paths, pattern=DEFAULT_PATTERN, preset=None):
    """先算不动手：返回 [{'src','meta','new_name'}]，供界面预览确认。"""
    out = []
    for p in paths:
        try:
            meta = extract_meta(p, preset)
        except Exception as exc:
            logger.warning('归档：读取失败（%s）', exc)
            meta = {k: '' for k in FIELD_KEYS}
            meta['原文件名'] = os.path.splitext(os.path.basename(p))[0]
        ext = os.path.splitext(p)[1] or '.docx'
        out.append({'src': p, 'meta': meta,
                    'new_name': render_name(pattern, meta, ext)})
    return out


def append_ledger(ledger_path, rows):
    """把若干行追加到台账 CSV；文件不存在就先写表头。

    用 utf-8-sig 存：不带 BOM 的 UTF-8，Excel 双击打开会把中文显示成乱码，
    而台账正是要给人双击看的。
    """
    exists = os.path.exists(ledger_path) and os.path.getsize(ledger_path) > 0
    d = os.path.dirname(ledger_path)
    if d and not os.path.isdir(d):
        os.makedirs(d)
    with io.open(ledger_path, 'a', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(LEDGER_COLUMNS)
        for row in rows:
            w.writerow([row.get(c, '') for c in LEDGER_COLUMNS])
    return len(rows)


def archive(items, out_dir, move=False, ledger_path=None):
    """执行归档：复制（或移动）到 out_dir，并按需追加台账。

    items 来自 plan()，界面上改过 new_name 也照改后的来。
    返回 (成功的 [(源, 归档后路径)], 失败的 [(源, 原因)])。
    """
    import datetime
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    done, failed, rows = [], [], []
    stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    for it in items:
        src = it['src']
        dst = unique_path(out_dir, it['new_name'])
        try:
            if move:
                shutil.move(src, dst)
            else:
                shutil.copy2(src, dst)
        except Exception as exc:
            failed.append((src, str(exc)))
            continue
        done.append((src, dst))
        row = dict(it['meta'])
        row['归档时间'] = stamp
        row['归档文件名'] = os.path.basename(dst)
        rows.append(row)
    if ledger_path and rows:
        try:
            append_ledger(ledger_path, rows)
        except Exception as exc:
            failed.append((ledger_path, '台账写入失败：{}'.format(exc)))
    return done, failed
