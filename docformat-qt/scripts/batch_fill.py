# -*- coding: utf-8 -*-
"""批量套打：从表格读一批数据，一次生成一叠套打件。

表头就是字段名，一行一份。字段名对不上的列直接忽略——用户的表里往往还有
序号、备注这类跟模板无关的列，为几列多余的东西报错太不体贴。

xlsx 自己解，不引第三方库：xlsx 就是个 zip，共享字符串在
xl/sharedStrings.xml、单元格在 xl/worksheets/sheet1.xml，取值只需要这两个。
软件要在离线信创机器上跑，能少一个依赖是一个。csv 则要过一遍编码探测——
国内的表格多半是 GBK 存的，按 UTF-8 硬读会满屏乱码。
"""
import csv
import io
import logging
import os
import re

logger = logging.getLogger('docformat.batch')

# 文件名里不能出现的字符（Windows 最严，按它来）
_BAD_NAME = re.compile(r'[\\/:*?"<>|\r\n\t]')


def _decode(raw):
    """猜编码。国内表格多是 GBK，先按 UTF-8 试，不行再退 GBK。"""
    for enc in ('utf-8-sig', 'utf-8', 'gbk', 'gb18030'):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', 'replace'), 'utf-8(替换)'


def _read_csv(path):
    with open(path, 'rb') as f:
        text, enc = _decode(f.read())
    logger.info('读表 %s 编码=%s', os.path.basename(path), enc)
    delim = '\t' if path.lower().endswith(('.tsv', '.txt')) else ','
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    return [r for r in rows if any(str(c).strip() for c in r)]


def _col_index(ref):
    """A1 → 0，AB12 → 27。xlsx 里空单元格会整个跳过，得靠列标补位。"""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _read_xlsx(path):
    import zipfile
    from xml.etree import ElementTree as ET
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(ns + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(ns + 't')))
        names = [n for n in z.namelist()
                 if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
        if not names:
            raise ValueError('这个 xlsx 里没有工作表')
        root = ET.fromstring(z.read(sorted(names)[0]))
    rows = []
    for tr in root.iter(ns + 'row'):
        cells = {}
        for tc in tr.findall(ns + 'c'):
            ref = tc.get('r') or ''
            idx = _col_index(ref) if ref else len(cells)
            v = tc.find(ns + 'v')
            if tc.get('t') == 's' and v is not None:
                try:
                    text = shared[int(v.text)]
                except (ValueError, IndexError):
                    text = ''
            elif tc.get('t') == 'inlineStr':
                text = ''.join(t.text or '' for t in tc.iter(ns + 't'))
            else:
                text = (v.text or '') if v is not None else ''
            cells[idx] = text
        if not cells:
            continue
        width = max(cells) + 1
        row = [cells.get(i, '') for i in range(width)]
        if any(str(c).strip() for c in row):
            rows.append(row)
    return rows


def read_table(path):
    """读表，返回 (表头列表, [每行的 {字段: 值}])。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xlsx':
        rows = _read_xlsx(path)
    elif ext in ('.csv', '.tsv', '.txt'):
        rows = _read_csv(path)
    else:
        raise ValueError('只认 .xlsx / .csv / .tsv；'
                         '别的格式请在 Excel/WPS 里另存为 CSV')
    if not rows:
        raise ValueError('这张表是空的')
    header = [str(c).strip() for c in rows[0]]
    if not any(header):
        raise ValueError('第一行应该是字段名（表头），现在是空的')
    out = []
    for r in rows[1:]:
        item = {}
        for i, name in enumerate(header):
            if name:
                item[name] = str(r[i]).strip() if i < len(r) else ''
        out.append(item)
    return header, out


def safe_name(text, fallback='未命名'):
    """把一段文字修成能当文件名的样子"""
    text = _BAD_NAME.sub('_', str(text or '')).strip(' .')
    return (text or fallback)[:60]


def plan_batch(template_path, header):
    """对一下表头和模板字段，返回 (能用的字段, 表里多余的列, 模板里没填上的字段)。

    多余的列不算错——用户的表里常有序号、备注这些跟模板无关的东西。
    模板里没对上的字段会留空打印，也不算错，但要让人看见。
    """
    from . import overprint
    fields = overprint.scan_fields(template_path)
    matched = [h for h in header if h in fields]
    extra = [h for h in header if h and h not in fields]
    missing = [f for f in fields if f not in header]
    return matched, extra, missing


def batch_fill(template_path, rows, out_dir, name_field=None, prefix='',
               progress=None, **fill_kw):
    """逐行生成，返回 [(输出路径, 提示列表)]；单行出错不拖垮整批。

    一批几十份，中间某一行数据有问题就整批中断是最气人的——那一行记下来
    继续跑，最后一起报。
    """
    from . import overprint
    os.makedirs(out_dir, exist_ok=True)
    made, failed = [], []
    used = set()
    for i, values in enumerate(rows, 1):
        base = safe_name(values.get(name_field) if name_field else '',
                         '第{}行'.format(i))
        stem = '{}{}'.format(prefix, base)
        name = stem
        n = 1
        while name.lower() in used:          # 同名的加序号，别互相覆盖
            n += 1
            name = '{}({})'.format(stem, n)
        used.add(name.lower())
        out = os.path.join(out_dir, name + '.docx')
        try:
            _n, notes = overprint.fill_form(template_path, values, out, **fill_kw)
            made.append((out, notes))
        except Exception as exc:
            failed.append((i, '{}: {}'.format(exc.__class__.__name__, exc)))
            logger.warning('第 %d 行生成失败：%s', i, exc.__class__.__name__)
        if progress:
            progress(i, len(rows))
    return made, failed
