# -*- coding: utf-8 -*-
"""自动编号 → 文字转换

Windows:   用 Word COM 转换（最可靠，100% 准确）
非 Windows: 解析 numbering.xml + numPr 查表转换（标准格式覆盖）
失败时:    返回未转换的段落列表供预览高亮提示
"""
import os
import re
import sys
import zipfile
from io import BytesIO


def _has_auto_numbering(docx_path):
    """快速检测文档是否含自动编号"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        if 'word/numbering.xml' not in z.namelist():
            return False
        doc_xml = z.read('word/document.xml').decode('utf-8')
        return 'w:numPr' in doc_xml


# ============================================================
# Windows: Word COM 转换
# ============================================================

def _convert_via_word_com(input_path, output_path):
    """用 Word COM 把自动编号转为纯文字，失败时返回 False"""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        return False
    try:
        try:
            from win32com.client import Dispatch
            word = Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            try:
                doc = word.Documents.Open(input_path)
                doc.Content.ListFormat.ConvertNumbersToText()
                doc.SaveAs2(output_path, FileFormat=16)
                doc.Close()
                return True
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
        except Exception:
            return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ============================================================
# 非 Windows: 解析 numbering.xml
# ============================================================

def _to_chinese(n):
    """1..9999 阿拉伯数字 → 中文（公文常见范围）：11→十一, 20→二十, 105→一百零五"""
    if n <= 0:
        return str(n)
    digits = '零一二三四五六七八九'
    units = ['', '十', '百', '千']
    if n < 10:
        return digits[n]
    if n <= 99:                      # 10..99：十一、二十、二十三
        tens, ones = divmod(n, 10)
        s = ('' if tens == 1 else digits[tens]) + '十'
        return s + (digits[ones] if ones else '')
    if n > 9999:
        return str(n)
    s = ''
    prev_zero = False
    for i in range(3, -1, -1):
        d = (n // (10 ** i)) % 10
        if d == 0:
            prev_zero = True
        else:
            if prev_zero and s:
                s += '零'
            s += digits[d] + units[i]
            prev_zero = False
    return s


def _to_upper_roman(n):
    if not (1 <= n <= 3999):
        return str(n)
    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
    out = ''
    for v, sym in vals:
        while n >= v:
            out += sym; n -= v
    return out


# Word 编号格式 → python 映射
_NUMFMT_MAP = {
    'decimal':           lambda n: str(n),
    'decimalZero':       lambda n: '{:02d}'.format(n),
    'chineseCounting':   _to_chinese,
    'chineseCountingThousand': _to_chinese,
    'ideographTraditional':    _to_chinese,
    'ideographDigital':  _to_chinese,
    'upperLetter':       lambda n: chr(64 + n) if 1 <= n <= 26 else str(n),
    'lowerLetter':       lambda n: chr(96 + n) if 1 <= n <= 26 else str(n),
    'upperRoman':        _to_upper_roman,
    'lowerRoman':        lambda n: _to_upper_roman(n).lower(),
    'japaneseCounting':  _to_chinese,
    'taiwaneseCounting': _to_chinese,
}

# 编号前后缀映射（常见的 Word 列表模板）
_PREFIX_SUFFIX = {
    ('chineseCounting', 0): ('', '、'),           # 一、
    ('chineseCounting', 1): ('（', '）'),          # （一）
    ('decimal', 0):          ('', '.'),            # 1.
    ('decimal', 1):          ('（', '）'),          # （1）
    ('decimal', 2):          ('', ')'),            # 1)
}


def _parse_numbering(zf):
    """解析 word/numbering.xml 返回 {numId: {ilvl: (fmt_func, prefix, suffix)}}"""
    if 'word/numbering.xml' not in zf.namelist():
        return {}
    import xml.etree.ElementTree as ET
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    tree = ET.parse(BytesIO(zf.read('word/numbering.xml')))
    definitions = {}
    # 抽象编号定义（abstractNum）
    abstract = {}
    for an in tree.iter(ns + 'abstractNum'):
        an_id = an.get(ns + 'abstractNumId')
        levels = {}
        for lvl in an.iter(ns + 'lvl'):
            ilvl = int(lvl.get(ns + 'ilvl', '0'))
            fmt_tag = lvl.find(ns + 'numFmt')
            fmt_val = fmt_tag.get(ns + 'val') if fmt_tag is not None else 'decimal'
            lvl_text = lvl.find(ns + 'lvlText')
            lvl_text_val = lvl_text.get(ns + 'val') if lvl_text is not None else '%1'
            start_tag = lvl.find(ns + 'start')
            try:
                start = int(start_tag.get(ns + 'val')) if start_tag is not None else 1
            except (TypeError, ValueError):
                start = 1
            # 只保留本级占位符 %k（k=ilvl+1）周围的前后缀；把其它级别的
            # 占位符（%1.%2 里的 %1）连同分隔符一并作为前缀，得到如 "1." 里的
            # 前缀部分。lvlText 形如 "%1."→前缀'' 后缀'.'；"%1.%2."→前缀'' 后缀'.'
            prefix, suffix = '', ''
            if lvl_text_val:
                cur = '%{}'.format(ilvl + 1)
                if cur in lvl_text_val:
                    parts = lvl_text_val.split(cur, 1)
                    prefix = re.sub(r'%\d+', '', parts[0])
                    suffix = re.sub(r'%\d+', '', parts[1]) if len(parts) > 1 else ''
                else:
                    parts = lvl_text_val.split('%1', 1)
                    prefix = parts[0] if parts else ''
                    suffix = re.sub(r'%\d+', '', parts[1]) if len(parts) > 1 else ''
            levels[ilvl] = (fmt_val, prefix, suffix, start)
        if an_id is not None:
            abstract[an_id] = levels
    # 编号实例（num）→ 引用 abstractNum
    for num in tree.iter(ns + 'num'):
        num_id = num.get(ns + 'numId')
        ref = num.find(ns + 'abstractNumId')
        if num_id is not None and ref is not None:
            ref_id = ref.get(ns + 'val')
            if ref_id and ref_id in abstract:
                result = {}
                for ilvl, lvldef in abstract[ref_id].items():
                    fmt, prefix, suffix = lvldef[0], lvldef[1], lvldef[2]
                    start = lvldef[3] if len(lvldef) > 3 else 1
                    fmt_func = _NUMFMT_MAP.get(fmt, lambda n: str(n))
                    result[ilvl] = (fmt_func, prefix, suffix, start)
                definitions[num_id] = result
    return definitions


def _convert_via_xml(input_path, output_path):
    """解析 numbering.xml + 文档 XML，转换自动编号为纯文字。异常时静默跳过。"""
    try:
        with zipfile.ZipFile(input_path, 'r') as zf:
            definitions = _parse_numbering(zf)
            if not definitions:
                return False, set()
            doc_xml = zf.read('word/document.xml').decode('utf-8')
            modified, unconverted = _patch_document_xml(doc_xml, definitions)
            if modified == doc_xml:
                return False, set()
            import shutil
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as out:
                for item in zf.infolist():
                    if item.filename == 'word/document.xml':
                        out.writestr(item, modified.encode('utf-8'))
                    else:
                        out.writestr(item, zf.read(item.filename))
        return True, unconverted
    except Exception:
        return False, set()


def _patch_document_xml(doc_xml, definitions):
    """替换文档 XML 中的自动编号为文字。异常时返回原文。"""
    try:
        return _patch_document_xml_impl(doc_xml, definitions)
    except Exception:
        return doc_xml, set()


def _patch_document_xml_impl(doc_xml, definitions):
    import xml.etree.ElementTree as ET
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    ET.register_namespace('', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    tree = ET.parse(BytesIO(doc_xml.encode('utf-8')))
    body = tree.find('.//' + ns + 'body')
    if body is None:
        return doc_xml, set()

    # 状态追踪：{numId: {ilvl: 当前计数}}
    counters = {}
    unconverted = set()  # 未能转换的段落索引
    para_idx = 0

    for para in body.iter(ns + 'p'):
        pPr = para.find(ns + 'pPr')
        if pPr is None:
            para_idx += 1
            continue
        numPr = pPr.find(ns + 'numPr')
        if numPr is None:
            para_idx += 1
            # 检查是否有需要重置的计数器（新编号段落后的连续无编号段落）
            continue
        numId_elem = numPr.find(ns + 'numId')
        ilvl_elem = numPr.find(ns + 'ilvl')
        if numId_elem is None:
            para_idx += 1
            continue
        numId = numId_elem.get(ns + 'val')
        ilvl = int(ilvl_elem.get(ns + 'val', '0')) if ilvl_elem is not None else 0
        if numId not in definitions or ilvl not in definitions[numId]:
            unconverted.add(para_idx)
            para_idx += 1
            continue

        lvldef = definitions[numId][ilvl]
        fmt_func, prefix, suffix = lvldef[0], lvldef[1], lvldef[2]
        start = lvldef[3] if len(lvldef) > 3 else 1
        # 递增计数器（首次出现按 start 起始值）
        c = counters.setdefault(numId, {})
        if ilvl not in c:
            c[ilvl] = start
        else:
            c[ilvl] = c[ilvl] + 1
        # 重置更低层级（下次出现时从各自 start 开始）
        for l in range(ilvl + 1, 10):
            c.pop(l, None)
        number = c[ilvl]
        text = prefix + fmt_func(number) + suffix + ' '

        # 在段落第一个 run 之前插入编号文字
        first_run = para.find(ns + 'r')
        if first_run is not None:
            first_t = first_run.find(ns + 't')
            if first_t is not None:
                first_t.text = text + (first_t.text or '')
        # 从 pPr 中移除 numPr
        pPr.remove(numPr)
        # 清理空 ilvl 元素
        for junk in list(pPr):
            if junk.tag == ns + 'ind':
                pass  # 保留缩进
        para_idx += 1

    result = ET.tostring(tree.getroot(), encoding='unicode')
    return result, unconverted


# ============================================================
# 统一入口
# ============================================================

def convert_auto_numbering(input_path, output_path):
    """把文档中所有自动编号转为纯文字。

    返回: (success: bool, unconverted: set) — unconverted 是未能转换的段落索引（空表示全部成功）
    """
    if not _has_auto_numbering(input_path):
        return True, set()

    if sys.platform == 'win32':
        ok = _convert_via_word_com(input_path, output_path)
        if ok:
            return True, set()
        # COM 失败回退到 XML 方案
        ok2, unc = _convert_via_xml(input_path, output_path)
        return ok2, unc
    else:
        return _convert_via_xml(input_path, output_path)
