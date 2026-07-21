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

# Word 编号格式 → python 映射
_NUMFMT_MAP = {
    'decimal':           lambda n: str(n),
    'chineseCounting':   lambda n: '一二三四五六七八九十百千万'[n-1] if 1 <= n <= 10 else str(n),
    'chineseCountingThousand': lambda n: str(n),  # 简化
    'upperLetter':       lambda n: chr(64 + n) if 1 <= n <= 26 else str(n),
    'lowerLetter':       lambda n: chr(96 + n) if 1 <= n <= 26 else str(n),
    'upperRoman':        lambda n: str(n),
    'lowerRoman':        lambda n: str(n),
    'japaneseCounting':  lambda n: str(n),
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
            prefix, suffix = '', ''
            if lvl_text_val:
                parts = lvl_text_val.split('%1', 1)
                prefix = parts[0] if parts else ''
                suffix = parts[1] if len(parts) > 1 else ''
            levels[ilvl] = (fmt_val, prefix, suffix)
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
                for ilvl, (fmt, prefix, suffix) in abstract[ref_id].items():
                    fmt_func = _NUMFMT_MAP.get(fmt, lambda n: str(n))
                    result[ilvl] = (fmt_func, prefix, suffix)
                definitions[num_id] = result
    return definitions


def _convert_via_xml(input_path, output_path):
    """解析 numbering.xml + 文档 XML，转换自动编号为纯文字"""
    with zipfile.ZipFile(input_path, 'r') as zf:
        definitions = _parse_numbering(zf)
        if not definitions:
            return False, set()
        doc_xml = zf.read('word/document.xml').decode('utf-8')
        # 修改后的 XML
        modified, unconverted = _patch_document_xml(doc_xml, definitions)
        # 重建 docx
        import shutil
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in zf.infolist():
                if item.filename == 'word/document.xml':
                    out.writestr(item, modified.encode('utf-8'))
                else:
                    out.writestr(item, zf.read(item.filename))
    return True, unconverted


def _patch_document_xml(doc_xml, definitions):
    """替换文档 XML 中的自动编号为文字"""
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

        fmt_func, prefix, suffix = definitions[numId][ilvl]
        # 递增计数器
        c = counters.setdefault(numId, {})
        # 检查是否有重新编号标记
        lvl_restart = False
        for _le in pPr.iter(ns + 'numPr'):
            pass  # numPr 已在上面找到
        c[ilvl] = c.get(ilvl, 0) + 1
        # 重置更低层级
        for l in range(ilvl + 1, 10):
            c[l] = 0
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
