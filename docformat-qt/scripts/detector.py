# -*- coding: utf-8 -*-
"""
段落检测 — 从 formatter.py 拆分

包含：detect_para_type、_compile_rules、_build_text_context、
      日期/标题识别、上下文感知逻辑、DEFAULT_DETECT_RULES
"""

import re
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ===== 日期行的默认识别模式 =====
_DATE_PATTERNS = (
    r'^\d{4}年\d{1,2}月\d{1,2}日$',
    r'^\d{4}年\d{1,2}月\d{1,2}$',
    r'^\d{4}年\d{1,2}月$',
    r'^\d{4}\.\d{1,2}\.\d{1,2}$',
    r'^\d{4}\.\d{1,2}$',
    r'^\d{4}/\d{1,2}/\d{1,2}$',
    r'^\d{4}/\d{1,2}$',
    r'^\d{4}-\d{1,2}-\d{1,2}$',
    r'^\d{4}-\d{1,2}$',
    r'^二[○〇零oO0][一二三四五六七八九零〇○oO0]{2}年.{1,3}月.{1,3}日$',
    r'^二[○〇零oO0][一二三四五六七八九零〇○oO0]{2}年.{1,3}月$',
)

_CLOSING_PATTERNS = (
    r'^特此(说明|通知|报告|函复|函告|批复|公告|通报)。?$',
    r'^此致$',
    r'^敬礼[！!]?$',
    r'^以上(报告|意见|方案).{0,10}$',
    r'^妥否.{0,10}$',
    r'^请.{0,15}(批示|审批|审议|指示|核准)。?$',
)

DEFAULT_DETECT_RULES = {
    'security': r'^(绝密|机密|秘密)\s*[★\*]?\s*([一二三四五六七八九十0-9]+\s*(年|个月|月))?\s*$',
    'docnum': r'^[一-鿿]{2,20}[〔\[](19|20)\d{2}[〕\]]\s*第?\s*\d+\s*号$',
    'heading1': r'^[一二三四五六七八九十]+、',
    'heading2': r'^（[一二三四五六七八九十]+）|^\([一二三四五六七八九十]+\)',
    'heading3': r'^\d+\.\s*[^\d.\s]',
    'heading4': r'^（\d+）|^\(\d+\)',
    'recipient': r'^[一-鿿\d、，,（）()\s]+[：:]$',
    'attachment': r'^附件\d*([：:．.\s]|$)',
    'closing': '|'.join(_CLOSING_PATTERNS),
    'date': '|'.join(_DATE_PATTERNS),
    'signature': r'(公司|局|委|部|厅|院|所|室|中心|办公室|集团|银行|学校|大学|医院'
                 r'|指挥部|领导小组|委员会|管理处|管委会)$',
}

_CLOSING_RES = [re.compile(p) for p in _CLOSING_PATTERNS]
_DATE_RES = [re.compile(p) for p in _DATE_PATTERNS]


def _build_text_context(doc):
    """Collect non-empty paragraph texts and map document indexes to text indexes."""
    all_texts = []
    all_texts_idx_map = {}
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            all_texts_idx_map[i] = len(all_texts)
            all_texts.append(text)
    return all_texts, all_texts_idx_map


def _compile_rules(overrides):
    """合并用户自定义识别规则，非法正则自动回退默认值"""
    rules = {}
    overrides = overrides or {}
    for key, default in DEFAULT_DETECT_RULES.items():
        pattern = overrides.get(key) or default
        try:
            rules[key] = re.compile(pattern)
        except re.error:
            rules[key] = re.compile(default)
    return rules


def _is_date_text(text, date_patterns):
    """检查文本是否匹配日期模式"""
    for pattern in date_patterns:
        if pattern.match(text):
            return True
    return False


def _normalize_date_text(text):
    """标准化中文日期格式，去除多余空格"""
    return re.sub(r'\s+', '', text)


def _standardize_date_text(text):
    """标准化日期文本，统一为 YYYY年M月D日 格式"""
    text = _normalize_date_text(text)
    # 2026.7.19 → 2026年7月19日
    m = re.match(r'^(\d{4})\.(\d{1,2})\.(\d{1,2})$', text)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    # 2026-7-19 → 2026年7月19日
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', text)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    # 2026/7/19 → 2026年7月19日
    m = re.match(r'^(\d{4})/(\d{1,2})/(\d{1,2})$', text)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    # 2026.7 → 2026年7月
    m = re.match(r'^(\d{4})\.(\d{1,2})$', text)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月"
    return text


def detect_para_type(text, index, total, alignment, all_texts, all_texts_index=None, prev_para_type=None, rules=None, flags=None):
    """检测段落类型

    返回: 'title', 'subtitle', 'recipient', 'heading1'..4, 'body', 'signature',
          'date', 'attachment', 'attachment_label', 'closing', 'security', 'docnum',
          'copynum', 'urgency', 'signatory', 'cc', 'issuer'（后几类需 flags 开启）
    flags: 可选 dict，控制默认关闭的版头/版记/副标题识别：
           subtitle_enabled / header_elements / record_elements
    """
    text = text.strip()
    if not text:
        return 'empty'

    flags = flags or {}
    _rules = rules if isinstance(rules, dict) and all(hasattr(v, 'match') for v in rules.values()) else _compile_rules(rules)

    _early_idx = all_texts_index if all_texts_index is not None else index
    _tot = len(all_texts) if all_texts else total

    # ===== 版头要素（默认关闭，header_elements 开启后识别）=====
    if flags.get('header_elements') and _early_idx < 8:
        # 份号：纯数字行（通常 6 位），位于版头最前
        if re.match(r'^\d{4,8}$', text):
            return 'copynum'
        # 紧急程度：特急/加急/平急
        if re.match(r'^(特急|加急|平急|急件)$', text):
            return 'urgency'
        # 签发人：签发人：×××
        if re.match(r'^签发人[：:]\s*\S', text):
            return 'signatory'

    # ===== 密级标识检测 =====
    if _early_idx < 3 and _rules['security'].match(text):
        return 'security'

    # ===== 发文字号检测 =====
    if _early_idx < 6 and _rules['docnum'].match(text):
        return 'docnum'

    closing_patterns = [_rules['closing']]
    date_patterns = [_rules['date']]

    # ===== 副标题（默认关闭）：标题下方以破折号/括号引起的居中短行 =====
    if flags.get('subtitle_enabled') and prev_para_type in ('title', 'subtitle'):
        if re.match(r'^(——|--|－－|—|-)\s*\S', text) or re.match(r'^[（(].{2,40}[）)]$', text):
            if len(text) < 50:
                return 'subtitle'

    # ===== 版记要素（默认关闭，record_elements 开启后识别）=====
    if flags.get('record_elements'):
        if re.match(r'^抄送[：:]', text):
            return 'cc'
        # 印发机关和印发日期：××× 印发 / ×××××年×月×日印发
        if re.search(r'印发$', text) or re.search(r'印发\s*$', text):
            return 'issuer'

    # ===== 标题续行检测 =====
    if prev_para_type == 'title':
        heading_prefix = re.match(r'^[一二三四五六七八九十（\(\d]', text)
        is_recipient_end = re.search(r'[：:]\s*$', text)
        is_attachment = _rules['attachment'].match(text)
        is_closing = any(pattern.match(text) for pattern in closing_patterns)
        is_date = _is_date_text(text, date_patterns)
        is_sentence_end = re.search(r'[。！？.!?；;]\s*$', text)
        if not heading_prefix and not is_recipient_end and not is_attachment and not is_closing and not is_date and not is_sentence_end and len(text) < 50:
            return 'title'

    # ===== 早期日期检测 =====
    if all_texts_index is not None:
        is_in_tail = all_texts_index >= len(all_texts) * 2 // 3
    else:
        is_in_tail = index >= total * 2 // 3
    if is_in_tail:
        if _is_date_text(text, date_patterns):
            return 'date'

    # ===== 附件块延续识别 =====
    if prev_para_type == 'attachment':
        # 编号延续（2.XXX / 3、XXX）
        if re.match(r'^\s*\d{1,2}[.、]\s*\S', text):
            return 'attachment'
        # 换行延续：无编号的较长纯文本行（>15字），且不匹配署名/日期/结束语
        if (len(text) > 15
                and not _rules['signature'].search(text)
                and not _is_date_text(text, date_patterns)
                and not any(pattern.match(text) for pattern in closing_patterns)):
            return 'attachment'

    # ===== 一级标题："一、" "二、" 等 =====
    if _rules['heading1'].match(text):
        return 'heading1'

    # ===== 二级标题："（一）" "（二）" 等 =====
    if _rules['heading2'].match(text):
        return 'heading2'

    # ===== 三级标题："1." "2." 等 =====
    if _rules['heading3'].match(text) and len(text) < 60:
        return 'heading3'

    # ===== 四级标题："（1）" "（2）" 等 =====
    if _rules['heading4'].match(text) and len(text) < 60:
        return 'heading4'

    # ===== 主送机关 =====
    if _rules['recipient'].match(text) and len(text) < 30:
        body_indicators = (
            r'(现将|为了|根据|按照|经研究|为贯彻|为落实|为进一步|为深入|'
            r'如下|以下|特此|兹将|报告如下|说明如下|通知如下|汇报如下|'
            r'的意见|的通知|的报告|的决定|的请示|的函)'
        )
        if not re.search(body_indicators, text):
            if all_texts_index is not None:
                next_texts = all_texts[all_texts_index + 1: all_texts_index + 2]
                for nt in next_texts:
                    nt = nt.strip()
                    if re.match(r'^关于.+的(通知|报告|请示|函|意见|决定|公告|通报|批复)', nt):
                        break
                    if 15 < len(nt) < 80 and not re.search(r'[。！？，、；：]$', nt):
                        break
                else:
                    return 'recipient'
            else:
                return 'recipient'

    # ===== 附件标识行（独占一行的"附件""附件1""附件一"，无冒号无后续内容）=====
    # 附件另起页后，"附件"及序号顶格黑体；与正文里的"附件1：说明"(悬挂缩进) 区分。
    # 关键区分：带冒号/有后续内容 → 说明；纯"附件N" → 标识。
    # 再加位置保险：仅文档中后段（非空段过 1/3 处）才认定为标识，避免正文
    # 前部偶发的纯"附件N"被误判——正文末尾列附件通常带冒号，不受影响。
    if re.match(r'^附件\s*[一二三四五六七八九十\d]*\s*$', text):
        _pos = all_texts_index if all_texts_index is not None else index
        _tot = len(all_texts) if all_texts else total
        if _tot and _pos >= _tot / 3.0:
            return 'attachment_label'

    # ===== 附件说明行 =====
    if _rules['attachment'].match(text):
        return 'attachment'

    # ===== 结束语 =====
    for pattern in closing_patterns:
        if pattern.match(text):
            return 'closing'

    # ===== 落款日期 =====
    if _is_date_text(text, date_patterns):
        return 'date'

    # ===== 落款单位 =====
    if index >= total - 10 and len(text) < 60:
        allow_signature_check = True
        title_region_idx = all_texts_index if all_texts_index is not None else index
        if title_region_idx < 5:
            previous_texts = all_texts[:title_region_idx] if all_texts_index is not None else all_texts[:index]
            has_document_body_started = any(
                re.search(r'[：:]\s*$', pt.strip()) or re.match(r'^[一二三四五六七八九十]+、', pt.strip())
                for pt in previous_texts
            )
            if not has_document_body_started:
                allow_signature_check = False

        if allow_signature_check:
            if _rules['signature'].search(text):
                return 'signature'
            if re.search(r'[。！？.!?；;]\s*$', text):
                return 'body'
            if all_texts_index is not None:
                remaining_texts = all_texts[all_texts_index + 1:]
            else:
                remaining_texts = []
            for next_text in remaining_texts[:3]:
                if _is_date_text(next_text, date_patterns):
                    return 'signature'

    # ===== 主标题 =====
    title_region_idx = all_texts_index if all_texts_index is not None else index
    if title_region_idx < 5:
        _check_idx = all_texts_index if all_texts_index is not None else 0
        _title_region_ended = False
        for pt in all_texts[:_check_idx]:
            pt_s = pt.strip()
            if re.search(r'[：:]\s*$', pt_s) and len(pt_s) < 50:
                _title_region_ended = True
                break
            if re.match(r'^[一二三四五六七八九十]+、', pt_s):
                _title_region_ended = True
                break

        if not _title_region_ended:
            title_patterns = [
                r'^关于.+的(通知|报告|请示|函|意见|决定|公告|通报|批复|说明|方案|总结|汇报|复函|答复|建议)$',
                r'^.{2,30}(通知|报告|请示|函|意见|决定|公告|通报|批复|工作方案|工作总结|实施方案|管理办法|暂行规定)$',
                r'^[\u4e00-\u9fff]{2,20}(委员会|办公室|局|厅|院|部|委|中心|公司|集团|学校|大学)$',
            ]
            for pattern in title_patterns:
                if re.match(pattern, text):
                    return 'title'

            if 15 < len(text) < 80 and not re.search(r'[。！？，、；：.!?,;:]$', text):
                if not re.match(r'^[一二三四五六七八九十\d（(]', text):
                    return 'title'

            if alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) < 60:
                return 'title'

    # ===== 其他都是正文 =====
    return 'body'
