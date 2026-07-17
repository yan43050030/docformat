# -*- coding: utf-8 -*-
"""排版前后对比预览：左侧原文，右侧按当前预设模拟排版效果，确认后才真正处理"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QComboBox, QDialog, QHBoxLayout, QLabel,
                             QPushButton, QSplitter, QTextBrowser, QVBoxLayout,
                             QWidget)

from scripts.formatter import (_build_text_context, _compile_rules,
                               detect_para_type)

TYPE_LABELS = {
    'security': '密级', 'title': '标题', 'recipient': '主送机关',
    'heading1': '一级标题', 'heading2': '二级标题', 'heading3': '三级标题',
    'heading4': '四级标题', 'body': '正文', 'signature': '署名',
    'date': '日期', 'attachment': '附件', 'closing': '结尾', 'empty': '',
}
MAX_PARAS = 500

ALIGN_CSS = {'left': 'left', 'center': 'center', 'right': 'right', 'justify': 'justify'}


def _read_paragraphs(path):
    """返回 [(text, alignment)]，非 docx 返回 None"""
    if not path.lower().endswith('.docx'):
        return None
    from docx import Document
    doc = Document(path)
    paras = []
    for p in doc.paragraphs[:MAX_PARAS]:
        align = p.paragraph_format.alignment
        paras.append((p.text, align))
    return paras, doc


def _html_shell(body, base_size=12):
    return ('<html><head><style>'
            'body {{ font-family: "SimSun", serif; font-size: {}pt; margin: 14px; }}'
            'p {{ margin: 0 0 4px 0; white-space: pre-wrap; }}'
            '.tag {{ font-size: 8pt; color: #999; background: #F0EDE6; '
            'border-radius: 3px; padding: 0 4px; margin-right: 4px; }}'
            '</style></head><body>{}</body></html>').format(base_size, body)


def render_before_html(paras):
    parts = []
    for text, align in paras:
        if not text.strip():
            parts.append('<p>&nbsp;</p>')
            continue
        a = 'center' if align is not None and 'CENTER' in str(align) else (
            'right' if align is not None and 'RIGHT' in str(align) else 'left')
        parts.append('<p style="text-align:{}">{}</p>'.format(a, _esc(text)))
    return _html_shell(''.join(parts))


def render_after_html(paras, preset):
    rules = _compile_rules(preset.get('detect_rules'))
    texts = [t for t, _ in paras]
    all_texts = [t.strip() for t in texts if t.strip()]
    idx_map = {}
    n = 0
    for i, t in enumerate(texts):
        if t.strip():
            idx_map[i] = n
            n += 1

    parts = []
    prev_type = None
    total = len(paras)
    for i, (text, align) in enumerate(paras):
        if not text.strip():
            parts.append('<p>&nbsp;</p>')
            continue
        ptype = detect_para_type(
            text.strip(), i, total, align, all_texts,
            all_texts_index=idx_map.get(i), prev_para_type=prev_type, rules=rules)
        prev_type = ptype
        fmt = preset.get(ptype if ptype in preset else 'body', preset.get('body', {}))

        style = [
            'font-family:"{}"'.format(fmt.get('font_cn', '仿宋_GB2312')),
            'font-size:{}pt'.format(fmt.get('size', 16)),
            'text-align:{}'.format(ALIGN_CSS.get(fmt.get('align', 'left'), 'left')),
        ]
        indent = fmt.get('indent', 0) or 0
        if indent:
            style.append('text-indent:{}pt'.format(indent))
        ls = fmt.get('line_spacing')
        if ls:
            style.append('line-height:{}pt'.format(ls))
        if fmt.get('bold'):
            style.append('font-weight:bold')

        tag = TYPE_LABELS.get(ptype, ptype)
        parts.append('<p style="{}"><span class="tag">{}</span>{}</p>'.format(
            '; '.join(style), tag, _esc(text)))
    return _html_shell(''.join(parts))


def _esc(text):
    return (text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


class PreviewDialog(QDialog):
    def __init__(self, files, preset, parent=None):
        super(PreviewDialog, self).__init__(parent)
        self.setWindowTitle("排版效果预览（尚未修改任何文件）")
        self.resize(1080, 720)
        self.files = files
        self.preset = preset

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(8)

        top = QHBoxLayout()
        top.addWidget(QLabel("预览文件:"))
        self.file_combo = QComboBox()
        for f in files:
            self.file_combo.addItem(os.path.basename(f), f)
        self.file_combo.currentIndexChanged.connect(self._load_current)
        top.addWidget(self.file_combo, 1)
        tip = QLabel("右侧为预设「{}」的模拟效果，实际以 Word 渲染为准".format(preset.get('name', '')))
        tip.setProperty("muted", "true")
        top.addWidget(tip)
        root.addLayout(top)

        header = QHBoxLayout()
        lab_l = QLabel("排版前（原文）")
        lab_l.setProperty("sectionTitle", "true")
        lab_r = QLabel("排版后（模拟预览，含段落类型标注）")
        lab_r.setProperty("sectionTitle", "true")
        header.addWidget(lab_l, 1)
        header.addWidget(lab_r, 1)
        root.addLayout(header)

        split = QSplitter(Qt.Horizontal)
        self.view_before = QTextBrowser()
        self.view_after = QTextBrowser()
        split.addWidget(self.view_before)
        split.addWidget(self.view_after)
        split.setSizes([500, 500])
        root.addWidget(split, 1)

        btns = QHBoxLayout()
        btns.addStretch(1)
        cancel = QPushButton("关闭")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确认无误，开始排版")
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

        self._load_current()

    def _load_current(self):
        path = self.file_combo.currentData()
        if not path:
            return
        try:
            result = _read_paragraphs(path)
        except Exception as e:
            msg = _html_shell('<p>读取失败: {}</p>'.format(_esc(str(e))))
            self.view_before.setHtml(msg)
            self.view_after.setHtml(msg)
            return
        if result is None:
            msg = _html_shell(
                '<p>该文件为 .doc/.wps 格式，预览暂仅支持 .docx。</p>'
                '<p>点击「开始排版」时会自动转换并处理，不影响最终效果。</p>')
            self.view_before.setHtml(msg)
            self.view_after.setHtml(msg)
            return
        paras, _doc = result
        self.view_before.setHtml(render_before_html(paras))
        self.view_after.setHtml(render_after_html(paras, self.preset))
