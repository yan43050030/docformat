# -*- coding: utf-8 -*-
"""检查项选择面板（分组勾选，记忆上次选择）。

版式检查与用语检查各用一个子类：查的东西不同、说明不同、默认值不同，
但勾选面板的骨架是一样的，共用一个基类，别抄两遍。
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts.compliance import (CHECK_GROUPS, DEFAULT_OPTIONS, LAYOUT_KEYS,
                                WORDING_GROUPS, WORDING_KEYS, only)
from app.theme import settings


class _OptionsDialog(QDialog):
    """勾选面板骨架。子类给 GROUPS / KEYS / 标题 / 说明。"""

    GROUPS = []
    KEYS = []
    TITLE = ''
    INTRO = ''
    OK_TEXT = '开始检查'

    def __init__(self, parent=None):
        super(_OptionsDialog, self).__init__(parent)
        self.setWindowTitle(self.TITLE)
        self.resize(480, 560)
        s = settings()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(8)

        tip = QLabel(self.INTRO)
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        scroll.setWidget(host)
        body = QVBoxLayout(host)
        body.setContentsMargins(2, 2, 2, 2)
        body.setSpacing(4)

        self._checks = {}
        for group_name, items in self.GROUPS:
            gl = QLabel(group_name)
            gl.setProperty("sectionTitle", "true")
            body.addSpacing(6)
            body.addWidget(gl)
            for key, label in items:
                cb = QCheckBox(label)
                cb.setChecked(s.value('compliance/' + key,
                                      DEFAULT_OPTIONS.get(key, True), type=bool))
                self._checks[key] = cb
                body.addWidget(cb)
        self._extra(body)
        body.addStretch(1)
        root.addWidget(scroll, 1)

        btns = QHBoxLayout()
        all_btn = QPushButton("全选")
        all_btn.setProperty("flat", "true")
        all_btn.setCursor(Qt.PointingHandCursor)
        all_btn.clicked.connect(lambda: self._set_all(True))
        none_btn = QPushButton("全不选")
        none_btn.setProperty("flat", "true")
        none_btn.setCursor(Qt.PointingHandCursor)
        none_btn.clicked.connect(lambda: self._set_all(False))
        btns.addWidget(all_btn)
        btns.addWidget(none_btn)
        btns.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton(self.OK_TEXT)
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self._accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

    def _extra(self, _body):
        """子类往面板底部加自己的东西"""

    def _set_all(self, on):
        for cb in self._checks.values():
            cb.setChecked(on)

    def _accept(self):
        s = settings()
        for key, cb in self._checks.items():
            s.setValue('compliance/' + key, cb.isChecked())
        self.accept()

    def get_options(self):
        """只返回本面板管的那几项，其余一律关掉。

        两个入口各管一摊：查版式时不该顺带报错别字，查用语时也不该跳出来
        说页边距不对——用户是带着一个明确目的点进来的。
        """
        return only(self.KEYS,
                    {k: cb.isChecked() for k, cb in self._checks.items()})


class ComplianceOptionsDialog(_OptionsDialog):
    GROUPS = CHECK_GROUPS
    KEYS = LAYOUT_KEYS
    TITLE = "公文合规检查 — 选择检查项"
    INTRO = ("检查标准来自你当前选中的预设——你的公文与国标有差异时，改预设即可，"
             "检查会自动跟着变。段落项会逐段按识别出的类型（标题/正文/各级小标题…）"
             "对照预设核对，全部通过即等于排版合规。<br><br>"
             "查的是<b>版式</b>；文字本身对不对（错别字、数字用法、文种搭配）"
             "在首页的「公文用语检查」里。")


class WordingOptionsDialog(_OptionsDialog):
    GROUPS = WORDING_GROUPS
    KEYS = WORDING_KEYS
    TITLE = "公文用语检查 — 选择检查项"
    INTRO = ("只查<b>有明文规定</b>的用语问题——错别字、数字用法、文种与结语搭配、"
             "标题标点这一类，白纸黑字有依据的才报。文风好坏、写得顺不顺，"
             "机器不懂，一概不管。<br><br>"
             "查出来的问题会先列成清单，涉及改动文字的<b>一律走 Word 修订</b>，"
             "你逐条看过、认可了才落笔。")
