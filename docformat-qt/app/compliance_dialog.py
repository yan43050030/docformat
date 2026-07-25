# -*- coding: utf-8 -*-
"""公文合规检查——检查项选择面板（分组勾选，记忆上次选择）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts.compliance import CHECK_GROUPS
from app.theme import settings


class ComplianceOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super(ComplianceOptionsDialog, self).__init__(parent)
        self.setWindowTitle("公文合规检查 — 选择检查项")
        self.resize(480, 560)
        s = settings()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(8)

        tip = QLabel("检查标准来自你当前选中的预设——你的公文与国标有差异时，改预设即可，"
                     "检查会自动跟着变。段落项会逐段按识别出的类型（标题/正文/各级小标题…）"
                     "对照预设核对，全部通过即等于排版合规。")
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
        for group_name, items in CHECK_GROUPS:
            gl = QLabel(group_name)
            gl.setProperty("sectionTitle", "true")
            body.addSpacing(6)
            body.addWidget(gl)
            for key, label in items:
                cb = QCheckBox(label)
                cb.setChecked(s.value('compliance/' + key, True, type=bool))
                self._checks[key] = cb
                body.addWidget(cb)
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
        ok = QPushButton("开始检查")
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self._accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

    def _set_all(self, on):
        for cb in self._checks.values():
            cb.setChecked(on)

    def _accept(self):
        s = settings()
        for key, cb in self._checks.items():
            s.setValue('compliance/' + key, cb.isChecked())
        self.accept()

    def get_options(self):
        return {key: cb.isChecked() for key, cb in self._checks.items()}
