# -*- coding: utf-8 -*-
"""公文合规检查——检查项选择面板（可勾选查哪些，记忆上次选择）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QHBoxLayout, QLabel,
                             QPushButton, QVBoxLayout)

from scripts.compliance import CHECK_ITEMS
from app.theme import settings


class ComplianceOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super(ComplianceOptionsDialog, self).__init__(parent)
        self.setWindowTitle("公文合规检查 — 选择检查项")
        self.resize(460, 420)
        s = settings()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(8)

        tip = QLabel("检查标准来自你当前选中的预设方案——你的公文与国标有差异时，"
                     "改预设即可，检查会自动跟着变。勾选要检查的项：")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        self._checks = {}
        for key, label in CHECK_ITEMS:
            cb = QCheckBox(label)
            cb.setChecked(s.value('compliance/' + key, True, type=bool))
            self._checks[key] = cb
            root.addWidget(cb)

        root.addStretch(1)
        btns = QHBoxLayout()
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

    def _accept(self):
        s = settings()
        for key, cb in self._checks.items():
            s.setValue('compliance/' + key, cb.isChecked())
        self.accept()

    def get_options(self):
        return {key: cb.isChecked() for key, cb in self._checks.items()}
