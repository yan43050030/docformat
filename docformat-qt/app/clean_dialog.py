# -*- coding: utf-8 -*-
"""格式清洗——清洗项选择面板（分组勾选，记忆上次选择）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts.cleaner import CLEAN_GROUPS, DEFAULT_CLEAN
from app.theme import settings


class CleanItemsDialog(QDialog):
    def __init__(self, current=None, parent=None):
        super(CleanItemsDialog, self).__init__(parent)
        self.setWindowTitle("格式清洗 — 选择清洗项")
        self.resize(560, 600)
        s = settings()
        cur = dict(DEFAULT_CLEAN)
        if current:
            cur.update(current)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(8)

        tip = QLabel("排版出怪问题时，多半是原文档里藏着这些看不见的脏格式。"
                     "默认已勾选安全项；「段落对齐」「修订痕迹」「域代码转文字」"
                     "影响较大，需要时再手动开启。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        scroll.setWidget(host)
        body = QVBoxLayout(host)
        body.setContentsMargins(2, 2, 2, 2)
        body.setSpacing(2)

        self._checks = {}
        for group_name, items in CLEAN_GROUPS:
            gl = QLabel(group_name)
            gl.setProperty("sectionTitle", "true")
            body.addSpacing(8)
            body.addWidget(gl)
            for key, name, desc in items:
                saved = s.value('clean/' + key, cur.get(key, True), type=bool)
                cb = QCheckBox(name)
                cb.setChecked(saved)
                self._checks[key] = cb
                body.addWidget(cb)
                dl = QLabel('　　' + desc)
                dl.setProperty("muted", "true")
                dl.setWordWrap(True)
                body.addWidget(dl)
        body.addStretch(1)
        root.addWidget(scroll, 1)

        btns = QHBoxLayout()
        reset = QPushButton("恢复默认")
        reset.setProperty("flat", "true")
        reset.setCursor(Qt.PointingHandCursor)
        reset.clicked.connect(self._reset)
        btns.addWidget(reset)
        btns.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确定")
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self._accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

    def _reset(self):
        for key, cb in self._checks.items():
            cb.setChecked(DEFAULT_CLEAN.get(key, True))

    def _accept(self):
        s = settings()
        for key, cb in self._checks.items():
            s.setValue('clean/' + key, cb.isChecked())
        self.accept()

    def get_items(self):
        return {key: cb.isChecked() for key, cb in self._checks.items()}
