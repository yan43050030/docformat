# -*- coding: utf-8 -*-
"""首次启动引导：三个任务入口 + 一句话说明，只显示一次"""
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (QDialog, QFrame, QHBoxLayout, QLabel,
                             QPushButton, QVBoxLayout)

# (标题, 说明, 目标页索引)
_TASKS = [
    ("排版现有文件", "拖入 Word/文本文件，一键排成规范公文", 0),
    ("起草新公文", "选模板、填字段，直接产出排好版的文书", 2),
    ("历史件做模板", "把一份历史公文挖空成可复用的起草模板", 3),
]


class OnboardingDialog(QDialog):
    """exec 后读取 chosen_page（None = 直接开始，停留当前页）"""
    def __init__(self, parent=None):
        super(OnboardingDialog, self).__init__(parent)
        self.setWindowTitle("欢迎使用 DocFormat Pro")
        self.setModal(True)
        self.resize(560, 380)
        self.chosen_page = None

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 18)
        root.setSpacing(12)

        title = QLabel("你想做什么？")
        title.setProperty("h1", "true")
        sub = QLabel("三步完成：选任务 → 拖入文件/填内容 → 开始处理。以后可按 F1 查看使用说明。")
        sub.setProperty("muted", "true")
        sub.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(sub)
        root.addSpacing(6)

        for text, desc, page in _TASKS:
            card = QFrame()
            card.setProperty("modeCard", "true")
            card.setCursor(Qt.PointingHandCursor)
            v = QVBoxLayout(card)
            v.setContentsMargins(16, 12, 16, 12)
            v.setSpacing(3)
            t = QLabel(text)
            t.setProperty("modeCardTitle", "true")
            d = QLabel(desc)
            d.setProperty("muted", "true")
            d.setWordWrap(True)
            v.addWidget(t)
            v.addWidget(d)
            card.mousePressEvent = (lambda _e, p=page: self._choose(p))
            root.addWidget(card)

        root.addStretch(1)
        btns = QHBoxLayout()
        btns.addStretch(1)
        skip = QPushButton("直接开始使用")
        skip.setProperty("primary", "true")
        skip.setCursor(Qt.PointingHandCursor)
        skip.clicked.connect(self.accept)
        btns.addWidget(skip)
        root.addLayout(btns)

    def _choose(self, page):
        self.chosen_page = page
        self.accept()
