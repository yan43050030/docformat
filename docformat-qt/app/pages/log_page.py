# -*- coding: utf-8 -*-
"""日志页：等级着色的处理日志"""
import time

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QHBoxLayout, QLabel, QPushButton, QTextEdit,
                             QVBoxLayout, QWidget)

LEVEL_COLORS = {
    'info': '#5C7CFA',
    'success': '#2E7D32',
    'warning': '#B26A00',
    'error': '#C62828',
}
LEVEL_LABELS = {'info': '信息', 'success': '成功', 'warning': '警告', 'error': '错误'}


class LogPage(QWidget):
    def __init__(self, parent=None):
        super(LogPage, self).__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 16)
        root.setSpacing(12)

        head = QHBoxLayout()
        title = QLabel("处理日志")
        title.setProperty("h1", "true")
        clear_btn = QPushButton("清空日志")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self.clear)
        head.addWidget(title)
        head.addStretch(1)
        head.addWidget(clear_btn)
        root.addLayout(head)

        self.view = QTextEdit()
        self.view.setObjectName("LogView")
        self.view.setReadOnly(True)
        root.addWidget(self.view, 1)

    def append(self, level, message):
        color = LEVEL_COLORS.get(level, '#888888')
        label = LEVEL_LABELS.get(level, level)
        ts = time.strftime('%H:%M:%S')
        safe = (message.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;').replace('\n', '<br>'))
        self.view.append(
            '<span style="color:#8A8375">[{}]</span> '
            '<b style="color:{}">[{}]</b> {}'.format(ts, color, label, safe))

    def clear(self):
        self.view.clear()
