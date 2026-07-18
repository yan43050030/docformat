# -*- coding: utf-8 -*-
"""已选文件列表：文件名 + 处理状态 + 单项移除"""
import os

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
                             QPushButton, QWidget)

# 状态文本；颜色由主题 QSS 按 statusLevel 属性着色（深色主题自动用亮色）
_STATUS_TEXT = {
    'pending': '',
    'processing': '⏳ 处理中…',
    'ok': '✓ 完成',
    'fail': '✗ 失败',
}


class FileList(QListWidget):
    fileRemoved = pyqtSignal(int)

    def __init__(self, parent=None):
        super(FileList, self).__init__(parent)
        self.setSelectionMode(QListWidget.NoSelection)
        self.setMaximumHeight(200)
        self._status_labels = {}   # normpath -> QLabel

    def set_files(self, paths):
        self.clear()
        self._status_labels = {}
        for i, p in enumerate(paths):
            item = QListWidgetItem(self)
            row = QWidget()
            lay = QHBoxLayout(row)
            lay.setContentsMargins(6, 2, 6, 2)
            lay.setSpacing(8)

            ext = os.path.splitext(p)[1].lower().lstrip('.')
            tag = QLabel(ext)
            tag.setProperty("badge", "true")

            name = QLabel(os.path.basename(p))
            name.setToolTip(p)

            status = QLabel("")
            status.setProperty("statusLevel", "pending")
            self._status_labels[os.path.normpath(p)] = status

            btn = QPushButton("✕")
            btn.setProperty("removeBtn", "true")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setToolTip("从列表移除该文件（不会删除文件本身）")
            btn.setFixedWidth(30)
            btn.clicked.connect(lambda _=False, idx=i: self.fileRemoved.emit(idx))

            lay.addWidget(tag)
            lay.addWidget(name, 1)
            lay.addWidget(status)
            lay.addSpacing(8)
            lay.addWidget(btn)

            hint = row.sizeHint()
            item.setSizeHint(QSize(hint.width(), max(34, hint.height())))
            self.setItemWidget(item, row)

    def set_status(self, path, state, tooltip=''):
        """state: pending / processing / ok / fail"""
        label = self._status_labels.get(os.path.normpath(path))
        if label is None:
            return
        label.setText(_STATUS_TEXT.get(state, ''))
        label.setToolTip(tooltip)
        label.setProperty("statusLevel", state)
        # 属性变化后需重新抛光样式才会生效
        label.style().unpolish(label)
        label.style().polish(label)

    def reset_statuses(self):
        for label in self._status_labels.values():
            label.setText('')
            label.setToolTip('')
            label.setProperty("statusLevel", "pending")
            label.style().unpolish(label)
            label.style().polish(label)
