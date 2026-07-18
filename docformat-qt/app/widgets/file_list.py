# -*- coding: utf-8 -*-
"""已选文件列表：显示文件名 + 处理状态 + 单项移除"""
import os

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
                             QPushButton, QWidget)

# 状态: text, color
_STATUS_STYLES = {
    'pending': ('', ''),
    'processing': ('处理中…', '#B26A00'),
    'ok': ('✓ 完成', '#2E7D32'),
    'fail': ('✗ 失败', '#C62828'),
}


class FileList(QListWidget):
    fileRemoved = pyqtSignal(int)

    def __init__(self, parent=None):
        super(FileList, self).__init__(parent)
        self.setSelectionMode(QListWidget.NoSelection)
        self.setMaximumHeight(180)
        self._status_labels = {}   # normpath -> QLabel

    def set_files(self, paths):
        self.clear()
        self._status_labels = {}
        for i, p in enumerate(paths):
            item = QListWidgetItem(self)
            row = QWidget()
            lay = QHBoxLayout(row)
            lay.setContentsMargins(4, 0, 4, 0)
            name = QLabel(os.path.basename(p))
            name.setToolTip(p)
            ext = os.path.splitext(p)[1].lower().lstrip('.')
            tag = QLabel(ext)
            tag.setProperty("badge", "true")
            status = QLabel("")
            self._status_labels[os.path.normpath(p)] = status
            btn = QPushButton("移除")
            btn.setProperty("flat", "true")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, idx=i: self.fileRemoved.emit(idx))
            lay.addWidget(tag)
            lay.addWidget(name, 1)
            lay.addWidget(status)
            lay.addWidget(btn)
            item.setSizeHint(row.sizeHint())
            self.setItemWidget(item, row)

    def set_status(self, path, state, tooltip=''):
        """state: pending / processing / ok / fail"""
        label = self._status_labels.get(os.path.normpath(path))
        if label is None:
            return
        text, color = _STATUS_STYLES.get(state, ('', ''))
        label.setText(text)
        label.setStyleSheet('color: {}; background: transparent;'.format(color) if color else '')
        label.setToolTip(tooltip)

    def reset_statuses(self):
        for label in self._status_labels.values():
            label.setText('')
            label.setToolTip('')
