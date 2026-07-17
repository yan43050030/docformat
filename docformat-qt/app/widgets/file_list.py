# -*- coding: utf-8 -*-
"""已选文件列表：显示文件名 + 单项移除"""
import os

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
                             QPushButton, QWidget)


class FileList(QListWidget):
    fileRemoved = pyqtSignal(int)

    def __init__(self, parent=None):
        super(FileList, self).__init__(parent)
        self.setSelectionMode(QListWidget.NoSelection)
        self.setMaximumHeight(180)

    def set_files(self, paths):
        self.clear()
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
            btn = QPushButton("移除")
            btn.setProperty("flat", "true")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, idx=i: self.fileRemoved.emit(idx))
            lay.addWidget(tag)
            lay.addWidget(name, 1)
            lay.addWidget(btn)
            item.setSizeHint(row.sizeHint())
            self.setItemWidget(item, row)
