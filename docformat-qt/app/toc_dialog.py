# -*- coding: utf-8 -*-
"""生成目录——形式与层级选择（记忆上次选择）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QButtonGroup, QComboBox, QDialog, QHBoxLayout,
                             QLabel, QPushButton, QRadioButton, QVBoxLayout)

from app.theme import settings


class TocOptionsDialog(QDialog):
    """exec_() 返回 Accepted 后，用 get_mode()/get_levels() 取选择。"""

    def __init__(self, parent=None):
        super(TocOptionsDialog, self).__init__(parent)
        self.setWindowTitle("生成目录")
        self.resize(520, 340)
        s = settings()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(10)

        tip = QLabel("排版时已按识别出的标题层级写入大纲级别，两种目录都据此生成。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        self.group = QButtonGroup(self)

        self.rb_auto = QRadioButton("Word 自动目录（推荐）")
        self.rb_auto.setCursor(Qt.PointingHandCursor)
        auto_desc = QLabel("插入 Word 目录域。在 Word/WPS 里右键「更新域」即可得到真实页码，"
                           "条目可点击跳转；正文改动后再更新一次即可，页码自动跟着变。")
        auto_desc.setProperty("muted", "true")
        auto_desc.setWordWrap(True)

        self.rb_manual = QRadioButton("静态目录页")
        self.rb_manual.setCursor(Qt.PointingHandCursor)
        man_desc = QLabel("直接写成文字，不依赖更新域，发给谁看到的都一样。"
                          "点引导线用制表位对齐；本机有 Word/WPS 或 LibreOffice 时"
                          "自动填入真实页码，否则留占位符。")
        man_desc.setProperty("muted", "true")
        man_desc.setWordWrap(True)

        self.group.addButton(self.rb_auto)
        self.group.addButton(self.rb_manual)
        root.addWidget(self.rb_auto)
        root.addWidget(auto_desc)
        root.addSpacing(6)
        root.addWidget(self.rb_manual)
        root.addWidget(man_desc)

        lv_row = QHBoxLayout()
        lv_row.addWidget(QLabel("收录层级："))
        self.levels = QComboBox()
        for label, val in [('到一级标题', 1), ('到二级标题', 2),
                           ('到三级标题（常用）', 3), ('到四级标题', 4)]:
            self.levels.addItem(label, val)
        lv_row.addWidget(self.levels)
        lv_row.addStretch(1)
        root.addSpacing(6)
        root.addLayout(lv_row)

        saved_mode = s.value('toc/mode', 'auto')
        self.rb_manual.setChecked(saved_mode == 'manual')
        self.rb_auto.setChecked(saved_mode != 'manual')
        idx = self.levels.findData(s.value('toc/levels', 3, type=int))
        self.levels.setCurrentIndex(idx if idx >= 0 else 2)

        root.addStretch(1)
        btns = QHBoxLayout()
        btns.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("生成目录")
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self._accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

    def _accept(self):
        s = settings()
        s.setValue('toc/mode', self.get_mode())
        s.setValue('toc/levels', self.get_levels())
        self.accept()

    def get_mode(self):
        return 'manual' if self.rb_manual.isChecked() else 'auto'

    def get_levels(self):
        return self.levels.currentData() or 3
