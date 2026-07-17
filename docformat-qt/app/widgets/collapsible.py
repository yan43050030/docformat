# -*- coding: utf-8 -*-
"""折叠分组卡片（预设编辑器用）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFrame, QToolButton, QVBoxLayout, QWidget


class CollapsibleSection(QFrame):
    def __init__(self, title, expanded=False, parent=None):
        super(CollapsibleSection, self).__init__(parent)
        self.setProperty("card", "true")

        self._header = QToolButton()
        self._header.setProperty("collapsibleHeader", "true")
        self._header.setText(title)
        self._header.setCheckable(True)
        self._header.setChecked(expanded)
        self._header.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self._header.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self._header.setCursor(Qt.PointingHandCursor)
        self._header.clicked.connect(self._on_toggle)

        self._body = QWidget()
        self._body.setVisible(expanded)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 10)
        lay.setSpacing(8)
        lay.addWidget(self._header)
        lay.addWidget(self._body)

    def _on_toggle(self):
        expanded = self._header.isChecked()
        self._header.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self._body.setVisible(expanded)

    def set_body_layout(self, layout):
        layout.setContentsMargins(2, 4, 2, 4)
        self._body.setLayout(layout)

    def set_editable(self, editable):
        """只读模式：内容禁用但折叠头仍可展开查看"""
        self._body.setEnabled(editable)
