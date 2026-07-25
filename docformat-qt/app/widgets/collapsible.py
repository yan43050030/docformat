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
        # 只读模式下仍可用的控件（如"规则测试"这类纯查看工具）
        self._always_enabled = []

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

    def mark_always_enabled(self, *widgets):
        """标记只读模式下也保持可用的控件（连同其子控件）。

        典型用途：内置模板参数只读，但"规则测试"只是查看识别结果、
        不修改任何内容，不应被写保护连坐禁用。
        """
        for w in widgets:
            if w is not None and w not in self._always_enabled:
                self._always_enabled.append(w)

    def _is_exempt(self, widget):
        node = widget
        while node is not None and node is not self._body:
            if node in self._always_enabled:
                return True
            node = node.parentWidget()
        return False

    def set_editable(self, editable):
        """只读模式：内容禁用但折叠头仍可展开查看。

        注意不能直接禁用 _body——Qt 中父控件禁用后子控件无法单独启用，
        豁免控件会被连坐。因此逐个控件设置，跳过豁免项。
        """
        self._body.setEnabled(True)
        if editable:
            for w in self._body.findChildren(QWidget):
                w.setEnabled(True)
            return
        for w in self._body.findChildren(QWidget):
            w.setEnabled(self._is_exempt(w))
