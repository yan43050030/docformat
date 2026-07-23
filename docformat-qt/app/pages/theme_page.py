# -*- coding: utf-8 -*-
"""主题页：四主题卡片切换"""
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (QFrame, QGridLayout, QHBoxLayout, QLabel,
                             QPushButton, QVBoxLayout, QWidget)

from app.theme import THEMES, current_theme_id, raw_theme_id, save_theme_id


class ThemePage(QWidget):
    themeSelected = pyqtSignal(str)

    def __init__(self, parent=None):
        super(ThemePage, self).__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 16)
        root.setSpacing(14)

        title = QLabel("界面主题")
        title.setProperty("h1", "true")
        sub = QLabel("选择喜欢的配色方案，选择立即生效并自动记忆")
        sub.setProperty("muted", "true")
        root.addWidget(title)
        root.addWidget(sub)

        grid = QGridLayout()
        grid.setSpacing(14)
        self._buttons = {}
        for i, (tid, c) in enumerate(THEMES.items()):
            card = QFrame()
            card.setProperty("card", "true")
            lay = QVBoxLayout(card)
            lay.setContentsMargins(16, 14, 16, 14)
            lay.setSpacing(8)

            sw_row = QHBoxLayout()
            sw_row.setSpacing(4)
            for color in [c['bg'], c['card'], c['accent'], c['teal'], c['ink']]:
                sw = QLabel()
                sw.setFixedSize(26, 26)
                sw.setStyleSheet(
                    "background: {}; border-radius: 13px; border: 1px solid rgba(0,0,0,40);".format(color))
                sw_row.addWidget(sw)
            sw_row.addStretch(1)
            lay.addLayout(sw_row)

            name = QLabel(c['name'])
            name.setProperty("sectionTitle", "true")
            desc = QLabel(c['desc'])
            desc.setProperty("muted", "true")
            btn = QPushButton("使用此主题")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, t=tid: self._select(t))
            lay.addWidget(name)
            lay.addWidget(desc)
            lay.addWidget(btn)
            self._buttons[tid] = btn
            grid.addWidget(card, i // 2, i % 2)
        root.addLayout(grid)

        # 跟随系统明暗
        auto_row = QHBoxLayout()
        self.auto_btn = QPushButton("🌓 跟随系统明暗（自动在纸质/暗夜间切换）")
        self.auto_btn.setCursor(Qt.PointingHandCursor)
        self.auto_btn.clicked.connect(lambda: self._select('auto'))
        auto_row.addWidget(self.auto_btn)
        auto_row.addStretch(1)
        root.addLayout(auto_row)

        root.addStretch(1)
        self._refresh_checked()

    def _select(self, tid):
        save_theme_id(tid)
        from app.theme import resolve_theme_id
        self.themeSelected.emit(resolve_theme_id(tid))
        self._refresh_checked()

    def _refresh_checked(self):
        cur = raw_theme_id()
        if hasattr(self, 'auto_btn'):
            is_auto = (cur == 'auto')
            self.auto_btn.setText("✓ 已跟随系统" if is_auto else "🌓 跟随系统明暗（自动在纸质/暗夜间切换）")
            self.auto_btn.setProperty("primary", "true" if is_auto else "false")
            self.auto_btn.style().unpolish(self.auto_btn)
            self.auto_btn.style().polish(self.auto_btn)
        for tid, btn in self._buttons.items():
            if tid == cur:
                btn.setText("✓ 当前主题")
                btn.setProperty("primary", "true")
            else:
                btn.setText("使用此主题")
                btn.setProperty("primary", "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
