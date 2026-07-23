# -*- coding: utf-8 -*-
"""轻量提示条：右下角一闪而过，替代"已保存/已导出"等打断式弹窗。

用法：Toast.show_message(parent_window, "已导出模板")
重要确认/错误仍用 QMessageBox。
"""
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QPoint
from PyQt5.QtWidgets import QLabel, QGraphicsOpacityEffect


class Toast(QLabel):
    _active = []   # 防止被 GC

    def __init__(self, parent, text, kind='info'):
        super(Toast, self).__init__(parent)
        self.setText(text)
        self.setWordWrap(True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        colors = {
            'info': ('#2E2A24', '#FFFFFF'),
            'success': ('#2E7D32', '#FFFFFF'),
            'warning': ('#B26A00', '#FFFFFF'),
            'error': ('#C62828', '#FFFFFF'),
        }
        bg, fg = colors.get(kind, colors['info'])
        self.setStyleSheet(
            'background: {}; color: {}; border-radius: 9px; '
            'padding: 10px 16px; font-size: 13px;'.format(bg, fg))
        self.adjustSize()
        self._effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self._effect)

    @classmethod
    def show_message(cls, parent, text, kind='info', msec=2200):
        if parent is None:
            return
        win = parent.window()
        t = cls(win, text, kind)
        # 定位到右下角
        margin = 24
        t.move(win.width() - t.width() - margin, win.height() - t.height() - margin)
        t.show()
        t.raise_()
        cls._active.append(t)

        fade_in = QPropertyAnimation(t._effect, b'opacity', t)
        fade_in.setDuration(180)
        fade_in.setStartValue(0.0)
        fade_in.setEndValue(1.0)
        fade_in.start()
        t._fade_in = fade_in

        def _fade_out():
            fo = QPropertyAnimation(t._effect, b'opacity', t)
            fo.setDuration(280)
            fo.setStartValue(1.0)
            fo.setEndValue(0.0)
            fo.finished.connect(lambda: (t.deleteLater(),
                                         cls._active.remove(t) if t in cls._active else None))
            fo.start()
            t._fade_out = fo

        QTimer.singleShot(msec, _fade_out)
        return t
