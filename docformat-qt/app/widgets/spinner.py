# -*- coding: utf-8 -*-
"""轻量转圈指示器：处理进行中显示，让人知道在动、没卡死。"""
from PyQt5.QtCore import Qt, QTimer, QRectF
from PyQt5.QtGui import QColor, QPainter, QPen
from PyQt5.QtWidgets import QWidget


class Spinner(QWidget):
    def __init__(self, size=16, parent=None):
        super(Spinner, self).__init__(parent)
        self._size = size
        self.setFixedSize(size, size)
        self._angle = 0
        self._color = QColor('#C0392B')
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self.hide()

    def set_color(self, color):
        self._color = QColor(color)
        self.update()

    def start(self):
        self.show()
        if not self._timer.isActive():
            self._timer.start(80)

    def stop(self):
        self._timer.stop()
        self.hide()

    def _tick(self):
        self._angle = (self._angle + 45) % 360
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        s = self._size
        m = 2.0
        rect = QRectF(m, m, s - 2 * m, s - 2 * m)
        pen = QPen(self._color, 2.0)
        pen.setCapStyle(Qt.RoundCap)
        p.setPen(pen)
        # 画一段 270° 弧，从当前角度起步
        p.drawArc(rect, -self._angle * 16, 270 * 16)
        p.end()
