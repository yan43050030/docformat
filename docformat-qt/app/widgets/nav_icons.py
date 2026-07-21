# -*- coding: utf-8 -*-
"""程序绘制的界面图标：不依赖字体字符（麒麟/UOS 上 emoji 字体不全），
用 QPainter 画简单线条图形，颜色跟随主题。"""
from PyQt5.QtCore import QPointF, QRectF, Qt
from PyQt5.QtGui import QColor, QIcon, QPainter, QPen, QPixmap


def _painter(pm, color, width=1.7):
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(QColor(color))
    pen.setWidthF(width)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    return p


def _pixmap(kind, color, size=18):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = _painter(pm, color)
    s = size

    if kind == 'doc':           # 文档 + 三行文字（格式处理）
        p.drawRoundedRect(QRectF(s*0.22, s*0.12, s*0.56, s*0.76), 2, 2)
        for i, y in enumerate((0.34, 0.5, 0.66)):
            p.drawLine(QPointF(s*0.34, s*y), QPointF(s*0.66 if i < 2 else s*0.54, s*y))
    elif kind == 'layout':      # 版式网格（版式方案）
        p.drawRoundedRect(QRectF(s*0.15, s*0.15, s*0.7, s*0.7), 2, 2)
        p.drawLine(QPointF(s*0.15, s*0.42), QPointF(s*0.85, s*0.42))
        p.drawLine(QPointF(s*0.46, s*0.42), QPointF(s*0.46, s*0.85))
    elif kind == 'pencil':      # 铅笔（文书起草）
        p.drawLine(QPointF(s*0.25, s*0.75), QPointF(s*0.68, s*0.32))
        p.drawLine(QPointF(s*0.68, s*0.32), QPointF(s*0.78, s*0.42))
        p.drawLine(QPointF(s*0.78, s*0.42), QPointF(s*0.35, s*0.85))
        p.drawLine(QPointF(s*0.35, s*0.85), QPointF(s*0.22, s*0.88))
        p.drawLine(QPointF(s*0.22, s*0.88), QPointF(s*0.25, s*0.75))
    elif kind == 'maker':       # 文档 + 加号（模板制作）
        p.drawRoundedRect(QRectF(s*0.18, s*0.15, s*0.5, s*0.66), 2, 2)
        p.drawLine(QPointF(s*0.72, s*0.55), QPointF(s*0.72, s*0.85))
        p.drawLine(QPointF(s*0.57, s*0.7), QPointF(s*0.87, s*0.7))
    elif kind == 'theme':       # 半填充圆（主题）
        p.drawEllipse(QRectF(s*0.18, s*0.18, s*0.64, s*0.64))
        p.setBrush(QColor(color))
        p.drawPie(QRectF(s*0.18, s*0.18, s*0.64, s*0.64), 90*16, 180*16)
    elif kind == 'log':         # 列表行（日志）
        for y in (0.25, 0.5, 0.75):
            p.drawEllipse(QPointF(s*0.22, s*y), s*0.03, s*0.03)
            p.drawLine(QPointF(s*0.34, s*y), QPointF(s*0.8, s*y))
    elif kind == 'drop':        # 文档 + 下落箭头（拖拽区）
        p.drawRoundedRect(QRectF(s*0.28, s*0.08, s*0.44, s*0.5), 2, 2)
        p.drawLine(QPointF(s*0.5, s*0.62), QPointF(s*0.5, s*0.9))
        p.drawLine(QPointF(s*0.38, s*0.78), QPointF(s*0.5, s*0.9))
        p.drawLine(QPointF(s*0.62, s*0.78), QPointF(s*0.5, s*0.9))

    p.end()
    return pm


def make_icon(kind, off_color, on_color=None, size=18):
    """生成主题着色图标；on_color 用于按钮选中态（checkable 按钮自动切换）"""
    icon = QIcon()
    icon.addPixmap(_pixmap(kind, off_color, size), QIcon.Normal, QIcon.Off)
    if on_color:
        icon.addPixmap(_pixmap(kind, on_color, size), QIcon.Normal, QIcon.On)
    return icon


def make_pixmap(kind, color, size=44):
    return _pixmap(kind, color, size)


NAV_ICON_KINDS = ['doc', 'layout', 'pencil', 'maker', 'theme', 'log']
