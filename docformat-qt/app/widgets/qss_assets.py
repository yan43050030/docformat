# -*- coding: utf-8 -*-
"""按主题渲染并缓存控件指示器图片（复选框/单选/下拉箭头），供 QSS 引用。

原生复选框/单选钮在麒麟/UOS 与暗色主题下外观参差，自绘统一为主题色勾选。
图片缓存到配置目录，按主题+版本命名，缺失时重绘。QSS 用 url() 引用。
"""
import os

from PyQt5.QtCore import QPointF, QRectF, Qt
from PyQt5.QtGui import QColor, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import QApplication

_ASSET_VERSION = 2   # 改动绘制逻辑时 +1 触发重绘


def _cache_dir():
    from app.presets import config_dir
    d = config_dir() / 'ui_cache'
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def _new_pixmap(size):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    return pm


def _painter(pm):
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    return p


def _draw_checkbox(size, bg, border, fill=None, check=None):
    pm = _new_pixmap(size)
    p = _painter(pm)
    r = QRectF(1.5, 1.5, size - 3, size - 3)
    p.setPen(QPen(QColor(border), 1.5))
    p.setBrush(QColor(fill) if fill else QColor(bg))
    p.drawRoundedRect(r, 4, 4)
    if check:
        pen = QPen(QColor(check), 2.0)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        p.setPen(pen)
        s = size
        p.drawPolyline(QPointF(s*0.28, s*0.52), QPointF(s*0.44, s*0.68), QPointF(s*0.74, s*0.32))
    p.end()
    return pm


def _draw_radio(size, bg, border, fill=None, dot=None):
    pm = _new_pixmap(size)
    p = _painter(pm)
    r = QRectF(1.5, 1.5, size - 3, size - 3)
    p.setPen(QPen(QColor(border), 1.5))
    p.setBrush(QColor(fill) if fill else QColor(bg))
    p.drawEllipse(r)
    if dot:
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(dot))
        c = size / 2.0
        rr = size * 0.22
        p.drawEllipse(QPointF(c, c), rr, rr)
    p.end()
    return pm


def _draw_chevron(size, color):
    pm = _new_pixmap(size)
    p = _painter(pm)
    pen = QPen(QColor(color), 1.6)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    s = size
    p.drawPolyline(QPointF(s*0.3, s*0.42), QPointF(s*0.5, s*0.62), QPointF(s*0.7, s*0.42))
    p.end()
    return pm


def ensure_assets(tid, colors):
    """确保当前主题的指示器图片已生成，返回 {名称: posix 路径} 供 QSS url() 使用。

    无 QApplication（极早期）时返回 None，QSS 退化为无图（仍可用）。
    """
    if QApplication.instance() is None:
        return None
    d = _cache_dir()
    c = colors
    specs = {
        'cb_off':   lambda: _draw_checkbox(20, c['bg'], c['border_medium']),
        'cb_on':    lambda: _draw_checkbox(20, c['bg'], c['accent'], fill=c['accent'], check=c['accent_fg']),
        'radio_off': lambda: _draw_radio(20, c['bg'], c['border_medium']),
        'radio_on': lambda: _draw_radio(20, c['bg'], c['accent'], dot=c['accent']),
        'chevron':  lambda: _draw_chevron(18, c['ink_muted']),
    }
    out = {}
    for name, fn in specs.items():
        fp = d / '{}_{}_{}.png'.format(name, tid, _ASSET_VERSION)
        if not fp.exists():
            try:
                fn().save(str(fp), 'PNG')
            except Exception:
                return None
        out[name] = fp.as_posix()
    return out
