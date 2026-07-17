#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DocFormat Pro (Qt) — 公文格式自动排版工具

排版引擎复用自 docformat-gui (MIT License, Copyright KaguraNanaga)
https://github.com/KaguraNanaga/docformat-gui
"""
import os
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication

from app.main_window import MainWindow


def main():
    # HiDPI 支持（国产整机 4K 屏）
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("DocFormat Pro")
    app.setOrganizationName("DocFormatPro")

    # 中文字体回退链：Windows 用微软雅黑，麒麟/UOS 用系统默认无衬线
    font = QFont()
    if sys.platform == 'win32':
        font.setFamily("Microsoft YaHei UI")
    font.setPointSize(10)
    app.setFont(font)

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
