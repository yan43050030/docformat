#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DocFormat Pro (Qt) — 公文格式自动排版工具

排版引擎复用自 docformat-gui (MIT License, Copyright KaguraNanaga)
https://github.com/KaguraNanaga/docformat-gui
"""
import os
import sys
import traceback

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


def _setup_crash_handler():
    """安装全局异常钩子：任何未捕获异常都弹出诊断报告窗口"""
    # 保存原始 excepthook
    _original_excepthook = sys.excepthook

    def _global_excepthook(exc_type, exc_value, exc_tb):
        # 忽略 KeyboardInterrupt（用户主动 Ctrl+C）
        if exc_type is KeyboardInterrupt:
            sys.exit(1)

        # 打印到 stderr 供终端可见
        traceback.print_exception(exc_type, exc_value, exc_tb)

        # 尝试弹出诊断对话框
        try:
            from app.diagnostic import show_crash_dialog
            show_crash_dialog(exc_type, exc_value, exc_tb)
        except Exception:
            pass

        # 调用原始 hook
        _original_excepthook(exc_type, exc_value, exc_tb)

    sys.excepthook = _global_excepthook


_setup_crash_handler()


def _print_version_and_exit():
    """--version / -v：不开界面就报个版本号。

    必须放在 **import PyQt5 之前**。装机、远程支持、打包后的冒烟验证都靠
    它——尤其是绿色版，"到底能不能在这台机器上跑起来"用它一句话就能验，
    前提是这条路径压根不碰 Qt：一旦碰了，没有 X 的环境（服务器、SSH、
    CI）就会先撞上 "Could not connect to any X display" 而不是打出版本号。
    版本号来自 app.version —— 那个模块只有一行常量、不 import 任何东西，
    正是为这条路径准备的（app.main_window 会顺带拉起整个界面模块）。
    """
    if len(sys.argv) > 1 and sys.argv[1] in ('--version', '-v'):
        from app.version import VERSION
        print('DocFormat Pro {}'.format(VERSION))
        raise SystemExit(0)


_print_version_and_exit()

from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QFont, QGuiApplication, QIcon
from PyQt5.QtWidgets import QApplication

from app.main_window import MainWindow


def resource_path(name):
    base = getattr(sys, '_MEIPASS', APP_DIR)
    return os.path.join(base, name)


def main():
    # HiDPI 支持（国产整机 4K 屏 / Windows 125%、150% 缩放）
    # PassThrough 让 1.25x/1.5x 等非整数缩放按真实比例渲染，避免模糊或过大
    if hasattr(Qt, 'HighDpiScaleFactorRoundingPolicy'):
        QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    # 去掉对话框标题栏无用的 "?" 帮助按钮（Qt 5.10+；旧版 Qt 无此属性则跳过）
    if hasattr(Qt, 'AA_DisableWindowContextHelpButton'):
        QCoreApplication.setAttribute(Qt.AA_DisableWindowContextHelpButton, True)

    app = QApplication(sys.argv)
    app.setApplicationName("DocFormat Pro")
    app.setOrganizationName("DocFormatPro")

    icon_file = resource_path(os.path.join('assets', 'icon.ico'))
    if os.path.exists(icon_file):
        app.setWindowIcon(QIcon(icon_file))

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
