# -*- coding: utf-8 -*-
"""文件拖拽区：Qt 原生拖拽（Windows / 麒麟 X11 均可靠）"""
import os

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QFrame, QLabel, QVBoxLayout

ALLOWED_EXTS = ('.docx', '.doc', '.wps', '.txt', '.md')


def filter_paths(urls):
    """从 QMimeData urls 提取合法的本地文档路径（含文件夹展开）"""
    paths = []
    for url in urls:
        p = url.toLocalFile()
        if not p:
            continue
        if os.path.isdir(p):
            for root, _dirs, files in os.walk(p):
                for fn in files:
                    if fn.lower().endswith(ALLOWED_EXTS) and not fn.startswith('~$'):
                        paths.append(os.path.join(root, fn))
        elif p.lower().endswith(ALLOWED_EXTS) and not os.path.basename(p).startswith('~$'):
            paths.append(p)
    return paths


class DropZone(QFrame):
    filesDropped = pyqtSignal(list)
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super(DropZone, self).__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(150)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 28, 16, 28)
        lay.setSpacing(6)
        lay.setAlignment(Qt.AlignCenter)

        self.icon = QLabel()
        self.icon.setAlignment(Qt.AlignCenter)
        self.icon.setStyleSheet("background: transparent;")
        self.set_icon_color('#8A8375')   # 启动默认色，主题应用后会重设
        self.title = QLabel("拖拽文件到此处，或点击选择")
        self.title.setObjectName("DropTitle")
        self.title.setAlignment(Qt.AlignCenter)
        self.hint = QLabel("支持 .docx / .doc / .wps / .txt / .md 格式，可多选文件或拖入文件夹")
        self.hint.setObjectName("DropHint")
        self.hint.setAlignment(Qt.AlignCenter)

        lay.addWidget(self.icon)
        lay.addWidget(self.title)
        lay.addWidget(self.hint)

    def set_icon_color(self, color):
        """用程序绘制的图标替代字体字符（信创系统 emoji 字体不全）"""
        from app.widgets.nav_icons import make_pixmap
        self.icon.setPixmap(make_pixmap('drop', color, 44))

    def _set_drag_over(self, on):
        self.setProperty("dragOver", "true" if on else "false")
        self.title.setText("释放文件以添加" if on else "拖拽文件到此处，或点击选择")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and filter_paths(event.mimeData().urls()):
            event.acceptProposedAction()
            self._set_drag_over(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drag_over(False)

    def dropEvent(self, event):
        self._set_drag_over(False)
        paths = filter_paths(event.mimeData().urls())
        if paths:
            self.filesDropped.emit(paths)
            event.acceptProposedAction()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
