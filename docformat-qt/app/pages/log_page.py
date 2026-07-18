# -*- coding: utf-8 -*-
"""日志页：等级着色的处理日志"""
import time

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QFileDialog, QHBoxLayout, QLabel, QMessageBox,
                             QPushButton, QTextEdit, QVBoxLayout, QWidget)

LEVEL_COLORS = {
    'info': '#5C7CFA',
    'success': '#2E7D32',
    'warning': '#B26A00',
    'error': '#C62828',
}
LEVEL_LABELS = {'info': '信息', 'success': '成功', 'warning': '警告', 'error': '错误'}


class LogPage(QWidget):
    def __init__(self, parent=None):
        super(LogPage, self).__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 16)
        root.setSpacing(12)

        head = QHBoxLayout()
        title = QLabel("处理日志")
        title.setProperty("h1", "true")
        export_btn = QPushButton("导出日志")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.clicked.connect(self.export_log)
        clear_btn = QPushButton("清空日志")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self.clear)
        head.addWidget(title)
        head.addStretch(1)
        head.addWidget(export_btn)
        head.addWidget(clear_btn)
        root.addLayout(head)

        self.view = QTextEdit()
        self.view.setObjectName("LogView")
        self.view.setReadOnly(True)
        # 上限防止长时间批量运行内存无限增长
        self.view.document().setMaximumBlockCount(5000)
        root.addWidget(self.view, 1)

    def append(self, level, message):
        color = LEVEL_COLORS.get(level, '#888888')
        label = LEVEL_LABELS.get(level, level)
        ts = time.strftime('%H:%M:%S')
        safe = (message.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;').replace('\n', '<br>'))
        self.view.append(
            '<span style="color:#8A8375">[{}]</span> '
            '<b style="color:{}">[{}]</b> {}'.format(ts, color, label, safe))

    def clear(self):
        self.view.clear()

    def export_log(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "导出日志", time.strftime('处理日志_%Y%m%d_%H%M%S.txt'), "文本文件 (*.txt)")
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.view.toPlainText())
            QMessageBox.information(self, "导出成功", "日志已保存到：\n" + path)
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))
