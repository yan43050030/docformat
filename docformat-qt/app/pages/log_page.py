# -*- coding: utf-8 -*-
"""日志页：等级着色的处理日志（含可选脱敏的持久化日志文件）"""
import os
import time

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (QCheckBox, QFileDialog, QHBoxLayout, QLabel,
                             QMessageBox, QPushButton, QTextEdit, QVBoxLayout,
                             QWidget)

from app import file_logger

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
        openfolder_btn = QPushButton("打开日志文件夹")
        openfolder_btn.setCursor(Qt.PointingHandCursor)
        openfolder_btn.clicked.connect(self._open_log_folder)
        clearfile_btn = QPushButton("清空日志文件")
        clearfile_btn.setCursor(Qt.PointingHandCursor)
        clearfile_btn.clicked.connect(self._clear_log_file)
        head.addWidget(title)
        head.addStretch(1)
        head.addWidget(openfolder_btn)
        head.addWidget(clearfile_btn)
        head.addWidget(export_btn)
        head.addWidget(clear_btn)
        root.addLayout(head)

        # 隐私控制行
        priv = QHBoxLayout()
        self.persist_chk = QCheckBox("保存日志到文件（方便回溯与反馈问题）")
        self.persist_chk.setChecked(file_logger.persist_enabled())
        self.persist_chk.stateChanged.connect(
            lambda: file_logger.set_persist(self.persist_chk.isChecked()))
        self.redact_chk = QCheckBox("脱敏保存（文件名/路径不明文写入，涉密推荐）")
        self.redact_chk.setChecked(file_logger.redact_enabled())
        self.redact_chk.stateChanged.connect(self._on_redact_toggle)
        priv.addWidget(self.persist_chk)
        priv.addWidget(self.redact_chk)
        priv.addStretch(1)
        root.addLayout(priv)

        hint = QLabel("日志文件保存在本地配置目录，不会上传。默认脱敏：文件名以哈希代替、"
                      "目录与用户名不写入。可随时「清空日志文件」。")
        hint.setProperty("muted", "true")
        hint.setWordWrap(True)
        root.addWidget(hint)

        self.view = QTextEdit()
        self.view.setObjectName("LogView")
        self.view.setReadOnly(True)
        # 上限防止长时间批量运行内存无限增长
        self.view.document().setMaximumBlockCount(5000)
        root.addWidget(self.view, 1)

    def append(self, level, message):
        # 屏幕日志显示完整信息（仅本机操作者可见，会话结束即消失）
        color = LEVEL_COLORS.get(level, '#888888')
        label = LEVEL_LABELS.get(level, level)
        ts = time.strftime('%H:%M:%S')
        safe = (message.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;').replace('\n', '<br>'))
        self.view.append(
            '<span style="color:#8A8375">[{}]</span> '
            '<b style="color:{}">[{}]</b> {}'.format(ts, color, label, safe))
        # 持久化日志（默认脱敏）
        file_logger.write(level, message)

    def _open_log_folder(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_logger.log_dir())))

    def _clear_log_file(self):
        ret = QMessageBox.question(self, "清空日志文件",
                                   "确定删除本地保存的日志文件吗？此操作不可恢复。",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            ok = file_logger.clear()
            QMessageBox.information(self, "已清空" if ok else "提示",
                                    "日志文件已删除" if ok else "没有可删除的日志文件")

    def _on_redact_toggle(self):
        if not self.redact_chk.isChecked():
            ret = QMessageBox.warning(
                self, "关闭脱敏",
                "关闭后，日志文件将明文记录文件名与路径（可能含涉密名称或人名）。\n"
                "仅建议在本机排查问题时临时关闭。确定关闭脱敏吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                self.redact_chk.setChecked(True)
                return
        file_logger.set_redact(self.redact_chk.isChecked())

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
