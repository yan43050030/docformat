# -*- coding: utf-8 -*-
"""格式诊断报告对话框：等宽正文 + 导出 TXT + 一键转入修复"""
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (QDialog, QFileDialog, QHBoxLayout, QLabel,
                             QMessageBox, QPlainTextEdit, QPushButton,
                             QVBoxLayout)


class ReportDialog(QDialog):
    """exec_() 返回 Accepted 表示用户点击了「立即一键修复」"""

    def __init__(self, report, parent=None):
        super(ReportDialog, self).__init__(parent)
        self.setWindowTitle("格式诊断报告")
        self.resize(760, 560)
        self.report = report

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(8)

        tip = QLabel("以下为诊断结果（未修改任何文件），日志页可回看：")
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        self.view = QPlainTextEdit()
        self.view.setReadOnly(True)
        self.view.setPlainText(report)
        mono = QFont("Courier New")
        mono.setStyleHint(QFont.Monospace)
        self.view.setFont(mono)
        root.addWidget(self.view, 1)

        btns = QHBoxLayout()
        export_btn = QPushButton("导出 TXT")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.clicked.connect(self._export)
        btns.addWidget(export_btn)
        btns.addStretch(1)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        fix_btn = QPushButton("立即一键修复")
        fix_btn.setProperty("primary", "true")
        fix_btn.setCursor(Qt.PointingHandCursor)
        fix_btn.setToolTip("切换到「智能一键处理」模式并立即处理这批文件")
        fix_btn.clicked.connect(self.accept)
        btns.addWidget(close_btn)
        btns.addWidget(fix_btn)
        root.addLayout(btns)

    def _export(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出诊断报告", "诊断报告.txt", "文本文件 (*.txt)")
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.report)
            from app.widgets.toast import Toast
            Toast.show_message(self, "报告已导出", "success")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))
