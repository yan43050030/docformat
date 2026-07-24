# -*- coding: utf-8 -*-
"""公文合规检查结果对话框（可交互修正）

列出每个文件的检查结果：
- 可自动修正的偏差 → 前置勾选框，用户勾选「认可」后才会被修改；
- 不可自动修正的偏差（如结构缺失）→ 只提示，标注需手动处理；
- 合格项 → ✓ 简要列出。
点击「应用所选修改并另存」后，仅对勾选项动手，原文件不动，结果另存。
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QFrame, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts.compliance import FIX_LABELS


class ComplianceReportDialog(QDialog):
    """exec_() 返回 Accepted 表示用户点了「应用所选修改」，且至少选了一项。"""

    def __init__(self, results, parent=None):
        super(ComplianceReportDialog, self).__init__(parent)
        self.setWindowTitle("公文合规检查结果")
        self.resize(720, 620)
        self._results = results
        # {result_index: {fix_key: QCheckBox}}
        self._boxes = {}
        self._fixable_total = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(8)

        tip = QLabel("检查标准来自你当前选中的预设。勾选你认可、希望自动修正的偏差，"
                     "其余保持不动——修正结果会另存为新文件，原文件不改。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        scroll.setWidget(host)
        body = QVBoxLayout(host)
        body.setContentsMargins(2, 2, 2, 2)
        body.setSpacing(12)

        for ri, res in enumerate(results):
            body.addWidget(self._build_file_block(ri, res))
        body.addStretch(1)
        root.addWidget(scroll, 1)

        # 底部操作行
        btns = QHBoxLayout()
        self.select_all = QCheckBox("全选可修正项")
        self.select_all.stateChanged.connect(self._toggle_all)
        if self._fixable_total == 0:
            self.select_all.setEnabled(False)
        btns.addWidget(self.select_all)
        btns.addStretch(1)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        self.apply_btn = QPushButton("应用所选修改并另存")
        self.apply_btn.setProperty("primary", "true")
        self.apply_btn.setCursor(Qt.PointingHandCursor)
        self.apply_btn.setEnabled(False)
        self.apply_btn.clicked.connect(self._on_apply)
        btns.addWidget(close_btn)
        btns.addWidget(self.apply_btn)
        root.addLayout(btns)

    def _build_file_block(self, ri, res):
        card = QFrame()
        card.setProperty("card", "true")
        v = QVBoxLayout(card)
        v.setContentsMargins(14, 12, 14, 12)
        v.setSpacing(6)

        findings = res.get('findings', [])
        warns = [f for f in findings if f['level'] == 'warn']
        head = QLabel('◆ {}'.format(res.get('display', '')))
        head.setProperty("sectionTitle", "true")
        v.addWidget(head)

        meta = '对照预设：{}　·　{}'.format(
            res.get('preset_name', '') or '当前预设',
            '存在 {} 项偏差'.format(len(warns)) if warns else '未发现偏差 ✓')
        meta_lbl = QLabel(meta)
        meta_lbl.setProperty("muted", "true")
        v.addWidget(meta_lbl)

        fixable = res.get('fix_input') is not None
        self._boxes[ri] = {}
        if not fixable and warns:
            note = QLabel("· 此文件非 .docx 格式，暂不支持自动修正，请先在 Word/WPS 里另存为 .docx")
            note.setProperty("muted", "true")
            note.setWordWrap(True)
            v.addWidget(note)

        for f in findings:
            level = f['level']
            fix_key = f.get('fix_key')
            text = '【{}】{}'.format(f['item'], f['detail'])
            if level == 'warn' and fix_key and fixable:
                cb = QCheckBox('✗ {}　→ 可自动{}'.format(text, FIX_LABELS.get(fix_key, '修正')))
                cb.setStyleSheet("")  # 保持主题默认
                cb.stateChanged.connect(self._refresh_apply)
                self._boxes[ri][fix_key] = cb
                self._fixable_total += 1
                v.addWidget(cb)
            elif level == 'warn':
                row = QLabel('✗ {}{}'.format(text, '　（需手动处理）' if not fix_key else ''))
                row.setWordWrap(True)
                v.addWidget(row)
            else:
                mark = '✓' if level == 'ok' else '·'
                row = QLabel('{} {}'.format(mark, text))
                row.setProperty("muted", "true")
                row.setWordWrap(True)
                v.addWidget(row)

        return card

    def _toggle_all(self, state):
        on = state == Qt.Checked
        for keys in self._boxes.values():
            for cb in keys.values():
                cb.blockSignals(True)
                cb.setChecked(on)
                cb.blockSignals(False)
        self._refresh_apply()

    def _refresh_apply(self, *_a):
        self.apply_btn.setEnabled(any(
            cb.isChecked() for keys in self._boxes.values() for cb in keys.values()))

    def _on_apply(self):
        if any(cb.isChecked() for keys in self._boxes.values() for cb in keys.values()):
            self.accept()

    def selections(self):
        """返回 [{'fix_input','display','preset','fix_keys':[...]}]，仅含有勾选的文件。"""
        out = []
        for ri, res in enumerate(self._results):
            keys = [k for k, cb in self._boxes.get(ri, {}).items() if cb.isChecked()]
            if keys and res.get('fix_input'):
                out.append({
                    'fix_input': res['fix_input'],
                    'display': res.get('display', ''),
                    'preset': res.get('preset', {}),
                    'fix_keys': keys,
                })
        return out
