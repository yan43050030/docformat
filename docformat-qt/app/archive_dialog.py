# -*- coding: utf-8 -*-
"""归档命名与登记表：改名前先摆出来看，确认了才动手。"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QAbstractItemView, QApplication, QCheckBox,
                             QComboBox, QDialog, QFileDialog, QHBoxLayout,
                             QHeaderView, QLabel, QLineEdit, QMessageBox,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QVBoxLayout)

from app.theme import settings
from scripts import archive as A

# 几个常见的命名式，直接选，不必自己拼占位符
PRESET_PATTERNS = [
    ('日期-文号-标题', '{成文日期}-{文号}-{标题}'),
    ('文号-标题', '{文号}-{标题}'),
    ('年月-文种-标题', '{年}{月}-{文种}-{标题}'),
    ('密级-文号-标题', '{密级词}-{文号}-{标题}'),
    ('日期-标题', '{成文日期}-{标题}'),
]


class ArchiveDialog(QDialog):
    """exec_() 返回 Accepted 表示已归档；outputs 是归档后的文件路径。"""

    def __init__(self, paths, preset=None, parent=None):
        super(ArchiveDialog, self).__init__(parent)
        self.setWindowTitle("归档命名与登记表")
        self.resize(940, 620)
        self._paths = list(paths)
        self._preset = preset
        self._items = []
        self.outputs = []
        s = settings()

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(8)

        tip = QLabel(
            "文号、标题、成文日期、密级这些字，软件排版时已经认过一遍了，"
            "不必再抄一次。下面按命名式生成归档文件名，逐行可改；确认后"
            "<b>复制</b>到归档目录（原件不动），并把每份的信息追加到登记台账。")
        tip.setWordWrap(True)
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        # ---- 命名式 ----
        row = QHBoxLayout()
        row.addWidget(QLabel("命名式："))
        self.pat_combo = QComboBox()
        for label, pat in PRESET_PATTERNS:
            self.pat_combo.addItem(label, pat)
        self.pat_combo.addItem('自定义', None)
        self.pat_combo.currentIndexChanged.connect(self._on_pattern_pick)
        row.addWidget(self.pat_combo)
        self.pat_edit = QLineEdit(
            s.value('archive/pattern', A.DEFAULT_PATTERN, type=str))
        self.pat_edit.setToolTip(
            "可用的字段：\n" + '\n'.join(
                '  {{{}}}  {}'.format(k, d) for k, d, _f in A.FIELDS))
        self.pat_edit.textChanged.connect(self._refresh)
        row.addWidget(self.pat_edit, 1)
        root.addLayout(row)

        fields = QLabel('可用字段：' + '　'.join(
            '{{{}}}'.format(k) for k in A.FIELD_KEYS))
        fields.setProperty("muted", "true")
        fields.setWordWrap(True)
        root.addWidget(fields)

        # ---- 预览表 ----
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ['原文件名', '识别到的文号 / 日期 / 密级', '归档文件名（可改）', ''])
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked
                                   | QAbstractItemView.SelectedClicked)
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        root.addWidget(self.table, 1)

        # ---- 归档目录 ----
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("归档到："))
        self.dir_edit = QLineEdit(s.value('archive/dir', '', type=str))
        self.dir_edit.setPlaceholderText("选择存放归档件的文件夹")
        row2.addWidget(self.dir_edit, 1)
        b = QPushButton("浏览…")
        b.clicked.connect(self._pick_dir)
        row2.addWidget(b)
        self.move_check = QCheckBox("移动（不保留原件）")
        self.move_check.setToolTip(
            "默认是复制，原件留在原地。勾上就是移动——原件不再存在，请想清楚。")
        row2.addWidget(self.move_check)
        root.addLayout(row2)

        # ---- 台账 ----
        row3 = QHBoxLayout()
        self.led_check = QCheckBox("同时登记到台账：")
        self.led_check.setChecked(s.value('archive/ledger_on', True, type=bool))
        row3.addWidget(self.led_check)
        self.led_edit = QLineEdit(s.value('archive/ledger', '', type=str))
        self.led_edit.setPlaceholderText("登记台账 .csv（不存在会自动新建）")
        row3.addWidget(self.led_edit, 1)
        b2 = QPushButton("浏览…")
        b2.clicked.connect(self._pick_ledger)
        row3.addWidget(b2)
        root.addLayout(row3)

        warn = QLabel(
            "台账里存的是<b>明文</b>的标题、文号、密级——那正是台账的用处，"
            "但也意味着台账本身要按其中<b>最高密级</b>管理。存放位置请自行斟酌。")
        warn.setWordWrap(True)
        warn.setProperty("muted", "true")
        root.addWidget(warn)

        self.status = QLabel("")
        self.status.setProperty("muted", "true")
        self.status.setWordWrap(True)
        root.addWidget(self.status)

        btns = QHBoxLayout()
        btns.addStretch(1)
        c = QPushButton("取消")
        c.clicked.connect(self.reject)
        btns.addWidget(c)
        self.ok_btn = QPushButton("归档")
        self.ok_btn.setProperty("primary", "true")
        self.ok_btn.setCursor(Qt.PointingHandCursor)
        self.ok_btn.clicked.connect(self._run)
        btns.addWidget(self.ok_btn)
        root.addLayout(btns)

        self._sync_combo()
        self._reload()

    # ---------- 数据 ----------
    def _reload(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self._items = A.plan(self._paths, self.pat_edit.text(), self._preset)
        finally:
            QApplication.restoreOverrideCursor()
        self._fill()

    def _refresh(self):
        """只重算文件名，不重新读文档——读一次就够了，改命名式时不该再等一遍"""
        pat = self.pat_edit.text()
        for it in self._items:
            ext = os.path.splitext(it['src'])[1] or '.docx'
            it['new_name'] = A.render_name(pat, it['meta'], ext)
        self._sync_combo()
        self._fill()

    def _sync_combo(self):
        pat = self.pat_edit.text()
        idx = self.pat_combo.findData(pat)
        self.pat_combo.blockSignals(True)
        self.pat_combo.setCurrentIndex(idx if idx >= 0
                                       else self.pat_combo.count() - 1)
        self.pat_combo.blockSignals(False)

    def _on_pattern_pick(self):
        pat = self.pat_combo.currentData()
        if pat:
            self.pat_edit.setText(pat)

    def _fill(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self._items))
        weak = 0
        for r, it in enumerate(self._items):
            m = it['meta']
            self.table.setItem(r, 0, _ro(os.path.basename(it['src'])))
            got = '　'.join(x for x in (m.get('文号'), m.get('成文日期'),
                                       m.get('密级')) if x) or '（都没认出来）'
            cell = _ro(got)
            if not m.get('文号') and not m.get('成文日期'):
                weak += 1
                cell.setToolTip('文号和成文日期都没认出来，归档名只能靠标题')
            self.table.setItem(r, 1, cell)
            self.table.setItem(r, 2, QTableWidgetItem(it['new_name']))
            self.table.setItem(r, 3, _ro(''))
        self.table.blockSignals(False)
        msg = '共 {} 份'.format(len(self._items))
        if weak:
            msg += ('；其中 {} 份没认出文号和日期，归档名可能不理想，'
                    '可在表里直接改'.format(weak))
        # 同名撞车先说在前头，别等归档完了才发现多了一堆 (2)(3)
        names = [it['new_name'] for it in self._items]
        dup = sorted({n for n in names if names.count(n) > 1})
        if dup:
            msg += '；有 {} 组重名，归档时会自动加 (2)(3) 区分'.format(len(dup))
        self.status.setText(msg)

    def _collect(self):
        out = []
        for r, it in enumerate(self._items):
            cell = self.table.item(r, 2)
            name = (cell.text() if cell else '').strip() or it['new_name']
            if not os.path.splitext(name)[1]:
                name += os.path.splitext(it['src'])[1] or '.docx'
            out.append(dict(it, new_name=A.safe_stem(
                os.path.splitext(name)[0]) + os.path.splitext(name)[1]))
        return out

    # ---------- 交互 ----------
    def _pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择归档目录",
                                             self.dir_edit.text() or '')
        if d:
            self.dir_edit.setText(d)

    def _pick_ledger(self):
        start = self.led_edit.text() or os.path.join(
            self.dir_edit.text() or '', '公文登记台账.csv')
        p, _ = QFileDialog.getSaveFileName(
            self, "选择登记台账", start, "CSV 表格 (*.csv)")
        if p:
            self.led_edit.setText(p)
            self.led_check.setChecked(True)

    def _run(self):
        out_dir = self.dir_edit.text().strip()
        if not out_dir:
            QMessageBox.information(self, "提示", "请先选择归档目录")
            return
        led = self.led_edit.text().strip() if self.led_check.isChecked() else ''
        if self.led_check.isChecked() and not led:
            QMessageBox.information(self, "提示", "请选择登记台账文件，或取消勾选")
            return
        move = self.move_check.isChecked()
        if move:
            ret = QMessageBox.question(
                self, "移动原件",
                "勾了「移动」——归档之后原件就不在原来的位置了。\n\n"
                "确定要移动，而不是复制？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return

        items = self._collect()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            done, failed = A.archive(items, out_dir, move=move,
                                     ledger_path=led or None)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "归档失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        s = settings()
        s.setValue('archive/pattern', self.pat_edit.text())
        s.setValue('archive/dir', out_dir)
        s.setValue('archive/ledger', led)
        s.setValue('archive/ledger_on', self.led_check.isChecked())

        self.outputs = [dst for _src, dst in done]
        msg = '{} 份已{}到：\n{}'.format(
            len(done), '移动' if move else '复制', out_dir)
        if led and done:
            msg += '\n\n已登记 {} 行到台账：\n{}'.format(len(done), led)
        if failed:
            msg += '\n\n以下没成功：\n' + '\n'.join(
                '· {}：{}'.format(os.path.basename(p), why) for p, why in failed)
        QMessageBox.information(self, "归档完成", msg)
        if done:
            self.accept()


def _ro(text):
    it = QTableWidgetItem(text)
    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    return it
