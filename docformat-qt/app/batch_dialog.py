# -*- coding: utf-8 -*-
"""批量套打：一张表 → 一叠套打件。

表头就是字段名，一行一份。选完表先把"对上了哪些字段、哪些列多余、模板里
哪些字段没人填"摆出来给人看——批量最怕闷头跑完才发现全填错了。
"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog, QDialogButtonBox,
                             QFileDialog, QFormLayout, QHBoxLayout, QLabel,
                             QLineEdit, QMessageBox, QProgressBar, QPushButton,
                             QTextBrowser, QVBoxLayout)

from scripts import batch_fill


class BatchDialog(QDialog):
    def __init__(self, template_path, parent=None):
        super(BatchDialog, self).__init__(parent)
        self.setWindowTitle("批量套打 — 一张表打一叠")
        self.resize(680, 560)
        self._tpl = template_path
        self._rows = []
        self._header = []
        self.made = []

        root = QVBoxLayout(self)
        tip = QLabel(
            "把要打的内容整理成一张表：<b>第一行是字段名</b>（和模板里的字段对上就行），"
            "往下一行一份。支持 .xlsx / .csv / .tsv。<br>"
            "表里多出来的列（序号、备注之类）会自动忽略；模板里没人填的字段留空打印。")
        tip.setWordWrap(True)
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        row = QHBoxLayout()
        row.addWidget(QLabel("数据表："))
        self.path_label = QLabel("（未选择）")
        self.path_label.setWordWrap(True)
        row.addWidget(self.path_label, 1)
        b = QPushButton("选择…")
        b.clicked.connect(self._pick)
        row.addWidget(b)
        root.addLayout(row)

        self.report = QTextBrowser()
        self.report.setMinimumHeight(180)
        root.addWidget(self.report, 1)

        form = QFormLayout()
        self.name_combo = QComboBox()
        self.name_combo.setToolTip("拿哪一列当文件名。重名会自动加序号，不会互相覆盖")
        form.addRow("文件名取自：", self.name_combo)
        self.prefix = QLineEdit()
        self.prefix.setPlaceholderText("如：送审单_（可留空）")
        form.addRow("文件名前缀：", self.prefix)
        out_row = QHBoxLayout()
        self.out_label = QLabel("（未选择）")
        self.out_label.setWordWrap(True)
        out_row.addWidget(self.out_label, 1)
        b2 = QPushButton("选择…")
        b2.clicked.connect(self._pick_out)
        out_row.addWidget(b2)
        form.addRow("输出到：", out_row)
        root.addLayout(form)

        self.bar = QProgressBar()
        self.bar.setVisible(False)
        root.addWidget(self.bar)

        bb = QDialogButtonBox()
        bb.addButton("关闭", QDialogButtonBox.RejectRole)
        self.btn_go = bb.addButton("开始生成", QDialogButtonBox.AcceptRole)
        self.btn_go.setEnabled(False)
        bb.accepted.connect(self._run)
        bb.rejected.connect(self.reject)
        root.addWidget(bb)
        self._out_dir = ''

    # ---------- 选表 ----------
    def _pick(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择数据表", "", "表格 (*.xlsx *.csv *.tsv *.txt)")
        if not path:
            return
        try:
            self._header, self._rows = batch_fill.read_table(path)
        except Exception as exc:
            QMessageBox.warning(self, "这张表读不了", str(exc))
            return
        self.path_label.setText(path)
        if not self._out_dir:
            self._out_dir = os.path.join(os.path.dirname(path), '套打输出')
            self.out_label.setText(self._out_dir)
        matched, extra, missing = batch_fill.plan_batch(self._tpl, self._header)
        html = ['<div style="font-family:SimSun">',
                '共 <b>{}</b> 行数据。'.format(len(self._rows))]
        html.append('<p style="color:#1F7A4D">对上的字段（{}）：{}</p>'
                    .format(len(matched), '、'.join(matched) or '一个都没对上'))
        if extra:
            html.append('<p style="color:#888">表里多出来、会被忽略的列：{}</p>'
                        .format('、'.join(extra)))
        if missing:
            html.append('<p style="color:#B8860B">模板里这些字段表中没有，'
                        '打出来是空白：{}</p>'.format('、'.join(missing)))
        if not matched:
            html.append('<p style="color:#C0392B"><b>一个字段都没对上</b>——'
                        '请确认第一行是字段名，且和模板里的名字一致。</p>')
        if self._rows:
            html.append('<p>第一行预览：{}</p>'.format(
                '；'.join('{}={}'.format(k, v)
                          for k, v in list(self._rows[0].items())[:6])))
        html.append('</div>')
        self.report.setHtml(''.join(html))

        self.name_combo.clear()
        self.name_combo.addItem('（用行号）', '')
        for h in self._header:
            if h:
                self.name_combo.addItem(h, h)
        for pref in ('标题', '文号', '名称'):
            i = self.name_combo.findData(pref)
            if i >= 0:
                self.name_combo.setCurrentIndex(i)
                break
        self.btn_go.setEnabled(bool(matched and self._rows))

    def _pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self._out_dir or '')
        if d:
            self._out_dir = d
            self.out_label.setText(d)

    # ---------- 生成 ----------
    def _run(self):
        if not self._rows or not self._out_dir:
            QMessageBox.information(self, "还差一步", "请先选好数据表和输出文件夹。")
            return
        self.bar.setVisible(True)
        self.bar.setRange(0, len(self._rows))
        self.btn_go.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)

        def _tick(i, _n):
            self.bar.setValue(i)
            QApplication.processEvents()

        try:
            made, failed = batch_fill.batch_fill(
                self._tpl, self._rows, self._out_dir,
                name_field=self.name_combo.currentData() or None,
                prefix=self.prefix.text().strip(), progress=_tick)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            self.btn_go.setEnabled(True)
            QMessageBox.warning(self, "批量生成失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
            self.bar.setVisible(False)
        self.made = [p for p, _n in made]
        self.btn_go.setEnabled(True)

        msg = "生成 {} 份，存放在：\n{}".format(len(made), self._out_dir)
        if failed:
            msg += "\n\n有 {} 行没能生成：\n".format(len(failed))
            msg += '\n'.join('  第 {} 行 — {}'.format(i, why)
                             for i, why in failed[:8])
        notes = {n for _p, ns in made for n in (ns or [])}
        if notes:
            msg += "\n\n生成过程中的提示：\n" + '\n'.join(
                '  · ' + n for n in list(notes)[:6])
        QMessageBox.information(self, "批量生成完成", msg)
        from PyQt5.QtCore import QUrl
        from PyQt5.QtGui import QDesktopServices
        QDesktopServices.openUrl(QUrl.fromLocalFile(self._out_dir))
