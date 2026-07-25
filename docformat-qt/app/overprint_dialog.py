# -*- coding: utf-8 -*-
"""套打填写：选套打模板 → 填字段 → 生成可直接打到预印纸上的文件"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QComboBox, QDialog, QFileDialog, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                             QPlainTextEdit, QPushButton, QScrollArea,
                             QVBoxLayout, QWidget)

from app.theme import settings
from scripts import overprint

# 这些字段内容通常较长，用多行输入框
_LONG_FIELDS = {'拟办意见', '领导批示', '备注', '主要内容', '说明'}
# 这些字段每次都变，不做记忆
_NO_MEMORY = {'标题', '拟办意见', '领导批示', '备注', '年', '月', '日'}


class OverprintDialog(QDialog):
    def __init__(self, parent=None):
        super(OverprintDialog, self).__init__(parent)
        self.setWindowTitle("套打填写 — 打到预印红头纸上")
        self.resize(680, 640)
        self._editors = {}
        self._template_path = None

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(8)

        tip = QLabel("套打模板里预印在纸上的红色内容是白色文字（占位不显影），"
                     "只有你填的内容会被打印出来，位置与预印栏位严格对齐。"
                     "内容过长时会自动缩小字号以免撑高表格、导致错位。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        row = QHBoxLayout()
        row.addWidget(QLabel("套打模板："))
        self.tpl_combo = QComboBox()
        self._templates = overprint.list_templates()
        for name, path, builtin in self._templates:
            self.tpl_combo.addItem('{}{}'.format(name, '（自带）' if builtin else ''), path)
        self.tpl_combo.currentIndexChanged.connect(self._load_fields)
        row.addWidget(self.tpl_combo, 1)
        add_btn = QPushButton("添加模板…")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setToolTip("导入自己的套打模板 docx（含 {{字段名}} 占位符）")
        add_btn.clicked.connect(self._import_template)
        row.addWidget(add_btn)
        root.addLayout(row)

        src_row = QHBoxLayout()
        pick = QPushButton("从 docx 导入内容…")
        pick.setCursor(Qt.PointingHandCursor)
        pick.setToolTip("选一份已有的送审单/草稿 docx，自动识别各部分内容填进下面的字段；\n"
                        "日期会拆成年/月/日分别落位，识别不到的可手工补填")
        pick.clicked.connect(self._import_content)
        src_row.addWidget(pick)
        clear = QPushButton("清空")
        clear.setProperty("flat", "true")
        clear.setCursor(Qt.PointingHandCursor)
        clear.clicked.connect(self._clear_fields)
        src_row.addWidget(clear)
        src_row.addStretch(1)
        root.addLayout(src_row)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        root.addWidget(self.scroll, 1)

        self.status = QLabel("")
        self.status.setProperty("muted", "true")
        self.status.setWordWrap(True)
        root.addWidget(self.status)

        btns = QHBoxLayout()
        btns.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("生成套打文件")
        ok.setProperty("primary", "true")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self._generate)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        root.addLayout(btns)

        if not self._templates:
            self.status.setText("未找到套打模板。点「添加模板…」导入一份含 {{字段名}} 的 docx。")
        else:
            self._load_fields()

    # ---------- 字段 ----------
    def _load_fields(self, *_a):
        path = self.tpl_combo.currentData()
        self._template_path = path
        host = QWidget()
        form = QFormLayout(host)
        form.setContentsMargins(4, 6, 4, 6)
        form.setSpacing(8)
        self._editors = {}
        if not path or not os.path.exists(path):
            self.scroll.setWidget(host)
            return
        try:
            fields = overprint.scan_fields(path)
        except Exception as e:
            self.status.setText("读取模板失败：{}".format(e))
            self.scroll.setWidget(host)
            return
        s = settings()
        for name in fields:
            if name in _LONG_FIELDS:
                ed = QPlainTextEdit()
                ed.setMinimumHeight(110)
                ed.setPlaceholderText("内容过长会自动缩小字号，仍放不下会提示")
            else:
                ed = QLineEdit()
                if name not in _NO_MEMORY:
                    ed.setText(s.value('overprint/{}'.format(name), '') or '')
            self._editors[name] = ed
            form.addRow(name + '：', ed)
        self.scroll.setWidget(host)
        self.status.setText("共 {} 个可填字段；留空的字段打印出来就是空白。".format(len(fields)))

    def _values(self):
        out = {}
        for name, ed in self._editors.items():
            out[name] = (ed.toPlainText() if isinstance(ed, QPlainTextEdit)
                         else ed.text()).strip()
        return out

    def _clear_fields(self):
        for ed in self._editors.values():
            if isinstance(ed, QPlainTextEdit):
                ed.setPlainText('')
            else:
                ed.setText('')

    def _import_content(self):
        """从已有 docx 抽取内容填进字段（日期自动拆成年/月/日）"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先选择套打模板")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择要适配的文档", "", "Word 文档 (*.docx);;所有文件 (*.*)")
        if not path:
            return
        from PyQt5.QtWidgets import QApplication
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            values = overprint.extract_values(path, list(self._editors.keys()))
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "读取失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()
        for name, val in values.items():
            ed = self._editors.get(name)
            if ed is None:
                continue
            if isinstance(ed, QPlainTextEdit):
                ed.setPlainText(val)
            else:
                ed.setText(val)
        missing = [n for n in self._editors if not values.get(n)]
        msg = "已识别 {} 个字段".format(len(values))
        if missing:
            msg += "；未识别：{}（请手工补填）".format('、'.join(missing))
        self.status.setText(msg)

    def _import_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择套打模板", "", "Word 文档 (*.docx)")
        if not path:
            return
        try:
            fields = overprint.scan_fields(path)
        except Exception as e:
            QMessageBox.warning(self, "读取失败", str(e))
            return
        if not fields:
            QMessageBox.information(
                self, "没有可填字段",
                "这份 docx 里没找到 {{字段名}} 占位符。\n\n"
                "套打模板的做法：把纸上已预印的内容设为白色文字（占位不显影），"
                "要填的位置写成 {{字段名}}。")
            return
        import shutil
        d = overprint.user_overprint_dir()
        os.makedirs(d, exist_ok=True)
        dest = os.path.join(d, os.path.basename(path))
        try:
            shutil.copyfile(path, dest)
        except Exception as e:
            QMessageBox.warning(self, "导入失败", str(e))
            return
        self._templates = overprint.list_templates()
        self.tpl_combo.blockSignals(True)
        self.tpl_combo.clear()
        for name, p, builtin in self._templates:
            self.tpl_combo.addItem('{}{}'.format(name, '（自带）' if builtin else ''), p)
        idx = self.tpl_combo.findData(dest)
        self.tpl_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.tpl_combo.blockSignals(False)
        self._load_fields()

    # ---------- 生成 ----------
    def _generate(self):
        if not self._template_path:
            return
        values = self._values()
        if not any(values.values()):
            QMessageBox.information(self, "提示", "请至少填写一个字段")
            return
        stem = os.path.splitext(os.path.basename(self._template_path))[0]
        title = values.get('标题', '').strip()
        default = '{}{}.docx'.format(stem, '_' + title[:16] if title else '')
        out, _ = QFileDialog.getSaveFileName(self, "保存套打文件", default,
                                             "Word 文档 (*.docx)")
        if not out:
            return
        from PyQt5.QtWidgets import QApplication
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            n, notes = overprint.fill_form(self._template_path, values, out)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "生成失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()

        s = settings()
        for name, val in values.items():
            if name not in _NO_MEMORY and val:
                s.setValue('overprint/{}'.format(name), val)

        self.result_path = out
        msg = "已生成：\n{}\n\n共填入 {} 个字段。".format(out, n)
        if notes:
            msg += "\n\n注意：\n" + "\n".join('· ' + x for x in notes)
        msg += "\n\n打印时请用预印红头纸，并确认打印机「按实际大小/100%」不缩放。"
        QMessageBox.information(self, "生成完成", msg)
        self.accept()
