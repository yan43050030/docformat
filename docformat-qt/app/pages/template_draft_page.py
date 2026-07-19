# -*- coding: utf-8 -*-
"""
============================================================================
模块一：模板起草页 (TemplateDraftPage)
============================================================================

【用途】
给 docformat-qt 增加"起草"能力。选文书模板 → 填字段 → 生成初稿 → 排版引擎排版。

模板中的 {{占位符}} 会自动识别为表单字段：
  - 含"情况/说明/内容/简要/事实"等关键词的 → 多行文本区
  - 其他 → 单行输入框

【模板格式】见 template_common.py
============================================================================
"""
import os
import re
import shutil
import tempfile

from PyQt5.QtWidgets import (
    QWidget, QListWidget, QListWidgetItem, QLineEdit, QTextEdit,
    QPushButton, QLabel, QFormLayout, QVBoxLayout, QScrollArea,
    QMessageBox, QSplitter, QPlainTextEdit, QFileDialog, QHBoxLayout,
    QApplication, QComboBox, QInputDialog,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from app.template_common import (
    TEMPLATE_DIR, PLACEHOLDER_RE,
    extract_fields, parse_template, render, find_unfilled,
    load_template_dirs, save_template_dirs, scan_templates,
    bundled_templates_dir, is_bundled_dir,
)


def _build_plain_text(rendered):
    """将渲染后的结构化 dict 拼成纯文本"""
    lines = [rendered["title"], ""]
    lines += [b["text"] for b in rendered["body"]]
    for k, v in rendered["meta"].items():
        lines += ["", "{}: {}".format(k, v)]
    return "\n".join(lines)


class DraftFormatWorker(QThread):
    """起草文本 → 临时 docx → 标点修复 → 排版引擎 → 输出 .docx"""
    logMessage = pyqtSignal(str, str)
    finishedWith = pyqtSignal(bool, str)

    def __init__(self, text, out_path, preset_name, custom_settings, parent=None):
        super(DraftFormatWorker, self).__init__(parent)
        self.text = text
        self.out_path = out_path
        self.preset_name = preset_name
        self.custom_settings = custom_settings

    def run(self):
        tmp_root = tempfile.mkdtemp(prefix='docformat_draft_')
        try:
            from docx import Document
            tmp = os.path.join(tmp_root, 'draft.docx')
            doc = Document()
            for line in self.text.splitlines():
                doc.add_paragraph(line)
            doc.save(tmp)

            from scripts import punctuation
            from scripts.formatter import format_document
            tmp2 = os.path.join(tmp_root, 'punct.docx')
            punctuation.process_document(tmp, tmp2)
            format_document(tmp2, self.out_path,
                            preset_name=self.preset_name,
                            custom_settings=self.custom_settings)
            self.logMessage.emit('success', '已生成: {}'.format(self.out_path))
            self.finishedWith.emit(True, self.out_path)
        except PermissionError:
            msg = '无法写入输出文件，它可能正被 Word/WPS 打开，请关闭后重试'
            self.logMessage.emit('error', '生成失败: {}'.format(msg))
            self.finishedWith.emit(False, msg)
        except Exception as e:
            self.logMessage.emit('error', '生成失败: {}'.format(e))
            self.finishedWith.emit(False, str(e))
        finally:
            shutil.rmtree(tmp_root, ignore_errors=True)


class TemplateDraftPage(QWidget):
    def __init__(self, mgr, parent=None):
        super().__init__(parent)
        self.mgr = mgr
        self.current_template_text = ""
        self.current_template_path = ""
        self.field_inputs = {}
        self._all_templates = []  # [(display_name, full_path, source_dir), ...]
        self._build_ui()
        self._load_template_list()

    def _build_ui(self):
        splitter = QSplitter(Qt.Horizontal)

        # ===== 左侧：模板列表 + 搜索 + 目录管理 =====
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(8, 8, 4, 8)

        # 搜索框
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索模板（模糊匹配）...")
        self.search_edit.textChanged.connect(self._on_search)
        lv.addWidget(self.search_edit)

        # 模板列表
        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_select_template)
        lv.addWidget(self.list_widget)

        # 底部按钮行
        btn_row = QHBoxLayout()
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self._load_template_list)
        btn_row.addWidget(refresh_btn)

        dir_btn = QPushButton("目录管理")
        dir_btn.clicked.connect(self._on_manage_dirs)
        btn_row.addWidget(dir_btn)
        lv.addLayout(btn_row)

        hint = QLabel("搜索时自动从已配置目录中匹配模板")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888; font-size: 11px;")
        lv.addWidget(hint)
        splitter.addWidget(left)

        # ===== 中间：字段填写 =====
        mid = QWidget()
        mv = QVBoxLayout(mid)
        mv.setContentsMargins(4, 8, 4, 8)
        mv.addWidget(QLabel("填写字段（根据模板自动生成）"))
        self.form_scroll = QScrollArea()
        self.form_scroll.setWidgetResizable(True)
        self.form_container = QWidget()
        self.form_layout = QFormLayout(self.form_container)
        self.form_scroll.setWidget(self.form_container)
        mv.addWidget(self.form_scroll)
        self.gen_btn = QPushButton("生成初稿并排版")
        self.gen_btn.clicked.connect(self._on_generate)
        self.gen_btn.setStyleSheet("QPushButton { font-weight: bold; padding: 6px 20px; }")
        mv.addWidget(self.gen_btn)
        splitter.addWidget(mid)

        # ===== 右侧：预览 =====
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(4, 8, 8, 8)
        rv.addWidget(QLabel("正文预览（未填字段显示 {{字段名}}）"))
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        rv.addWidget(self.preview)
        splitter.addWidget(right)

        splitter.setSizes([240, 360, 420])
        outer = QVBoxLayout(self)
        outer.addWidget(splitter)

    # ---------- 模板列表 & 搜索 ----------
    def _load_template_list(self):
        """重新扫描所有目录，更新模板列表"""
        self._all_templates = scan_templates()
        self._refresh_list_display()

    def _on_search(self, query):
        self._refresh_list_display(query)

    def _refresh_list_display(self, query=None):
        """根据搜索词过滤并刷新列表显示"""
        self.list_widget.clear()
        q = (query or "").strip()
        for display, path, src_dir in self._all_templates:
            if q and q not in display:
                continue
            # 列表中显示模板名 + 来源目录
            home = os.path.expanduser("~")
            src_label = src_dir
            if src_label.startswith(home):
                src_label = "~" + src_label[len(home):]
            label = "{}  [{}]".format(display, src_label)
            it = QListWidgetItem(label)
            it.setData(Qt.UserRole, path)
            self.list_widget.addItem(it)

    def _on_manage_dirs(self):
        """管理模板目录：查看、添加、移除（自带目录不可移除）"""
        dirs = load_template_dirs()
        home = os.path.expanduser("~")
        labels = []
        for d in dirs:
            label = d
            if label.startswith(home):
                label = "~" + label[len(home):]
            labels.append(label)

        from PyQt5.QtWidgets import QDialog, QListWidget as QL2, QDialogButtonBox
        dlg = QDialog(self)
        dlg.setWindowTitle("管理模板目录")
        dlg.resize(500, 350)
        dlv = QVBoxLayout(dlg)

        hint = QLabel("「软件自带」目录随打包发布，不可移除\n「添加」选择新目录  「移除」删除用户目录")
        hint.setStyleSheet("color: #888; font-size: 11px;")
        dlv.addWidget(hint)

        dl_list = QL2()
        for label, d in zip(labels, dirs):
            if is_bundled_dir(d):
                label = label + "（软件自带）"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, d)
            dl_list.addItem(item)
        dlv.addWidget(dl_list)

        dbb = QDialogButtonBox()
        btn_add = QPushButton("添加目录")
        btn_remove = QPushButton("移除选中")
        btn_close = QPushButton("关闭")
        dbb.addButton(btn_add, QDialogButtonBox.ActionRole)
        dbb.addButton(btn_remove, QDialogButtonBox.ActionRole)
        dbb.addButton(btn_close, QDialogButtonBox.AcceptRole)
        dlv.addWidget(dbb)

        def do_add():
            d = QFileDialog.getExistingDirectory(dlg, "选择模板目录", os.path.expanduser("~"))
            if not d:
                return
            cur = load_template_dirs()
            if d not in cur:
                cur.append(d)
                save_template_dirs(cur)
            dlg.accept()

        def do_remove():
            sel = dl_list.currentItem()
            if sel is None:
                return
            d = sel.data(Qt.UserRole)
            if is_bundled_dir(d):
                QMessageBox.information(dlg, "提示", "软件自带模板目录不可移除")
                return
            cur = load_template_dirs()
            # 除了自带目录外至少保留一个
            non_bundled = [x for x in cur if not is_bundled_dir(x)]
            if len(non_bundled) <= 1 and not is_bundled_dir(d):
                QMessageBox.information(dlg, "提示", "至少保留一个用户模板目录")
                return
            if d in cur:
                cur.remove(d)
                save_template_dirs(cur)
            dlg.accept()

        btn_add.clicked.connect(do_add)
        btn_remove.clicked.connect(do_remove)
        btn_close.clicked.connect(dlg.accept)
        dlg.exec_()
        self._load_template_list()

    # ---------- 选择模板 ----------
    def _on_select_template(self, current, _prev):
        if current is None:
            self.current_template_text = ""
            self.current_template_path = ""
            return
        self.current_template_path = current.data(Qt.UserRole)
        with open(self.current_template_path, encoding="utf-8") as f:
            self.current_template_text = f.read()
        self._rebuild_form()

    # ---------- 动态表单 ----------
    def _rebuild_form(self):
        while self.form_layout.rowCount():
            self.form_layout.removeRow(0)
        self.field_inputs.clear()
        for name in extract_fields(self.current_template_text):
            if any(k in name for k in ("情况", "说明", "内容", "简要", "事实", "简历", "案情")):
                w = QTextEdit()
                w.setFixedHeight(60)
                w.textChanged.connect(self._update_preview)
            else:
                w = QLineEdit()
                w.textChanged.connect(self._update_preview)
            self.field_inputs[name] = w
            self.form_layout.addRow(name + "：", w)
        self._update_preview()

    def _collect_values(self):
        vals = {}
        for name, w in self.field_inputs.items():
            vals[name] = (w.toPlainText() if isinstance(w, QTextEdit) else w.text()).strip()
        return vals

    def _current_rendered(self):
        return render(parse_template(self.current_template_text), self._collect_values())

    def _update_preview(self):
        if not self.current_template_text:
            return
        self.preview.setPlainText(_build_plain_text(self._current_rendered()))

    # ---------- 生成 & 排版 ----------
    def _on_generate(self):
        if not self.current_template_text:
            QMessageBox.warning(self, "提示", "请先从左侧选择文书模板")
            return
        rendered = self._current_rendered()
        unfilled = find_unfilled(rendered)
        if unfilled:
            if QMessageBox.question(
                self, "存在未填字段",
                "未填写：\n" + "、".join(unfilled) + "\n\n仍要生成吗？",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return

        plain = _build_plain_text(rendered)
        title = rendered.get("title", "未命名")

        default_dir = os.path.expanduser("~/Desktop")
        if not os.path.isdir(default_dir):
            default_dir = os.path.expanduser("~")
        safe_title = "".join(c for c in title if c not in r'\/:*?"<>|') or "未命名"
        default_path = os.path.join(default_dir, safe_title + "_排版.docx")

        out_path, _ = QFileDialog.getSaveFileName(
            self, "保存排版后的公文", default_path, "Word 文档 (*.docx)")
        if not out_path:
            return

        preset_name = self.mgr.active_key
        custom_settings = None
        if preset_name.startswith('user_'):
            preset = self.mgr.get(preset_name)
            if preset:
                custom_settings = {k: v for k, v in preset.items()
                                   if k not in ('key', 'user_key', 'builtin_key', 'created')}
                preset_name = 'custom'

        self.gen_btn.setEnabled(False)
        self.gen_btn.setText("正在生成并排版...")
        QApplication.processEvents()

        self.worker = DraftFormatWorker(plain, out_path, preset_name, custom_settings)
        self.worker.finishedWith.connect(self._on_format_done)
        self.worker.start()

    def _on_format_done(self, success, result):
        self.gen_btn.setEnabled(True)
        self.gen_btn.setText("生成初稿并排版")
        if success:
            QMessageBox.information(self, "完成", "已排版生成：\n{}".format(result))
        else:
            QMessageBox.critical(self, "生成失败", result)
