# -*- coding: utf-8 -*-
"""
============================================================================
模块一：模板起草页 (TemplateDraftPage)
============================================================================

【用途】
给 docformat-qt 增加"起草"能力。选文书模板 → 填字段 → 预览/编辑 → 排版引擎排版。

【功能】
  - 搜索模板、多目录管理
  - 字段自动生成（含"情况/说明"等关键词 → 多行框）
  - 预览窗口可直接编辑文本，右键快捷插入常用事项
  - 可编辑模板源码并保存回 .md 文件
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
    QApplication, QMenu, QAction, QDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QDialogButtonBox, QInputDialog,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from app.template_common import (
    TEMPLATE_DIR, PLACEHOLDER_RE,
    extract_fields, parse_template, render, find_unfilled,
    load_template_dirs, save_template_dirs, scan_templates,
    bundled_templates_dir, is_bundled_dir,
    load_quick_inserts, save_quick_inserts, DEFAULT_QUICK_INSERTS,
    find_placeholder_at, search_templates, read_template_preview,
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


class QuickInsertDialog(QDialog):
    """管理快捷插入项目"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("管理快捷插入")
        self.resize(500, 400)
        self.items = load_quick_inserts()
        self._build()

    def _build(self):
        lv = QVBoxLayout(self)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["菜单名", "插入内容"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        lv.addWidget(self.table)

        for item in self.items:
            self._add_row(item["label"], item["text"])

        btn_row = QHBoxLayout()
        btn_add = QPushButton("添加")
        btn_del = QPushButton("删除选中")
        btn_reset = QPushButton("恢复默认")
        btn_save = QPushButton("保存")
        btn_add.clicked.connect(lambda: self._add_row("", ""))
        btn_del.clicked.connect(self._del_row)
        btn_reset.clicked.connect(self._reset_default)
        btn_save.clicked.connect(self._save)
        for b in [btn_add, btn_del, btn_reset, btn_save]:
            btn_row.addWidget(b)
        lv.addLayout(btn_row)

    def _add_row(self, label, text):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(label))
        self.table.setItem(r, 1, QTableWidgetItem(text))

    def _del_row(self):
        rows = set(i.row() for i in self.table.selectedIndexes())
        for r in sorted(rows, reverse=True):
            self.table.removeRow(r)

    def _reset_default(self):
        self.table.setRowCount(0)
        for item in DEFAULT_QUICK_INSERTS:
            self._add_row(item["label"], item["text"])

    def _save(self):
        items = []
        for r in range(self.table.rowCount()):
            label = (self.table.item(r, 0).text() if self.table.item(r, 0) else "").strip()
            text = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").strip()
            if label:
                items.append({"label": label, "text": text})
        save_quick_inserts(items)
        self.accept()


class TemplateDraftPage(QWidget):
    def __init__(self, mgr, parent=None):
        super().__init__(parent)
        self.mgr = mgr
        self.current_template_text = ""
        self.current_template_path = ""
        self.field_inputs = {}
        self._all_templates = []
        self._edit_source_mode = False   # False=预览模式 True=编辑模板源码
        self._build_ui()
        self._load_template_list()

    def _build_ui(self):
        splitter = QSplitter(Qt.Horizontal)

        # ===== 左侧：模板列表 + 搜索 + 目录管理 =====
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(8, 8, 4, 8)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索模板（模糊匹配）...")
        self.search_edit.textChanged.connect(self._on_search)
        lv.addWidget(self.search_edit)

        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_select_template)
        lv.addWidget(self.list_widget)

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

        # ===== 右侧：预览 / 编辑 =====
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(4, 8, 8, 8)

        # 工具栏
        preview_bar = QHBoxLayout()
        mode_label = QLabel("预览/编辑")
        preview_bar.addWidget(mode_label)
        preview_bar.addStretch()

        self.btn_toggle_mode = QPushButton("编辑模板源码")
        self.btn_toggle_mode.setCheckable(True)
        self.btn_toggle_mode.toggled.connect(self._on_toggle_mode)
        preview_bar.addWidget(self.btn_toggle_mode)

        self.btn_save_template = QPushButton("保存模板")
        self.btn_save_template.clicked.connect(self._on_save_template)
        self.btn_save_template.setEnabled(False)
        preview_bar.addWidget(self.btn_save_template)

        btn_quick = QPushButton("快捷插入")
        btn_quick.clicked.connect(lambda: self._show_quick_menu(btn_quick))
        preview_bar.addWidget(btn_quick)
        rv.addLayout(preview_bar)

        # 预览编辑区
        self.preview = QPlainTextEdit()
        self.preview.setContextMenuPolicy(Qt.CustomContextMenu)
        self.preview.customContextMenuRequested.connect(self._on_preview_context_menu)
        rv.addWidget(self.preview)
        splitter.addWidget(right)

        splitter.setSizes([240, 360, 420])
        outer = QVBoxLayout(self)
        outer.addWidget(splitter)

    # ===== 预览模式切换 =====
    def _on_toggle_mode(self, checked):
        self._edit_source_mode = checked
        self.btn_toggle_mode.setText("返回预览" if checked else "编辑模板源码")
        self.btn_save_template.setEnabled(checked)
        if checked:
            self.preview.setStyleSheet("QPlainTextEdit { background: #fffef5; }")
            if self.current_template_text:
                self.preview.setPlainText(self.current_template_text)
        else:
            self.preview.setStyleSheet("")
            self._update_preview()

    # ===== 保存模板（源码编辑模式下） =====
    def _on_save_template(self):
        if not self.current_template_path:
            QMessageBox.warning(self, "提示", "请先选择一个模板")
            return
        new_text = self.preview.toPlainText()
        if not new_text.strip():
            QMessageBox.warning(self, "提示", "模板内容不能为空")
            return
        try:
            with open(self.current_template_path, "w", encoding="utf-8") as f:
                f.write(new_text)
            self.current_template_text = new_text
            self._rebuild_form()
            QMessageBox.information(self, "已保存", "模板已更新：{}".format(self.current_template_path))
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    # ===== 快捷插入 =====
    def _show_quick_menu(self, anchor_widget=None):
        menu = QMenu(self)
        for item in load_quick_inserts():
            label = item.get("label", "")
            text = item.get("text", "")
            if not label:
                continue
            action = menu.addAction(label)
            action.setData(text)
            action.triggered.connect(lambda checked, t=text: self._insert_text(t))
        menu.addSeparator()
        menu.addAction("管理快捷插入...", self._on_manage_quick_inserts)
        if anchor_widget:
            menu.exec_(anchor_widget.mapToGlobal(anchor_widget.rect().bottomLeft()))
        else:
            menu.exec_(self.preview.mapToGlobal(self.preview.rect().center()))

    def _on_preview_context_menu(self, pos):
        menu = QMenu(self)
        cursor = self.preview.textCursor()
        cursor_pos = cursor.position()
        full_text = self.preview.toPlainText()
        selected = cursor.selectedText().strip()

        # 光标在占位符上时，提供占位符操作
        ph_info = find_placeholder_at(full_text, cursor_pos) if not selected else None
        action_rename = action_fill = action_unhole = None
        if ph_info:
            _, field_name, _, _ = ph_info
            menu.addAction("占位符 {{" + field_name + "}}").setEnabled(False)
            action_rename = menu.addAction("重命名字段...")
            action_fill = menu.addAction("填入具体值...")
            action_unhole = menu.addAction("取消挖空（删除 {{ }} 标记）")
            menu.addSeparator()

        # 选中文字时，增加挖空选项
        action_hole = None
        if selected:
            label = "挖空「{}」为占位符...".format(
                selected if len(selected) <= 15 else selected[:12] + "...")
            action_hole = menu.addAction(label)
            menu.addSeparator()

        for item in load_quick_inserts():
            label = item.get("label", "")
            text = item.get("text", "")
            if not label:
                continue
            action = menu.addAction(label)
            action.setData(text)

        menu.addSeparator()
        menu.addAction("管理快捷插入...", self._on_manage_quick_inserts)
        action = menu.exec_(self.preview.mapToGlobal(pos))

        if action == action_rename and ph_info:
            self._do_rename_placeholder(ph_info)
        elif action == action_fill and ph_info:
            self._do_fill_placeholder(ph_info)
        elif action == action_unhole and ph_info:
            self._do_unhole_placeholder(ph_info)
        elif action == action_hole:  # was action_hole
            self._do_hollow_in_preview(selected)
        elif action and action.data():
            self._insert_text(action.data())

    def _do_placeholder_op(self, ph_info, new_text):
        """替换占位符并刷新"""
        _, _, start, end = ph_info
        full = self.preview.toPlainText()
        new_full = full[:start] + new_text + full[end:]
        self.preview.setPlainText(new_full)
        self.current_template_text = new_full
        self._rebuild_form()

    def _do_rename_placeholder(self, ph_info):
        _, old_name, _, _ = ph_info
        new_name, ok = QInputDialog.getText(self, "重命名字段",
            "将 {} 重命名为：".format(old_name), text=old_name)
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        self._do_placeholder_op(ph_info, "{{" + new_name + "}}")

    def _do_fill_placeholder(self, ph_info):
        _, field_name, _, _ = ph_info
        value, ok = QInputDialog.getText(self, "填入具体值",
            "为 {} 填入具体值（将删除 {{}} 标记）：".format(field_name))
        if not ok:
            return
        self._do_placeholder_op(ph_info, value.strip())

    def _do_unhole_placeholder(self, ph_info):
        _, field_name, _, _ = ph_info
        # 删除 {{ }} 标记，保留字段名作为普通文字
        self._do_placeholder_op(ph_info, field_name)

    def _do_hollow_in_preview(self, selected):
        field, ok = QInputDialog.getText(self, "命名字段",
            "把「{}」变成占位符，字段名：".format(selected))
        field = field.strip()
        if not ok or not field:
            return
        placeholder = "{{" + field + "}}"
        full = self.preview.toPlainText()
        new_full = full.replace(selected, placeholder)
        self.preview.setPlainText(new_full)
        self.current_template_text = new_full
        self._rebuild_form()

    def _insert_text(self, text):
        cursor = self.preview.textCursor()
        cursor.insertText(text)

    def _on_manage_quick_inserts(self):
        dlg = QuickInsertDialog(self)
        dlg.exec_()

    # ===== 模板列表 & 搜索 =====
    def _load_template_list(self):
        self._refresh_list_display()

    def _on_search(self, query):
        self._refresh_list_display(query)

    def _refresh_list_display(self, query=None):
        self.list_widget.clear()
        q = (query or "").strip()

        # 搜不到时的回退：仍用 scan_templates 以免缓存未命中
        results = search_templates(q)
        if not results and not q:
            # 缓存未命中时走旧逻辑兜底
            self._all_templates = scan_templates()
            for display, path, src_dir in self._all_templates:
                home = os.path.expanduser("~")
                src_label = src_dir
                if src_label.startswith(home):
                    src_label = "~" + src_label[len(home):]
                label = "{}  [{}]".format(display, src_label)
                it = QListWidgetItem(label)
                it.setData(Qt.UserRole, path)
                self.list_widget.addItem(it)
            return

        # 存储当前完整结果以便后续使用
        self._all_templates = [(d, p, s) for d, p, s, _ in results]

        shown_count = 0
        for display, path, src_dir, match_hint in results:
            home = os.path.expanduser("~")
            src_label = src_dir
            if src_label.startswith(home):
                src_label = "~" + src_label[len(home):]

            # 构建列表项显示文本
            preview = read_template_preview(path)
            tags = preview.get("tags", [])
            tag_str = "  #{}".format(" #".join(tags)) if tags else ""

            if q and match_hint:
                label = "{}  [{}]  — {}".format(display, src_label, match_hint)
            elif q:
                label = "{}  [{}]".format(display, src_label)
            else:
                label = "{}  [{}]{}".format(display, src_label, tag_str)

            it = QListWidgetItem(label)
            it.setData(Qt.UserRole, path)
            if tags:
                it.setToolTip("标签: {}".format("、".join(tags)))
            self.list_widget.addItem(it)
            shown_count += 1

        if q and shown_count == 0:
            it = QListWidgetItem("未找到匹配的模板")
            it.setFlags(Qt.NoItemFlags)
            self.list_widget.addItem(it)

    def _on_manage_dirs(self):
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

    # ===== 选择模板 =====
    def _on_select_template(self, current, _prev):
        if current is None:
            self.current_template_text = ""
            self.current_template_path = ""
            return
        self.current_template_path = current.data(Qt.UserRole)
        with open(self.current_template_path, encoding="utf-8") as f:
            self.current_template_text = f.read()
        self.btn_save_template.setEnabled(False)
        if self._edit_source_mode:
            self.preview.setPlainText(self.current_template_text)
        self._rebuild_form()

    # ===== 动态表单 =====
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
        if self._edit_source_mode:
            return  # 源码编辑模式不自动刷新
        self.preview.setPlainText(_build_plain_text(self._current_rendered()))

    # ===== 生成 & 排版 =====
    def _on_generate(self):
        if not self.current_template_text:
            QMessageBox.warning(self, "提示", "请先从左侧选择文书模板")
            return

        if self._edit_source_mode:
            # 源码编辑模式下，使用预览区的文本作为最终正文
            plain = self.preview.toPlainText()
        else:
            rendered = self._current_rendered()
            unfilled = find_unfilled(rendered)
            if unfilled:
                if QMessageBox.question(
                    self, "存在未填字段",
                    "未填写：\n" + "、".join(unfilled) + "\n\n仍要生成吗？",
                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                    return
            plain = _build_plain_text(rendered)

        title = "未命名"
        for line in plain.splitlines():
            s = line.strip()
            if s:
                title = s
                break

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
