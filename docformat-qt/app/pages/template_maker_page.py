# -*- coding: utf-8 -*-
"""
============================================================================
模块二：模板制作页 (TemplateMakerPage) — 历史件转模板
============================================================================

【用途】
把一份"已脱密的历史公文"快速转成可复用模板。
核心是"机器抽结构 + 人工点选挖空"：

  · 机器自动做的：切分标题/正文段落/落款、识别"一、""（一）"层级。
  · 机器做不到的：判断哪些内容是"固定套话"（该保留）、哪些是
    "案件变量"（该挖成 {{占位符}}）。由用户框选文字 → 点按钮挖空。

【工作流】
  1. 载入脱密历史件（.docx/.doc/.wps 或直接粘贴纯文本）
  2. 软件自动切分结构，显示在可编辑文本区
  3. 在文本区选中要变成变量的文字 → 点"挖空为占位符" → 输入字段名
  4. 可以设置任意附加字段（法律名称、身份信息、事实陈述等），在正文中挖空
  5. 在"附加字段"表格中配置署名/日期等结构性字段
  6. 点"保存为模板" → 生成 .md 模板

【模板格式】见 template_common.py
============================================================================
"""
import os
import re
import sys
import tempfile
import shutil

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QLineEdit, QFileDialog, QMessageBox, QInputDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QAbstractItemView, QSplitter, QMenu, QCheckBox,
)
from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt

from app.template_common import (
    TEMPLATE_DIR, PLACEHOLDER_RE,
    load_template_dirs, save_template_dirs, is_bundled_dir,
    detect_auto_fields, find_placeholder_at,
    load_quick_inserts, save_quick_inserts, DEFAULT_QUICK_INSERTS,
)


def _ensure_docx(path):
    """非 .docx 格式先转换为 .docx，返回 (可读取的 .docx 路径, 临时目录或None)"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        return path, None

    if ext in ('.doc', '.wps'):
        tmp_dir = tempfile.mkdtemp(prefix='docformat_template_')
        tmp_path = os.path.join(tmp_dir, os.path.splitext(os.path.basename(path))[0] + '.docx')
        if sys.platform == 'win32':
            from scripts import converter
            converter.convert_to_docx(path, tmp_path)
        else:
            from app import converter_linux
            produced = converter_linux.convert_to_docx(path, tmp_path)
            if produced != tmp_path:
                tmp_path = produced
        return tmp_path, tmp_dir

    raise ValueError("不支持的文件格式：{}".format(ext))


def _load_docx_text(path):
    """读取 docx 全部段落纯文本"""
    from docx import Document
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def _rough_split(text):
    """粗切结构：首个非空行标为标题；其余按行保留。"""
    lines = [l.rstrip() for l in text.splitlines()]
    out, title_done = [], False
    for l in lines:
        if not l.strip():
            continue
        if not title_done:
            out.append("标题: " + l.strip())
            title_done = True
        else:
            out.append(l)
    return "\n".join(out)


class TemplateMakerPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.field_names = []
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)

        # 顶部：载入 + 目录
        top = QHBoxLayout()
        self.btn_open = QPushButton("载入历史件(.docx)")
        self.btn_open.setToolTip("支持 .docx / .doc / .wps")
        self.btn_open.clicked.connect(self._on_open)
        top.addWidget(self.btn_open)
        top.addWidget(QLabel("或直接粘贴到下方文本框"))
        top.addStretch()

        top.addWidget(QLabel("保存到："))
        self.dir_combo = QComboBox()
        self.dir_combo.setMinimumWidth(280)
        self.dir_combo.setEditable(False)
        self._refresh_dir_combo()
        top.addWidget(self.dir_combo)

        self.btn_add_dir = QPushButton("+")
        self.btn_add_dir.setFixedWidth(30)
        self.btn_add_dir.setToolTip("添加模板保存目录")
        self.btn_add_dir.clicked.connect(self._on_add_dir)
        top.addWidget(self.btn_add_dir)
        outer.addLayout(top)

        # 正文编辑区 + 工具条紧凑布局
        outer.addWidget(QLabel("正文（选中文字 → 点下方按钮挖空为占位符）"))
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(False)
        self.editor.setMinimumHeight(200)
        self.editor.setContextMenuPolicy(Qt.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self._on_editor_context_menu)
        outer.addWidget(self.editor)

        mid = QHBoxLayout()
        self.btn_hole = QPushButton("① 挖空选中文字为占位符")
        self.btn_hole.clicked.connect(self._on_make_hole)
        self.btn_mark_level = QPushButton("② 标记选中行标题层级 ▾")
        self.btn_mark_level.clicked.connect(self._on_mark_level_menu)
        self.btn_quick = QPushButton("③ 快捷插入")
        self.btn_quick.clicked.connect(lambda: self._show_quick_menu(self.btn_quick))
        mid.addWidget(self.btn_hole)
        mid.addWidget(self.btn_mark_level)
        mid.addWidget(self.btn_quick)
        mid.addStretch()
        # 模板标签和工具条同行
        mid.addWidget(QLabel("标签："))
        self.tags_edit = QLineEdit()
        self.tags_edit.setPlaceholderText("逗号分隔，辅助起草页搜索")
        self.tags_edit.setMaximumWidth(240)
        mid.addWidget(self.tags_edit)
        outer.addLayout(mid)

        # 附加字段（META 键值对）表格
        meta_label = QLabel("附加字段（正文挖空后自动同步到此表，也可手动增删）")
        meta_label.setContentsMargins(0, 4, 0, 0)
        outer.addWidget(meta_label)
        meta_bar = QHBoxLayout()
        self.meta_table = QTableWidget(0, 2)
        self.meta_table.setHorizontalHeaderLabels(["字段名", "字段值（可用 {{占位符}}）"])
        self.meta_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.meta_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.meta_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.meta_table.setMaximumHeight(140)
        self.meta_table.verticalHeader().setVisible(False)
        # 独立配色：在任何主题下都保持可读
        self.meta_table.setStyleSheet("""
            QTableWidget { background: #FAFAFA; border: 1px solid #CCC; border-radius: 6px;
                gridline-color: #E0E0E0; }
            QTableWidget::item { padding: 4px 8px; color: #1A1A1A; }
            QTableWidget::item:selected { background: #DEE2E6; }
            QHeaderView::section { background: #EEE; color: #333; border: none;
                border-bottom: 1px solid #CCC; padding: 6px 8px; font-weight: 600; }
        """)
        meta_bar.addWidget(self.meta_table)

        meta_btns = QVBoxLayout()
        self.btn_meta_add = QPushButton("＋")
        self.btn_meta_add.setFixedSize(30, 30)
        self.btn_meta_add.clicked.connect(self._on_meta_add_row)
        self.btn_meta_del = QPushButton("－")
        self.btn_meta_del.setFixedSize(30, 30)
        self.btn_meta_del.clicked.connect(self._on_meta_del_row)
        meta_btns.addWidget(self.btn_meta_add)
        meta_btns.addWidget(self.btn_meta_del)
        meta_btns.addStretch()
        meta_bar.addLayout(meta_btns)
        outer.addLayout(meta_bar)

        # 预设两条：落款单位、落款日期
        self._add_meta_row("落款单位", "{{办案单位}}")
        self._add_meta_row("落款日期", "{{落款日期}}")

        # 保存区
        bottom = QHBoxLayout()
        bottom.addWidget(QLabel("模板名:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("必填，如 技术调查措施请示")
        self.name_edit.setMinimumWidth(180)
        bottom.addWidget(self.name_edit)
        self.btn_browse_save = QPushButton("浏览...")
        self.btn_browse_save.setToolTip("自定义保存路径和文件名")
        self.btn_browse_save.clicked.connect(self._on_browse_save)
        bottom.addWidget(self.btn_browse_save)
        self.btn_save = QPushButton("保存为模板")
        self.btn_save.clicked.connect(self._on_save)
        bottom.addWidget(self.btn_save)
        bottom.addStretch()
        outer.addLayout(bottom)

    # ---------------- 目录管理 ----------------
    def _refresh_dir_combo(self):
        self.dir_combo.clear()
        dirs = load_template_dirs()
        for d in dirs:
            label = d
            home = os.path.expanduser("~")
            if label.startswith(home):
                label = "~" + label[len(home):]
            if is_bundled_dir(d):
                label = label + "（软件自带）"
            self.dir_combo.addItem(label, d)

    def _on_add_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择模板保存目录", os.path.expanduser("~"))
        if not d:
            return
        dirs = load_template_dirs()
        if d not in dirs:
            dirs.append(d)
            save_template_dirs(dirs)
        self._refresh_dir_combo()
        idx = self.dir_combo.findData(d)
        if idx >= 0:
            self.dir_combo.setCurrentIndex(idx)

    # ---------------- 载入 ----------------
    def _on_open(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择脱密历史件", "",
            "Word 文档 (*.docx *.doc *.wps)")
        if not path:
            return
        tmp_dir = None
        try:
            work_path, tmp_dir = _ensure_docx(path)
            if tmp_dir:
                QMessageBox.information(self, "格式转换",
                    "已将 {} 转换为 .docx，接下来请检查结构并挖空变量。".format(
                        os.path.splitext(path)[1]))
            text = _load_docx_text(work_path)
        except ValueError as e:
            QMessageBox.critical(self, "格式不支持", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))
            return
        finally:
            if tmp_dir:
                shutil.rmtree(tmp_dir, ignore_errors=True)
        self.editor.setPlainText(_rough_split(text))
        self._auto_detect_and_ask(self.editor.toPlainText())

    # ---------------- 挖空 ----------------
    def _on_make_hole(self):
        cursor = self.editor.textCursor()
        selected = cursor.selectedText().strip()
        if not selected:
            QMessageBox.information(self, "提示", "请先在正文中选中要挖空的文字")
            return
        field, ok = QInputDialog.getText(self, "命名字段",
            "把「{}」变成占位符，字段名：".format(selected))
        field = field.strip()
        if not ok or not field:
            return
        full = self.editor.toPlainText()
        placeholder = "{{" + field + "}}"
        new_full = full.replace(selected, placeholder)
        self.editor.setPlainText(new_full)
        if field not in self.field_names:
            self.field_names.append(field)
            # 同步到附加字段表格
            existing_keys = set()
            for r in range(self.meta_table.rowCount()):
                item = self.meta_table.item(r, 0)
                if item:
                    existing_keys.add(item.text().strip())
            if field not in existing_keys:
                self._add_meta_row(field, "{{" + field + "}}")
        QMessageBox.information(self, "已挖空",
            "已把所有「{}」替换为 {}".format(selected, placeholder))

    def _on_mark_level_menu(self):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText().strip()
        if not line:
            return
        menu = QMenu(self)
        for level, label in [(0, "主标题"), (1, "一级标题"), (2, "二级标题"), (3, "三级标题"), (4, "四级标题")]:
            action = menu.addAction(label)
            action.triggered.connect(lambda checked, l=level: self._mark_line_level(l))
        menu.exec_(self.btn_mark_level.mapToGlobal(self.btn_mark_level.rect().bottomLeft()))

    def _mark_line_level(self, level):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText().strip()
        if not line:
            return
        # 去掉已有的标记前缀
        clean = re.sub(r'^标题[:：]\d*[:：]\s*', '', line).strip()
        if level == 0:
            cursor.insertText("标题: " + clean)
        else:
            cursor.insertText("标题:{}: {}".format(level, clean))

    def _on_editor_context_menu(self, pos):
        """右键菜单：挖空选中文字 / 调整已有占位符"""
        cursor = self.editor.textCursor()
        cursor_pos = cursor.position()
        full_text = self.editor.toPlainText()
        selected = cursor.selectedText().strip()

        menu = QMenu(self)

        # 光标在占位符上时
        ph_info = find_placeholder_at(full_text, cursor_pos) if not selected else None
        action_rename = action_fill = action_unhole = None
        if ph_info:
            _, field_name, _, _ = ph_info
            menu.addAction("占位符 {{" + field_name + "}}").setEnabled(False)
            action_rename = menu.addAction("重命名字段...")
            action_fill = menu.addAction("填入具体值...")
            action_unhole = menu.addAction("取消挖空（删除 {{ }} 标记）")
            menu.addSeparator()

        action_hole = None
        if selected:
            action_hole = menu.addAction("挖空为占位符")
            menu.addSeparator()

        menu.addAction("全选", self.editor.selectAll)
        action = menu.exec_(self.editor.mapToGlobal(pos))

        if action == action_rename and ph_info:
            self._do_rename_in_editor(ph_info)
        elif action == action_fill and ph_info:
            self._do_fill_in_editor(ph_info)
        elif action == action_unhole and ph_info:
            self._do_unhole_in_editor(ph_info)
        elif action == action_hole:
            self._on_make_hole()

    def _do_replace_in_editor(self, ph_info, new_text):
        """替换占位符文本"""
        _, _, start, end = ph_info
        full = self.editor.toPlainText()
        new_full = full[:start] + new_text + full[end:]
        self.editor.setPlainText(new_full)

    def _do_rename_in_editor(self, ph_info):
        _, old_name, _, _ = ph_info
        new_name, ok = QInputDialog.getText(self, "重命名字段",
            "将 {} 重命名为：".format(old_name), text=old_name)
        if not ok or not new_name.strip():
            return
        self._do_replace_in_editor(ph_info, "{{" + new_name.strip() + "}}")

    def _do_fill_in_editor(self, ph_info):
        _, field_name, _, _ = ph_info
        value, ok = QInputDialog.getText(self, "填入具体值",
            "为 {} 填入具体值（将删除 {{}} 标记）：".format(field_name))
        if not ok:
            return
        self._do_replace_in_editor(ph_info, value.strip())

    def _do_unhole_in_editor(self, ph_info):
        _, field_name, _, _ = ph_info
        self._do_replace_in_editor(ph_info, field_name)

    def _auto_detect_and_ask(self, text):
        """扫描文本，检测身份证号/法律条款/日期等，提示用户批量挖空"""
        matches = detect_auto_fields(text)
        if not matches:
            return
        dlg = QMessageBox(self)
        dlg.setWindowTitle("自动检测到可挖空字段")
        dlg.setText("软件检测到以下可能反复变化的字段，\n勾选要自动挖空的项：")
        dlg.setIcon(QMessageBox.Question)

        # 用 QCheckBox 列表代替标准按钮
        widget = QWidget()
        lv = QVBoxLayout(widget)
        checks = []
        for i, (matched, field, label) in enumerate(matches):
            cb = QCheckBox("「{}」→ 挖空为 {{{{ {} }}}}（{}）".format(
                matched if len(matched) <= 20 else matched[:17] + "...",
                field, label))
            cb.setChecked(True)
            cb._field = field
            cb._matched = matched
            checks.append(cb)
            lv.addWidget(cb)
        dlg.layout().addWidget(widget, 1, 0, 1, dlg.layout().columnCount())

        dlg.addButton("全部挖空", QMessageBox.AcceptRole)
        dlg.addButton("跳过", QMessageBox.RejectRole)
        dlg.setDefaultButton(dlg.buttons()[0])
        if dlg.exec_() == QMessageBox.Rejected:
            return

        full = self.editor.toPlainText()
        for cb in checks:
            if cb.isChecked():
                placeholder = "{{" + cb._field + "}}"
                full = full.replace(cb._matched, placeholder)
                if cb._field not in self.field_names:
                    self.field_names.append(cb._field)
        self.editor.setPlainText(full)

    # ---------------- 快捷插入 ----------------
    def _show_quick_menu(self, anchor_widget=None):
        menu = QMenu(self)
        for item in load_quick_inserts():
            label = item.get("label", "")
            text = item.get("text", "")
            if not label:
                continue
            action = menu.addAction(label)
            action.setData(text)
            action.triggered.connect(lambda checked, t=text: self._insert_quick_text(t))
        menu.addSeparator()
        menu.addAction("管理快捷插入...", self._on_manage_quick_inserts)
        if anchor_widget:
            menu.exec_(anchor_widget.mapToGlobal(anchor_widget.rect().bottomLeft()))
        else:
            menu.exec_(self.editor.mapToGlobal(self.editor.rect().center()))

    def _insert_quick_text(self, text):
        cursor = self.editor.textCursor()
        cursor.insertText(text)
        self.editor.setTextCursor(cursor)

    def _on_manage_quick_inserts(self):
        from app.pages.template_draft_page import QuickInsertDialog
        dlg = QuickInsertDialog(self)
        dlg.exec_()

    # ---------------- META 动态表格 ----------------
    def _add_meta_row(self, key="", value=""):
        row = self.meta_table.rowCount()
        self.meta_table.insertRow(row)
        self.meta_table.setItem(row, 0, QTableWidgetItem(key))
        self.meta_table.setItem(row, 1, QTableWidgetItem(value))

    def _on_meta_add_row(self):
        self._add_meta_row()

    def _on_meta_del_row(self):
        rows = set(idx.row() for idx in self.meta_table.selectedIndexes())
        for row in sorted(rows, reverse=True):
            self.meta_table.removeRow(row)

    def _collect_meta(self):
        """收集 META 表格中的非空键值对"""
        meta = {}
        for row in range(self.meta_table.rowCount()):
            key_item = self.meta_table.item(row, 0)
            val_item = self.meta_table.item(row, 1)
            key = key_item.text().strip() if key_item else ""
            val = val_item.text().strip() if val_item else ""
            if key:
                meta[key] = val
        return meta

    # ---------------- 浏览保存路径 ----------------
    def _on_browse_save(self):
        name = self.name_edit.text().strip()
        if not name:
            name = "未命名模板"
        safe = "".join(c for c in name if c not in r'\/:*?"<>|')
        default_dir = self.dir_combo.currentData() or os.path.expanduser("~/Desktop")
        if not os.path.isdir(default_dir):
            default_dir = os.path.expanduser("~")
        default_path = os.path.join(default_dir, safe + ".md")

        path, _ = QFileDialog.getSaveFileName(
            self, "选择模板保存位置", default_path,
            "Markdown 模板 (*.md)")
        if not path:
            return
        self._do_save(path)

    # ---------------- 保存 ----------------
    def _on_save(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "请填写模板名")
            return
        body = self.editor.toPlainText().strip()
        if not body:
            QMessageBox.warning(self, "提示", "正文为空")
            return

        save_dir = self.dir_combo.currentData() or TEMPLATE_DIR
        os.makedirs(save_dir, exist_ok=True)
        safe = "".join(c for c in name if c not in r'\/:*?"<>|')
        path = os.path.join(save_dir, safe + ".md")
        if os.path.exists(path):
            if QMessageBox.question(self, "已存在",
                "{} 已存在，覆盖吗？".format(safe + ".md"),
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return
        self._do_save(path)

    def _do_save(self, path):
        """执行保存（组装模板内容并写盘）"""
        body = self.editor.toPlainText().strip()
        name = self.name_edit.text().strip() or os.path.splitext(os.path.basename(path))[0]

        # 自动检测标题层级（仅对未手动标记的行生效）
        # 不影响排版引擎——只给模板系统加结构标记
        from scripts.formatter import detect_para_type, DEFAULT_DETECT_RULES
        rules = {k: v for k, v in DEFAULT_DETECT_RULES.items()}
        lines = body.splitlines()
        all_texts = [l.strip() for l in lines if l.strip() and not l.strip().startswith("标题")]
        idx_map = {}
        n = 0
        for i, l in enumerate(lines):
            if l.strip() and not l.strip().startswith("标题"):
                idx_map[i] = n
                n += 1
        prev_type = None
        auto_lines = []
        for i, line in enumerate(lines):
            s = line.strip()
            if not s:
                auto_lines.append(line)
                continue
            if s.startswith("标题"):
                auto_lines.append(line)
                prev_type = None
                continue
            ai = idx_map.get(i)
            ptype = detect_para_type(s, i, len(lines), None, all_texts, ai, prev_type, rules=rules)
            prev_type = ptype
            level_map = {'heading1': 1, 'heading2': 2, 'heading3': 3, 'heading4': 4}
            if ptype == 'title':
                auto_lines.append("标题: " + s)
            elif ptype in level_map:
                auto_lines.append("标题:{}: {}".format(level_map[ptype], s))
            else:
                auto_lines.append(line)
        body = "\n".join(auto_lines)

        # 组装模板内容
        parts = [body]

        # 标签（统一使用 // tags: 行注释语法）
        tags = self.tags_edit.text().strip()
        if tags:
            parts.append("")
            parts.append("// tags: {}".format(tags))

        parts.append("")
        parts.append("---META---")
        for k, v in self._collect_meta().items():
            parts.append("{}: {}".format(k, v))
        # 如果 META 完全为空，留一行占位
        if not self._collect_meta():
            parts.append("# 可在此添加附加字段，如 落款单位: {{办案单位}}")
        content = "\n".join(parts) + "\n"

        save_dir = os.path.dirname(path)
        os.makedirs(save_dir, exist_ok=True)

        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

        # 汇总所有占位符（正文 + META）
        body_fields = re.findall(r"\{\{\s*([^}]+?)\s*\}\}", body)
        meta_fields = re.findall(r"\{\{\s*([^}]+?)\s*\}\}", "\n".join(parts))
        all_fields = sorted(set(body_fields + meta_fields))

        # 确保此目录在模板目录列表中
        dirs = load_template_dirs()
        if save_dir not in dirs:
            dirs.append(save_dir)
            save_template_dirs(dirs)

        QMessageBox.information(self, "已保存",
            "模板已存至：\n{}\n\n包含占位符：{}\n"
            "现在可到「模板起草」页选用它。".format(
                path,
                "、".join(all_fields) if all_fields else "（无）"))
