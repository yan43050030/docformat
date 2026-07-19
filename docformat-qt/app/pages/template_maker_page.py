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
    QAbstractItemView, QSplitter,
)
from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt

from app.template_common import (
    TEMPLATE_DIR, PLACEHOLDER_RE,
    load_template_dirs, save_template_dirs,
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

        # 正文编辑区
        outer.addWidget(QLabel("正文（选中要变成变量的文字，点下方按钮挖空）"))
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(False)
        outer.addWidget(self.editor)

        # 工具条
        mid = QHBoxLayout()
        self.btn_hole = QPushButton("① 挖空选中文字为占位符")
        self.btn_hole.clicked.connect(self._on_make_hole)
        self.btn_title = QPushButton("② 标记选中行为标题")
        self.btn_title.clicked.connect(self._on_mark_title)
        mid.addWidget(self.btn_hole)
        mid.addWidget(self.btn_title)
        mid.addStretch()
        outer.addLayout(mid)

        # 附加字段（META 键值对）表格
        outer.addWidget(QLabel("附加字段（落款单位、日期等，可自由增删）"))
        meta_bar = QHBoxLayout()
        self.meta_table = QTableWidget(0, 2)
        self.meta_table.setHorizontalHeaderLabels(["字段名", "字段值（可用 {{占位符}}）"])
        self.meta_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.meta_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.meta_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.meta_table.setMaximumHeight(140)
        self.meta_table.verticalHeader().setVisible(False)
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
        bottom.addWidget(self.name_edit)
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
            # 显示时用 ~ 代替 home 目录
            home = os.path.expanduser("~")
            if d.startswith(home):
                label = "~" + d[len(home):]
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
        QMessageBox.information(self, "已挖空",
            "已把所有「{}」替换为 {}".format(selected, placeholder))

    def _on_mark_title(self):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText().strip()
        if not line:
            return
        if line.startswith("标题:") or line.startswith("标题："):
            return
        cursor.insertText("标题: " + line)

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

        # 组装模板内容
        parts = [body, "", "---META---"]
        for k, v in self._collect_meta().items():
            parts.append("{}: {}".format(k, v))
        # 如果 META 完全为空，留一行占位
        if not self._collect_meta():
            parts.append("# 可在此添加附加字段，如 落款单位: {{办案单位}}")
        content = "\n".join(parts) + "\n"

        save_dir = self.dir_combo.currentData() or TEMPLATE_DIR
        os.makedirs(save_dir, exist_ok=True)
        safe = "".join(c for c in name if c not in r'\/:*?"<>|')
        path = os.path.join(save_dir, safe + ".md")
        if os.path.exists(path):
            if QMessageBox.question(self, "已存在",
                "{} 已存在，覆盖吗？".format(safe + ".md"),
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return
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
