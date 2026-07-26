# -*- coding: utf-8 -*-
"""套头对位校验：把内容叠到套头纸 PDF 上看，不用打印机也能验证对不对齐。"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog, QFileDialog,
                             QHBoxLayout, QLabel, QMessageBox, QPushButton,
                             QScrollArea, QVBoxLayout, QWidget)

from scripts import overlay, overprint


class AlignDialog(QDialog):
    """选套头 PDF → 叠加预览 → 导出合并 PDF"""

    ZOOMS = [('适应宽度', 0), ('50%', 0.5), ('75%', 0.75), ('100%', 1.0),
             ('150%', 1.5), ('200%', 2.0)]

    def __init__(self, template_path, values, title_shape='trapezoid_down',
                 title_lines=None, parent=None):
        super(AlignDialog, self).__init__(parent)
        self.setWindowTitle("套头对位校验 — 不用打印也能看准不准")
        self.resize(880, 900)
        self._template_path = template_path
        self._values = values
        self._title_shape = title_shape
        self._title_lines = title_lines
        self._merged = None
        self._letterhead = overprint.load_letterhead(template_path)
        self._page = 0

        root = QVBoxLayout(self)
        tip = QLabel(
            "把生成的内容<b>叠</b>到套头纸（红头文件纸）的 PDF 上："
            "<span style='color:#C0392B'>红色</span>是纸上已经印好的，"
            "<b>黑色</b>是打印机这次要印上去的。"
            "看黑字有没有落进对应的空格、有没有压线，就知道对不对得准，"
            "不用试打浪费纸。<br>"
            "对不准就回上一步用「打印位置微调」改数值，改完再看一次。")
        tip.setWordWrap(True)
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        row = QHBoxLayout()
        row.addWidget(QLabel("套头纸 PDF："))
        self.path_label = QLabel(self._letterhead or "（未选择）")
        self.path_label.setWordWrap(True)
        row.addWidget(self.path_label, 1)
        btn_pick = QPushButton("选择…")
        btn_pick.clicked.connect(self._pick)
        row.addWidget(btn_pick)
        self.btn_clear = QPushButton("清除")
        self.btn_clear.clicked.connect(self._clear)
        row.addWidget(self.btn_clear)
        root.addLayout(row)

        row2 = QHBoxLayout()
        self.btn_check = QPushButton("生成对位预览")
        self.btn_check.clicked.connect(self._check)
        row2.addWidget(self.btn_check)
        self.btn_export = QPushButton("导出合并 PDF…")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export)
        row2.addWidget(self.btn_export)
        row2.addWidget(QLabel("显示："))
        self.zoom = QComboBox()
        for text, val in self.ZOOMS:
            self.zoom.addItem(text, val)
        self.zoom.currentIndexChanged.connect(self._redraw)
        row2.addWidget(self.zoom)
        self.page_combo = QComboBox()
        self.page_combo.setVisible(False)
        self.page_combo.currentIndexChanged.connect(self._on_page)
        row2.addWidget(self.page_combo)
        row2.addStretch(1)
        root.addLayout(row2)

        self.canvas = QLabel("选好套头纸 PDF 后，点「生成对位预览」")
        self.canvas.setAlignment(Qt.AlignCenter)
        self.canvas.setProperty("muted", "true")
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setWidget(self.canvas)
        self.area = area
        root.addWidget(area, 1)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setProperty("muted", "true")
        root.addWidget(self.status)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.reject)
        bottom.addWidget(btn_close)
        root.addLayout(bottom)

        self._check_deps()

    # ---------- 依赖 ----------
    def _check_deps(self):
        """缺组件就直说缺哪个、影响什么，不静默降级"""
        msgs = []
        ok_merge, why_merge = overlay.can_merge()
        ok_render, why_render = overlay.can_render()
        if not ok_merge:
            msgs.append('无法合并 PDF（{}），本功能不可用'.format(why_merge))
        elif not ok_render:
            msgs.append('无法在窗口里渲染（{}），但仍可「导出合并 PDF」'
                        '用看图/阅读器打开核对'.format(why_render))
        self.btn_check.setEnabled(ok_merge)
        if msgs:
            self.status.setText('　'.join(msgs))

    # ---------- 套头纸 ----------
    def _pick(self):
        start = os.path.dirname(self._letterhead) if self._letterhead else ''
        path, _ = QFileDialog.getOpenFileName(
            self, "选择套头纸 PDF", start, "PDF 文件 (*.pdf)")
        if not path:
            return
        self._letterhead = path
        self.path_label.setText(path)
        overprint.save_letterhead(self._template_path, path)
        size = overlay.page_size_cm(path)
        n = overlay.page_count(path)
        self.status.setText(
            '套头纸：{} 页，{}'.format(
                n, '{:.1f}×{:.1f}cm'.format(*size) if size else '尺寸未知')
            + ('　（A4 应为 21.0×29.7cm）' if size and
               (abs(size[0] - 21.0) > 0.3 or abs(size[1] - 29.7) > 0.3) else ''))

    def _clear(self):
        self._letterhead = ''
        self.path_label.setText("（未选择）")
        overprint.save_letterhead(self._template_path, '')
        self.canvas.setPixmap(QPixmap())
        self.canvas.setText("选好套头纸 PDF 后，点「生成对位预览」")
        self.btn_export.setEnabled(False)

    # ---------- 生成 ----------
    def _check(self):
        if not self._letterhead or not os.path.exists(self._letterhead):
            QMessageBox.information(self, "提示", "请先选择套头纸的 PDF 文件。")
            return
        import tempfile
        out = os.path.join(tempfile.mkdtemp(prefix='align_'), '对位预览.pdf')
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            path, notes = overlay.build_alignment_pdf(
                self._template_path, self._values, self._letterhead, out,
                title_shape=self._title_shape, title_lines=self._title_lines)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "生成失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()
        self._merged = path
        self._page = 0
        self.btn_export.setEnabled(True)
        n = overlay.page_count(path)
        self.page_combo.setVisible(n > 1)
        if n > 1:
            self.page_combo.blockSignals(True)
            self.page_combo.clear()
            for i in range(n):
                self.page_combo.addItem('第 {} 页'.format(i + 1), i)
            self.page_combo.blockSignals(False)
        self.status.setText('；'.join(notes) if notes else
                            '已叠加。红色=纸上已印好，黑色=打印机这次要印的。')
        self._redraw()

    def _on_page(self, idx):
        self._page = max(0, idx)
        self._redraw()

    def _redraw(self):
        if not self._merged:
            return
        ok_render, why = overlay.can_render()
        if not ok_render:
            self.canvas.setText(
                '本机无法在窗口里渲染 PDF（{}）。\n'
                '请点「导出合并 PDF…」，用看图软件或 PDF 阅读器打开核对。'.format(why))
            return
        try:
            png = overlay.render_page_png(self._merged, self._page)
        except Exception as e:
            self.canvas.setText('渲染失败：{}'.format(e))
            return
        pix = QPixmap(png)
        try:
            os.remove(png)          # 渲染件是临时文件，读进内存就删
        except OSError:
            pass
        if pix.isNull():
            self.canvas.setText('渲染结果为空')
            return
        zoom = self.zoom.currentData()
        if not zoom:                # 适应宽度
            target = max(200, self.area.viewport().width() - 24)
        else:
            # 渲染固定 150dpi，100% 指的是"按真实纸张尺寸显示"，
            # 换算成屏幕像素要按 96dpi 折一下，否则 100% 会大得离谱
            target = int(pix.width() * zoom * 96.0 / 150.0)
        if target and abs(target - pix.width()) > 2:
            pix = pix.scaledToWidth(target, Qt.SmoothTransformation)
        self.canvas.setPixmap(pix)
        self.canvas.setText('')

    def _export(self):
        if not self._merged:
            return
        stem = os.path.splitext(os.path.basename(self._template_path))[0]
        base = os.path.dirname(self._letterhead) or ''
        default = os.path.join(base, '{}_套头对位.pdf'.format(stem))
        out, _ = QFileDialog.getSaveFileName(self, "导出合并 PDF", default,
                                             "PDF 文件 (*.pdf)")
        if not out:
            return
        import shutil
        try:
            shutil.copyfile(self._merged, out)
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))
            return
        QMessageBox.information(
            self, "已导出",
            "已保存：\n{}\n\n这份 PDF 是「套头纸 + 本次内容」的合成件，"
            "只用于核对位置；真正打印时仍用生成的 docx 打到预印纸上。"
            .format(out))
