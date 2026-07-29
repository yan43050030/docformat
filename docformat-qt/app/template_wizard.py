# -*- coding: utf-8 -*-
"""新建套打模板：在套头底图上点出每个位置 → 生成模板。

流程就四步：选套头 PDF → 在图上拖框标出每处内容 → 生成 → 立刻叠加校验。
拖框拖到哪儿，字就印到哪儿（实测偏差 0.02cm 以内）。
"""
import os

from PyQt5.QtCore import QPoint, QRect, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog, QDoubleSpinBox,
                             QFileDialog, QHBoxLayout, QLabel, QLineEdit,
                             QListWidget, QListWidgetItem, QMessageBox,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts import overlay, overprint
from scripts.template_builder import build_template


class Canvas(QLabel):
    """套头底图 + 已标注的框；在上面拖一个框就发一次信号"""

    boxDrawn = pyqtSignal(float, float)          # 纸面坐标 cm

    def __init__(self, parent=None):
        super(Canvas, self).__init__(parent)
        self.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.setMouseTracking(True)
        self._start = None
        self._cur = None
        self.cm_per_px = 0.0                     # 底图比例尺
        self.marks = []                          # [(x_cm, y_cm, 标签)]

    def mousePressEvent(self, e):
        if self.pixmap() and e.button() == Qt.LeftButton:
            self._start = e.pos()
            self._cur = e.pos()

    def mouseMoveEvent(self, e):
        if self._start is not None:
            self._cur = e.pos()
            self.update()

    def mouseReleaseEvent(self, e):
        if self._start is None or not self.cm_per_px:
            return
        p = self._start
        self._start = self._cur = None
        self.update()
        # 取框的**左上角**：模板生成器用的就是左上角坐标
        self.boxDrawn.emit(p.x() * self.cm_per_px, p.y() * self.cm_per_px)

    def paintEvent(self, e):
        super(Canvas, self).paintEvent(e)
        if not self.pixmap():
            return
        qp = QPainter(self)
        if self.cm_per_px:
            qp.setPen(QPen(QColor('#1F7A4D'), 1))
            for x, y, text in self.marks:
                px, py = x / self.cm_per_px, y / self.cm_per_px
                qp.drawLine(int(px) - 6, int(py), int(px) + 6, int(py))
                qp.drawLine(int(px), int(py) - 6, int(px), int(py) + 6)
                qp.drawText(int(px) + 8, int(py) + 4, text)
        if self._start and self._cur:
            qp.setPen(QPen(QColor('#C0392B'), 1, Qt.DashLine))
            qp.drawRect(QRect(self._start, self._cur).normalized())
        qp.end()


class TemplateWizard(QDialog):
    """选套头 → 标位置 → 生成模板"""

    KINDS = [('预印栏目名（纸上已有，白字占位）', 'label'),
             ('填写位（要打印的内容）', 'field')]

    def __init__(self, parent=None):
        super(TemplateWizard, self).__init__(parent)
        self.setWindowTitle("新建套打模板 — 在套头图上点出各处位置")
        self.resize(1080, 860)
        self._pdf = ''
        self._items = []
        self._page_cm = (21.0, 29.7)

        root = QVBoxLayout(self)
        tip = QLabel(
            "① 选一份<b>套头纸的 PDF</b>（扫描件也行，但要是未缩放的原尺寸）；"
            "② 在图上<b>拖一个框</b>圈住某处内容——取的是框的<b>左上角</b>；"
            "③ 在右边填它是什么：纸上已印好的<b>栏目名</b>就照抄那几个字，"
            "要打印的位置就给它起个<b>字段名</b>（如 标题、密级）；"
            "④ 全部标完点「生成模板」，随后可直接叠加校验、微调。<br>"
            "标错了在右边列表里选中删掉即可。")
        tip.setWordWrap(True)
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        row = QHBoxLayout()
        row.addWidget(QLabel("套头 PDF："))
        self.path_label = QLabel("（未选择）")
        self.path_label.setWordWrap(True)
        row.addWidget(self.path_label, 1)
        b = QPushButton("选择…")
        b.clicked.connect(self._pick)
        row.addWidget(b)
        root.addLayout(row)

        body = QHBoxLayout()
        self.canvas = Canvas()
        self.canvas.boxDrawn.connect(self._on_box)
        area = QScrollArea()
        area.setWidgetResizable(False)
        area.setWidget(self.canvas)
        body.addWidget(area, 3)

        side = QVBoxLayout()
        side.addWidget(QLabel("这一处是："))
        self.kind = QComboBox()
        for t, v in self.KINDS:
            self.kind.addItem(t, v)
        side.addWidget(self.kind)
        side.addWidget(QLabel("栏目名文字 / 字段名："))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("如：紧急程度：  或  标题")
        side.addWidget(self.name_edit)
        f_row = QHBoxLayout()
        f_row.addWidget(QLabel("字号"))
        self.pt = QDoubleSpinBox()
        self.pt.setRange(6.0, 48.0)
        self.pt.setValue(12.0)
        self.pt.setSingleStep(0.5)
        self.pt.setSuffix(" pt")
        f_row.addWidget(self.pt)
        side.addLayout(f_row)
        self.hint = QLabel("在左边图上拖框，标注会加到下面")
        self.hint.setWordWrap(True)
        self.hint.setProperty("muted", "true")
        side.addWidget(self.hint)
        side.addWidget(QLabel("已标注："))
        self.list = QListWidget()
        side.addWidget(self.list, 1)
        del_btn = QPushButton("删除选中")
        del_btn.clicked.connect(self._del)
        side.addWidget(del_btn)
        body.addLayout(side, 2)
        root.addLayout(body, 1)

        bottom = QHBoxLayout()
        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setProperty("muted", "true")
        bottom.addWidget(self.status, 1)
        self.btn_make = QPushButton("生成模板…")
        self.btn_make.setEnabled(False)
        self.btn_make.clicked.connect(self._make)
        bottom.addWidget(self.btn_make)
        close = QPushButton("关闭")
        close.clicked.connect(self.reject)
        bottom.addWidget(close)
        root.addLayout(bottom)

        ok, why = overlay.can_render()
        if not ok:
            self.status.setText('本机无法渲染 PDF（{}），无法在图上标注。'.format(why))

    # ---------- 底图 ----------
    def _pick(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择套头纸 PDF", "", "PDF 文件 (*.pdf)")
        if not path:
            return
        ok, why = overlay.can_render()
        if not ok:
            QMessageBox.information(self, "无法渲染", why)
            return
        try:
            png = overlay.render_page_png(path, 0)
        except Exception as e:
            QMessageBox.warning(self, "读取失败", str(e))
            return
        pix = QPixmap(png)
        try:
            os.remove(png)
        except OSError:
            pass
        if pix.isNull():
            QMessageBox.warning(self, "读取失败", "这份 PDF 渲染不出内容")
            return
        self._pdf = path
        self.path_label.setText(path)
        size = overlay.page_size_cm(path) or (21.0, 29.7)
        self._page_cm = size
        # 比例尺：一个像素等于多少厘米。有了它，图上任意点都能换算成纸面坐标
        self.canvas.cm_per_px = size[0] / pix.width()
        self.canvas.setPixmap(pix)
        self.canvas.resize(pix.size())
        self.canvas.marks = []
        self._items = []
        self.list.clear()
        self.btn_make.setEnabled(False)
        self.status.setText(
            '套头纸 {:.1f}×{:.1f}cm，底图 {}×{} 像素（1 像素 = {:.4f}cm）'
            .format(size[0], size[1], pix.width(), pix.height(),
                    self.canvas.cm_per_px))

    # ---------- 标注 ----------
    def _on_box(self, x_cm, y_cm):
        kind = self.kind.currentData()
        text = self.name_edit.text().strip()
        if not text:
            self.hint.setText('⚠ 先在上面填"栏目名文字/字段名"，再拖框')
            return
        item = {'x': round(x_cm, 2), 'y': round(y_cm, 2),
                'kind': kind, 'pt': self.pt.value()}
        if kind == 'field':
            item['name'] = text
        else:
            item['text'] = text
        self._items.append(item)
        self.canvas.marks.append((item['x'], item['y'], text[:8]))
        self.canvas.update()
        self.list.addItem(QListWidgetItem(
            '{}  ({:.2f}, {:.2f})cm  {:.0f}pt  {}'.format(
                text, item['x'], item['y'], item['pt'],
                '填写位' if kind == 'field' else '预印')))
        self.name_edit.clear()
        self.hint.setText('已标 {} 处'.format(len(self._items)))
        self.btn_make.setEnabled(True)

    def _del(self):
        r = self.list.currentRow()
        if r < 0:
            return
        self.list.takeItem(r)
        del self._items[r]
        del self.canvas.marks[r]
        self.canvas.update()
        self.btn_make.setEnabled(bool(self._items))

    # ---------- 生成 ----------
    def _make(self):
        if not self._items:
            return
        if not any(i.get('kind') == 'field' for i in self._items):
            QMessageBox.information(
                self, "还差填写位",
                "至少要标一个「填写位」，否则这个模板没有可填的内容。")
            return
        d = overprint.user_overprint_dir()
        os.makedirs(d, exist_ok=True)
        out, _ = QFileDialog.getSaveFileName(
            self, "保存套打模板", os.path.join(d, '新建套打模板.docx'),
            "Word 文档 (*.docx)")
        if not out:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            _p, fields = build_template(
                self._items, out,
                page_w_cm=self._page_cm[0], page_h_cm=self._page_cm[1])
            if self._pdf:
                overprint.save_letterhead(out, self._pdf)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "生成失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()
        self.result_path = out
        QMessageBox.information(
            self, "已生成",
            "模板已保存：\n{}\n\n可填字段（{} 个）：{}\n\n"
            "套头 PDF 已一并记住，回到「套打填写」选中这个模板，"
            "填几个字就能用「套头对位校验」看准不准。"
            .format(out, len(fields), '、'.join(fields)))
        self.accept()
