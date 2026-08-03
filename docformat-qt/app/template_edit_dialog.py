# -*- coding: utf-8 -*-
"""套打模板可视化编辑：在纸上点中一个字，就地改它。

以前改模板只有两条路：让开发改脚本重跑，或者自己拿 Word 硬调。前者用户
干不了，后者一改就把制表位、精确行距、白字这些讲究的地方弄坏——套打错
一毫米就废一张纸。这里给第三条路：整张纸按 A4 比例画出来，点哪个字改
哪个字，位置用尺子量出来直接填，讲究的地方由程序守着。

灰字 = 纸上已经印好的（不打印），黑字 = 要打印上去的填写位。
"""
import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QComboBox, QDialog, QDoubleSpinBox,
                             QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
                             QInputDialog, QLabel, QLineEdit, QMessageBox,
                             QPushButton, QScrollArea, QSplitter, QVBoxLayout,
                             QWidget)

from app.overprint_canvas import OverprintCanvas
from scripts import overprint
from scripts.tpl_edit import (PT_LABELS, EditSession, TemplateEditError,
                             pt_label)

CN_FONTS = ['方正楷体_GBK', '方正仿宋_GBK', '方正小标宋_GBK', '方正大标宋简体',
            '方正黑体_GBK', '方正书宋_GBK', '仿宋_GB2312', '楷体_GB2312',
            '黑体', '宋体', '华文中宋']
EN_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Cambria']


class TemplateEditDialog(QDialog):

    STEP_CM = 0.05

    def __init__(self, template_path, parent=None):
        super(TemplateEditDialog, self).__init__(parent)
        self.setWindowTitle('可视化编辑模板 — {}'.format(
            os.path.splitext(os.path.basename(template_path))[0]))
        self.resize(1180, 820)
        self.saved_path = None
        self._sess = EditSession(template_path)
        self._ref = None
        self._loading = False       # 正在往控件里灌值，别把它当成用户改动

        root = QVBoxLayout(self)
        tip = QLabel(
            '<b>灰字</b>是纸上已经预印好的内容（软件里存成白字，占位但不打印）；'
            '<b>黑字</b>是要打印上去的填写位 <code>{{字段名}}</code>。<br>'
            '点中纸上任意一个字即可编辑：拖动可挪位置，右侧可改内容、字体、'
            '字号和精确到厘米的坐标。')
        tip.setWordWrap(True)
        tip.setProperty('muted', 'true')
        root.addWidget(tip)

        split = QSplitter(Qt.Horizontal)

        # ---------------- 左：画布 ----------------
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        bar = QHBoxLayout()
        bar.addWidget(QLabel('显示：'))
        self.zoom = QComboBox()
        for t, z in (('适应窗口', 0), ('75%', 28.3), ('100%（等大）', 37.8),
                     ('150%', 56.7), ('200%', 75.6)):
            self.zoom.addItem(t, z)
        self.zoom.currentIndexChanged.connect(self._apply_zoom)
        bar.addWidget(self.zoom)
        self.btn_undo = QPushButton('撤销')
        self.btn_undo.clicked.connect(self._undo)
        bar.addWidget(self.btn_undo)
        self.btn_redo = QPushButton('重做')
        self.btn_redo.clicked.connect(self._redo)
        bar.addWidget(self.btn_redo)
        bar.addStretch(1)
        ll.addLayout(bar)

        self.canvas = OverprintCanvas()
        self.canvas.set_edit(True)
        self.canvas.refPicked.connect(self._on_pick)
        self.canvas.refMoved.connect(self._on_moved)
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setWidget(self.canvas)
        ll.addWidget(area, 1)
        split.addWidget(left)

        # ---------------- 右：属性 ----------------
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)

        self.box = QGroupBox('选中的内容')
        form = QFormLayout(self.box)
        form.setLabelAlignment(Qt.AlignRight)

        self.ed_text = QLineEdit()
        self.ed_text.setToolTip('预印内容直接写文字；填写位写成 {{字段名}}')
        self.ed_text.editingFinished.connect(self._apply_text)
        form.addRow('文字：', self.ed_text)

        self.cb_kind = QComboBox()
        self.cb_kind.addItem('预印（纸上已有，不打印）', True)
        self.cb_kind.addItem('填写位（要打印出来）', False)
        self.cb_kind.setToolTip(
            '套打的全部秘密就是这一条：预印内容存成白字，占准位置但印不出来。')
        self.cb_kind.currentIndexChanged.connect(self._apply_kind)
        form.addRow('类型：', self.cb_kind)

        self.cb_cn = QComboBox()
        self.cb_cn.setEditable(True)
        self.cb_cn.addItems(CN_FONTS)
        self.cb_cn.currentTextChanged.connect(self._apply_font)
        form.addRow('中文字体：', self.cb_cn)

        self.cb_en = QComboBox()
        self.cb_en.setEditable(True)
        self.cb_en.addItems(EN_FONTS)
        self.cb_en.currentTextChanged.connect(self._apply_font)
        form.addRow('西文/数字：', self.cb_en)

        size_row = QHBoxLayout()
        self.cb_pt = QComboBox()
        for v, _name in PT_LABELS:
            self.cb_pt.addItem(pt_label(v), v)
        self.cb_pt.currentIndexChanged.connect(self._apply_font)
        size_row.addWidget(self.cb_pt, 1)
        self.ck_bold = QCheckBox('加粗')
        self.ck_bold.stateChanged.connect(self._apply_font)
        size_row.addWidget(self.ck_bold)
        w = QWidget()
        w.setLayout(size_row)
        form.addRow('字号：', w)

        self.sp_x = self._spin(0.0, 21.0, '距纸张左边缘的厘米数')
        self.sp_x.editingFinished.connect(self._apply_x)
        form.addRow('距纸左边：', self._with_steps(self.sp_x, self._apply_x))

        self.sp_y = self._spin(0.0, 29.7, '距纸张上边缘的厘米数')
        self.sp_y.editingFinished.connect(self._apply_y)
        form.addRow('距纸上边：', self._with_steps(self.sp_y, self._apply_y))

        self.sp_w = self._spin(0.0, 21.0,
                               '这段字整体占多宽（靠字距收放）。\n'
                               '预印栏目名在纸上多是收着排的，宽度不对，\n'
                               '紧跟其后的黑字就整体偏。')
        self.sp_w.editingFinished.connect(self._apply_w)
        form.addRow('这段字总宽：', self._with_steps(self.sp_w, self._apply_w))

        rl.addWidget(self.box)

        act = QHBoxLayout()
        self.btn_del = QPushButton('删除这段')
        self.btn_del.clicked.connect(self._delete)
        act.addWidget(self.btn_del)
        act.addStretch(1)
        rl.addLayout(act)

        addbox = QGroupBox('新增')
        al = QVBoxLayout(addbox)
        hint = QLabel('新内容加在**当前选中那一行**的末尾，加好之后拖到位、'
                      '或直接填坐标。没选中就加不了——得先告诉软件加到哪一行。')
        hint.setWordWrap(True)
        hint.setProperty('muted', 'true')
        al.addWidget(hint)
        arow = QHBoxLayout()
        b1 = QPushButton('加预印栏目名…')
        b1.clicked.connect(lambda: self._add(False))
        arow.addWidget(b1)
        b2 = QPushButton('加填写位…')
        b2.clicked.connect(lambda: self._add(True))
        arow.addWidget(b2)
        al.addLayout(arow)
        rl.addWidget(addbox)

        self.info = QLabel('')
        self.info.setWordWrap(True)
        self.info.setProperty('muted', 'true')
        rl.addWidget(self.info)
        rl.addStretch(1)
        split.addWidget(right)
        split.setSizes([740, 420])
        root.addWidget(split, 1)

        self.status = QLabel('')
        self.status.setWordWrap(True)
        self.status.setProperty('muted', 'true')
        root.addWidget(self.status)

        btns = QHBoxLayout()
        btns.addStretch(1)
        b = QPushButton('取消')
        b.clicked.connect(self.reject)
        btns.addWidget(b)
        b = QPushButton('另存为…')
        b.clicked.connect(self._save_as)
        btns.addWidget(b)
        b = QPushButton('保存')
        b.setProperty('primary', 'true')
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(self._save)
        btns.addWidget(b)
        root.addLayout(btns)

        self._refresh()

    # ---------------- 小控件 ----------------
    def _spin(self, lo, hi, tip):
        sp = QDoubleSpinBox()
        sp.setRange(lo, hi)
        sp.setDecimals(2)
        sp.setSingleStep(self.STEP_CM)
        sp.setSuffix(' cm')
        sp.setToolTip(tip)
        return sp

    def _with_steps(self, spin, apply_fn):
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        for txt, d in (('－', -self.STEP_CM), ('＋', self.STEP_CM)):
            b = QPushButton(txt)
            b.setFixedWidth(30)
            b.clicked.connect(
                lambda _c=False, _s=spin, _d=d, _f=apply_fn: (
                    _s.setValue(max(_s.minimum(), _s.value() + _d)), _f()))
            row.addWidget(b)
        row.addWidget(spin, 1)
        w = QWidget()
        w.setLayout(row)
        return w

    # ---------------- 刷新 ----------------
    def _apply_zoom(self):
        self.canvas.set_zoom(self.zoom.currentData() or 0)

    def _refresh(self, keep=True):
        plan = self._sess.outline()
        self.canvas.set_plan(plan)
        self.canvas.select(self._ref if keep else None)
        self.btn_undo.setEnabled(self._sess.can_undo())
        self.btn_redo.setEnabled(self._sess.can_redo())
        self._load_props()

    def _load_props(self):
        on = self._ref is not None
        self.box.setEnabled(on)
        self.btn_del.setEnabled(on)
        if not on:
            self.info.setText('还没选中内容。点纸上任意一个字试试。')
            return
        try:
            d = self._sess.info(self._ref)
            pos = self._sess.positions().get(tuple(self._ref))
        except TemplateEditError as exc:
            self._ref = None
            self.status.setText(str(exc))
            self.box.setEnabled(False)
            return
        self._loading = True
        try:
            self.ed_text.setText(d['text'])
            self.cb_kind.setCurrentIndex(0 if d['white'] else 1)
            self.cb_cn.setEditText(d['font_cn'])
            self.cb_en.setEditText(d['font_en'])
            i = self.cb_pt.findData(round(d['pt'], 2))
            if i < 0:
                self.cb_pt.addItem(pt_label(d['pt']), d['pt'])
                i = self.cb_pt.count() - 1
            self.cb_pt.setCurrentIndex(i)
            self.ck_bold.setChecked(d['bold'])
            if pos:
                self.sp_x.setValue(round(pos[0], 2))
                self.sp_y.setValue(round(pos[1], 2))
            self.sp_w.setValue(round(d['width_cm'], 2))
        finally:
            self._loading = False
        self.info.setText(
            '字距 {:+.3f}cm　{}'.format(
                d['track_cm'],
                '这是填写位，生成时会换成用户填的内容' if d['is_field']
                else '这是预印内容，只占位、不打印'))

    # ---------------- 画布回调 ----------------
    def _on_pick(self, ref):
        self._ref = tuple(ref) if ref is not None else None
        self._load_props()

    def _on_moved(self, ref, dx, dy):
        if ref is None:
            return
        self._ref = tuple(ref)
        msg = []
        try:
            if abs(dx) > 0.005:
                self._sess.nudge_x(self._ref, dx)
            if abs(dy) > 0.005:
                pi = self._ref[0]
                got = self._sess.nudge_y(pi, dy)
                if abs(got - dy) > 0.005:
                    # 段前距不能为负，往上顶到头了。如实说清楚，
                    # 别让画布显示的位置和纸上的对不上
                    msg.append('往上只能挪到 {:.2f}cm 就顶住了'
                               '（Word 的段前距不能为负）'.format(abs(got)))
        except TemplateEditError as exc:
            msg.append(str(exc))
        self.status.setText('；'.join(msg))
        self._refresh()

    # ---------------- 编辑动作 ----------------
    def _guard(self):
        return self._ref is not None and not self._loading

    def _apply_text(self):
        if not self._guard():
            return
        new = self.ed_text.text()
        try:
            if new == self._sess.info(self._ref)['text']:
                return
            self._sess.set_text(self._ref, new)
        except TemplateEditError as exc:
            self.status.setText(str(exc))
            return
        self.status.setText('')
        self._refresh()

    def _apply_kind(self):
        if not self._guard():
            return
        self._sess.set_style(self._ref, white=bool(self.cb_kind.currentData()))
        self._refresh()

    def _apply_font(self):
        if not self._guard():
            return
        self._sess.set_style(self._ref,
                             font_cn=self.cb_cn.currentText().strip(),
                             font_en=self.cb_en.currentText().strip(),
                             pt=self.cb_pt.currentData(),
                             bold=self.ck_bold.isChecked())
        self._refresh()

    def _apply_x(self):
        if not self._guard():
            return
        cur = self._sess.positions().get(tuple(self._ref))
        if not cur:
            return
        d = round(self.sp_x.value() - cur[0], 3)
        if abs(d) < 0.005:
            return
        try:
            self._sess.nudge_x(self._ref, d)
        except TemplateEditError as exc:
            self.status.setText(str(exc))
            return
        self._refresh()

    def _apply_y(self):
        if not self._guard():
            return
        cur = self._sess.positions().get(tuple(self._ref))
        if not cur:
            return
        d = round(self.sp_y.value() - cur[1], 3)
        if abs(d) < 0.005:
            return
        got = self._sess.nudge_y(self._ref[0], d)
        if abs(got - d) > 0.005:
            self.status.setText(
                '往上只能挪 {:.2f}cm 就顶住了（Word 的段前距不能为负）；'
                '要再往上，得先把上一行往上挪。'.format(abs(got)))
        else:
            self.status.setText('')
        self._refresh()

    def _apply_w(self):
        if not self._guard():
            return
        self._sess.set_width(self._ref, self.sp_w.value())
        self._refresh()

    def _delete(self):
        if self._ref is None:
            return
        try:
            txt = self._sess.info(self._ref)['text']
        except TemplateEditError:
            return
        if QMessageBox.question(
                self, '删除', '确定删掉「{}」？'.format(txt),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) != QMessageBox.Yes:
            return
        self._sess.delete(self._ref)
        self._ref = None
        self.status.setText('已删除「{}」，可点「撤销」还原'.format(txt))
        self._refresh(keep=False)

    def _add(self, is_field):
        if self._ref is None:
            QMessageBox.information(
                self, '先选一行',
                '请先在纸上点中一个字，新内容会加在它所在的那一行末尾。')
            return
        pi = self._ref[0]
        if is_field:
            name, ok = QInputDialog.getText(
                self, '加填写位', '字段名（生成时按它向用户要内容）：')
            if not ok or not name.strip():
                return
            try:
                ref = self._sess.add_field(pi, name.strip(),
                                           x_cm=self.sp_x.value() + 1.0)
            except TemplateEditError as exc:
                QMessageBox.warning(self, '加不了', str(exc))
                return
        else:
            text, ok = QInputDialog.getText(
                self, '加预印栏目名', '文字（纸上已印好的内容，不会打印）：')
            if not ok or not text.strip():
                return
            ref = self._sess.add_run(pi, text.strip(),
                                     x_cm=self.sp_x.value() + 1.0, white=True)
        self._ref = ref
        self.status.setText('已添加，拖到位或直接填坐标')
        self._refresh()

    def _undo(self):
        if self._sess.undo():
            self._ref = None
            self.status.setText('已撤销')
            self._refresh(keep=False)

    def _redo(self):
        if self._sess.redo():
            self._ref = None
            self._refresh(keep=False)

    # ---------------- 存盘 ----------------
    def _check(self):
        """存之前替用户看一眼：有没有内容跑出纸外、字段名写坏了。"""
        warn = []
        plan = self._sess.outline()
        pg = plan['page']
        for seg, x, y in _walk(plan):
            txt = seg.get('text', '')
            if x < 0 or x > pg['width_cm'] or y < 0 or y > pg['height_cm']:
                warn.append('「{}」跑到纸外面了（{:.2f}, {:.2f}）'.format(
                    txt[:8], x, y))
            # 大括号只配了一半：生成时既换不掉、也印不出，属于改坏了
            if ('{{' in txt) != ('}}' in txt):
                warn.append('「{}」的 {{{{}}}} 括号不成对，这一栏不会被当成填写位'
                            .format(txt[:12]))
        return warn

    def _save(self):
        self._do_save(self._sess.path)

    def _save_as(self):
        d = overprint.user_overprint_dir()
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = ''
        path, _ = QFileDialog.getSaveFileName(
            self, '另存模板为', os.path.join(d, os.path.basename(
                self._sess.path)), 'Word 文档 (*.docx)')
        if path:
            self._do_save(path)

    def _do_save(self, path):
        warn = self._check()
        if warn:
            ret = QMessageBox.question(
                self, '有内容跑出纸外',
                '\n'.join(warn[:8]) + '\n\n仍然保存？',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return
        try:
            self.saved_path = self._sess.save(path)
        except Exception as exc:
            QMessageBox.warning(self, '保存失败', str(exc))
            return
        QMessageBox.information(
            self, '已保存',
            '模板已保存到：\n{}\n\n回到套打窗口填一份试试，'
            '再打一张空白纸比对，位置不对就回来接着调。'.format(self.saved_path))
        self.accept()

    def reject(self):
        if self._sess.dirty:
            ret = QMessageBox.question(
                self, '放弃修改', '模板改了还没保存，确定放弃？',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return
        super(TemplateEditDialog, self).reject()


def _walk(plan):
    return overprint.iter_seg_positions(plan, with_y=True)
