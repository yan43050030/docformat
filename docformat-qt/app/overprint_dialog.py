# -*- coding: utf-8 -*-
"""套打填写：选套打模板 → 填字段 → 生成可直接打到预印纸上的文件"""
import os

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QDialog,
                             QDialogButtonBox,
                             QDoubleSpinBox, QFileDialog, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                             QPlainTextEdit, QPushButton, QScrollArea,
                             QSplitter, QTextBrowser, QVBoxLayout, QWidget)

from app.theme import settings
from app.overprint_canvas import OverprintCanvas
from scripts import overprint

# 这些字段内容通常较长，用多行输入框
_LONG_FIELDS = {'拟办意见', '领导批示', '备注', '主要内容', '说明'}
# 这些字段每次都变，不做记忆
_NO_MEMORY = {'标题', '拟办意见', '领导批示', '备注', '年', '月', '日'}


class OffsetDialog(QDialog):
    """打印位置微调：把每个字段顶到距纸张左边指定的厘米数上。

    套打准不准，只取决于黑字落在哪儿。这里让用户拿尺子量真实预印单，
    把数值填进来，存成模板旁边的 .位置.json，跟着模板走。
    """

    STEP_CM = 0.05      # 微调步进；定位靠制表位，精度不受此限

    def __init__(self, template_path, fields, current, parent=None):
        super(OffsetDialog, self).__init__(parent)
        self.setWindowTitle("打印位置微调")
        self.resize(520, 520)
        self._template_path = template_path
        self._spins = {}
        self._spins_y = {}
        saved = overprint.load_offsets(template_path)
        saved_y = overprint.load_offsets_y(template_path)

        root = QVBoxLayout(self)
        tip = QLabel(
            "<b>怎么量、怎么填</b><br>"
            "① 拿一张真实的预印单，用直尺从<b>纸张左边缘</b>量到该栏空格的<b>左沿</b>；<br>"
            "② 把量到的<b>厘米数</b>填进下面对应的格子（例如 3.20）；<br>"
            "③ 生成、试打一张，对不准再回来微调，直到套准为止。<br><br>"
            "单位一律是 <b>厘米(cm)</b>，从纸张左边缘算起（含页边距），"
            "不是从版心或表格边算起。<br>"
            "留空（0.00）表示不指定，按模板原样排。<br>"
            "填多少就印在多少——定位用的是 Word 制表位，与字体无关，"
            "实测误差小于 0.01cm。<br>"
            "◀ ▶ 每次微调 {:.2f}cm。若某栏的预印栏目名本身就压过了你填的位置，"
            "保存后会明确提示“顶不过去”并告诉你最小可用值。".format(self.STEP_CM))
        tip.setWordWrap(True)
        tip.setProperty("muted", "true")
        root.addWidget(tip)

        host = QWidget()
        form = QFormLayout(host)
        form.setLabelAlignment(Qt.AlignRight)
        for name in fields:
            row = QHBoxLayout()
            sp = QDoubleSpinBox()
            sp.setRange(0.0, 21.0)
            sp.setDecimals(2)
            sp.setSingleStep(self.STEP_CM)
            sp.setSuffix(" cm")
            sp.setSpecialValueText("不指定")     # 0.00 显示成"不指定"
            sp.setValue(float(saved.get(name, 0.0)))
            sp.setToolTip("距纸张左边缘的厘米数；0 = 不指定，按模板原样。\n"
                          "键盘上下箭头也能微调")
            self._spins[name] = sp
            # 左右各一个微调键：对着实物一点一点挪，比反复键入数字快
            for _txt, _d in (('◀', -self.STEP_CM), ('▶', self.STEP_CM)):
                b = QPushButton(_txt)
                b.setFixedWidth(28)
                b.setToolTip('向{}移 {:.2f}cm'.format('左' if _d < 0 else '右',
                                                    abs(_d)))
                b.clicked.connect(
                    lambda _c=False, _s=sp, _dd=_d: _s.setValue(
                        max(0.0, (_s.value() or 0.0) + _dd)))
                row.addWidget(b)
            row.addWidget(sp)
            cur = current.get(name)
            if cur is not None:
                b2 = QPushButton("取当前 {:.2f}".format(cur))
                b2.setToolTip("把这一栏现在的实际落点填进去，再在此基础上微调")
                b2.clicked.connect(
                    lambda _c=False, _s=sp, _v=float(cur): _s.setValue(_v))
                row.addWidget(b2)
            row.addWidget(QLabel("　距上边"))
            spy = QDoubleSpinBox()
            spy.setRange(0.0, 29.7)
            spy.setDecimals(2)
            spy.setSingleStep(self.STEP_CM)
            spy.setSuffix(" cm")
            spy.setSpecialValueText("不指定")
            spy.setValue(float(saved_y.get(name, 0.0)))
            spy.setToolTip("距纸张上边缘的厘米数；0 = 不指定。\n"
                           "纵向只能**整行**挪（Word 里只有段前距能推一行），\n"
                           "同一行上的几个字段会一起动；表格格子里的字段挪不了，\n"
                           "行高是模板定死的。保存后若挪不动会明确告诉你。")
            self._spins_y[name] = spy
            row.addWidget(spy)
            lab = QLabel("")
            lab.setProperty("muted", "true")
            row.addWidget(lab)
            row.addStretch(1)
            w = QWidget()
            w.setLayout(row)
            form.addRow(name + "：", w)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(host)
        root.addWidget(scroll, 1)

        # ---- 整体平移：整张纸一起挪 ----
        shift_row = QHBoxLayout()
        shift_row.addWidget(QLabel("整体平移："))
        self._shift = {}
        dx0, dy0 = overprint.load_shift(template_path)
        for key, text, init in (('dx', '向右', dx0), ('dy', '向下', dy0)):
            shift_row.addWidget(QLabel(text))
            sp = QDoubleSpinBox()
            sp.setRange(-5.0, 5.0)
            sp.setDecimals(2)
            sp.setSingleStep(self.STEP_CM)
            sp.setSuffix(" cm")
            sp.setValue(float(init))
            sp.setToolTip("整张纸的内容一起挪，用来补打印机走纸、纸张裁切的整体误差。\n"
                          "负数就是往反方向挪。")
            self._shift[key] = sp
            shift_row.addWidget(sp)
        self._auto_btn = QPushButton("按扫描件自动对位…")
        self._auto_btn.setToolTip("拿一张**空白套头纸**的扫描件（原尺寸 100%，别裁边），\n"
                                  "程序会找出纸上的红线，和模板的框线一比，算出该挪多少。")
        self._auto_btn.clicked.connect(self._auto_align)
        shift_row.addWidget(self._auto_btn)
        shift_row.addStretch(1)
        root.addLayout(shift_row)

        path_lab = QLabel("保存到：{}".format(overprint.offsets_path(template_path)))
        path_lab.setWordWrap(True)
        path_lab.setProperty("muted", "true")
        root.addWidget(path_lab)

        bb = QDialogButtonBox()
        btn_reset = bb.addButton("恢复默认", QDialogButtonBox.ResetRole)
        bb.addButton("取消", QDialogButtonBox.RejectRole)
        bb.addButton("保存", QDialogButtonBox.AcceptRole)
        btn_reset.clicked.connect(self._reset)
        bb.accepted.connect(self._save)
        bb.rejected.connect(self.reject)
        root.addWidget(bb)

    def _reset(self):
        for sp in list(self._spins.values()) + list(self._spins_y.values()):
            sp.setValue(0.0)
        for sp in self._shift.values():
            sp.setValue(0.0)

    def values(self):
        return {k: sp.value() for k, sp in self._spins.items() if sp.value() > 0}

    def values_y(self):
        return {k: sp.value() for k, sp in self._spins_y.items() if sp.value() > 0}

    def shift(self):
        return (self._shift['dx'].value(), self._shift['dy'].value())

    def _auto_align(self):
        from scripts import scan_align
        ok, why = scan_align.available()
        if not ok:
            QMessageBox.information(self, "用不了自动对位", why)
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择空白套头纸的扫描件（PDF）", "", "PDF 文件 (*.pdf)")
        if not path:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            res = scan_align.align(path, self._template_path)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "量不出来", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
        text = scan_align.describe(res)
        if res.get('dx') is None and res.get('dy') is None:
            QMessageBox.information(self, "没找到基准线", text)
            return
        ret = QMessageBox.question(
            self, "自动对位结果",
            text + "\n\n把这个平移量填进去吗？（填完还要点“保存”才生效）",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if ret != QMessageBox.Yes:
            return
        # 量出来的是"模板要往哪儿挪"，正是平移量本身
        if res.get('dx') is not None:
            self._shift['dx'].setValue(round(res['dx'], 2))
        if res.get('dy') is not None:
            self._shift['dy'].setValue(round(res['dy'], 2))

    def _save(self):
        overprint.save_offsets(self._template_path, self.values(),
                               shift=self.shift(), offsets_y=self.values_y())
        self.accept()


class OverprintDialog(QDialog):
    def __init__(self, parent=None):
        super(OverprintDialog, self).__init__(parent)
        self.setWindowTitle("套打填写 — 打到预印红头纸上")
        self.resize(1120, 720)
        self._editors = {}
        self._template_path = None
        # 拖动改出来的位置先放内存里，随预览即时生效；点「保存位置」才落盘。
        # 拖的时候不停写文件既慢又难反悔
        self._offsets = {}
        self._offsets_y = {}
        self._shift = (0.0, 0.0)
        self._pos_dirty = False
        self._bg_cache = {}
        # 导入内容的源文件目录——生成时默认存到那里，省得每次翻文件夹
        self._source_dir = None
        # 输入后延迟重算预览，避免每敲一个字就跑一遍填充
        self._pv_timer = QTimer(self)
        self._pv_timer.setSingleShot(True)
        self._pv_timer.setInterval(350)
        self._pv_timer.timeout.connect(self._refresh_preview)

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
        new_btn = QPushButton("新建模板…")
        new_btn.setCursor(Qt.PointingHandCursor)
        new_btn.setToolTip("拿一份套头纸的 PDF，在图上点出各处位置，直接做出新模板")
        new_btn.clicked.connect(self._new_template)
        row.addWidget(new_btn)
        add_btn = QPushButton("添加模板…")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setToolTip("导入自己的套打模板 docx（含 {{字段名}} 占位符）")
        add_btn.clicked.connect(self._import_template)
        row.addWidget(add_btn)
        self.edit_btn = QPushButton("修改模板…")
        self.edit_btn.setCursor(Qt.PointingHandCursor)
        self.edit_btn.setToolTip(
            "在 Word/WPS 里打开模板修改，比如把白色的单位名改成你的真实单位名。\n"
            "自带模板随软件安装、不可直接改，会先复制一份到你的模板目录再打开。")
        self.edit_btn.clicked.connect(self._edit_template)
        row.addWidget(self.edit_btn)
        batch_btn = QPushButton("批量套打…")
        batch_btn.setToolTip("从 Excel/CSV 读一批数据，一次生成一叠套打件")
        batch_btn.clicked.connect(self._batch)
        row.addWidget(batch_btn)
        lib_btn = QPushButton("套头库…")
        lib_btn.setToolTip("几套红头纸集中管一处；还能自动认出哪张纸配这个模板")
        lib_btn.clicked.connect(self._letterhead_lib)
        row.addWidget(lib_btn)
        open_dir = QPushButton("模板目录")
        open_dir.setCursor(Qt.PointingHandCursor)
        open_dir.setToolTip("打开存放套打模板的文件夹")
        open_dir.clicked.connect(self._open_template_dir)
        row.addWidget(open_dir)
        root.addLayout(row)

        src_row = QHBoxLayout()
        pick = QPushButton("从 docx 导入内容…")
        pick.setCursor(Qt.PointingHandCursor)
        pick.setToolTip("选一份已有的送审单/草稿 docx，自动识别各部分内容填进下面的字段；\n"
                        "日期会拆成年/月/日分别落位，识别不到的可手工补填")
        pick.clicked.connect(self._import_content)
        src_row.addWidget(pick)
        drop_hint = QLabel("（也可把 docx 直接拖到本窗口）")
        drop_hint.setProperty("muted", "true")
        src_row.addWidget(drop_hint)
        clear = QPushButton("清空")
        clear.setProperty("flat", "true")
        clear.setCursor(Qt.PointingHandCursor)
        clear.clicked.connect(self._clear_fields)
        src_row.addWidget(clear)
        src_row.addStretch(1)
        root.addLayout(src_row)

        split = QSplitter(Qt.Horizontal)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        split.addWidget(self.scroll)

        pv_box = QWidget()
        pv_lay = QVBoxLayout(pv_box)
        pv_lay.setContentsMargins(6, 0, 0, 0)
        pv_lay.setSpacing(4)
        pv_head = QLabel("版面预览（灰字=纸上已预印，不会打印；黑字=本次打印内容）")
        pv_head.setProperty("sectionTitle", "true")
        pv_head.setWordWrap(True)
        pv_lay.addWidget(pv_head)

        shape_row = QHBoxLayout()
        shape_row.addWidget(QLabel("长标题回行："))
        self.shape_combo = QComboBox()
        for _t, _v in (('正梯形（上长下短）', 'trapezoid_down'),
                       ('倒梯形（上短下长）', 'trapezoid_up'),
                       ('不回行（由 Word 自动折行）', 'none')):
            self.shape_combo.addItem(_t, _v)
        self.shape_combo.setToolTip(
            "标题一行放不下时的分行方式。公文要求词意完整、排列对称；\n"
            "选“不回行”则交给 Word 自动折行，可能把词拆断。")
        self.shape_combo.currentIndexChanged.connect(self._schedule_preview)
        shape_row.addWidget(self.shape_combo)
        shape_row.addWidget(QLabel("分"))
        self.lines_combo = QComboBox()
        self.lines_combo.addItem("自动", None)
        # 不放"1 行"：一行放不下时 Word 照样会折，标不了它做不到的承诺；
        # "只占一行、不主动断"由左边的"不回行"表达
        for _n in (2, 3, 4):
            self.lines_combo.addItem("{} 行".format(_n), _n)
        self.lines_combo.setToolTip(
            "标题分成几行。选“自动”按长度决定；指定行数则一定分成那么多行。\n"
            "想精确控制断在哪个字，直接在左边标题框里按回车。")
        self.lines_combo.currentIndexChanged.connect(self._schedule_preview)
        shape_row.addWidget(self.lines_combo)
        shape_row.addStretch(1)
        self.btn_save_pos = QPushButton("保存位置")
        self.btn_save_pos.setEnabled(False)
        self.btn_save_pos.setToolTip("把拖出来的位置记到模板旁边，下次打开还是这样")
        self.btn_save_pos.clicked.connect(self._save_positions)
        shape_row.addWidget(self.btn_save_pos)
        self.btn_reset_pos = QPushButton("还原位置")
        self.btn_reset_pos.setToolTip("撤掉所有微调，回到模板原样（还要点「保存位置」才落盘）")
        self.btn_reset_pos.clicked.connect(self._reset_positions)
        shape_row.addWidget(self.btn_reset_pos)
        self.btn_offsets = QPushButton("打印位置微调…")
        self.btn_offsets.setToolTip(
            "拿尺子量真实预印单，指定每个字段距纸张左边缘多少厘米，\n"
            "试打几次调准后就固定下来（存在模板旁边，跟着模板走）")
        self.btn_offsets.clicked.connect(self._edit_offsets)
        shape_row.addWidget(self.btn_offsets)
        self.btn_align = QPushButton("套头对位校验…")
        self.btn_align.setToolTip(
            "选一份套头纸（红头文件纸）的 PDF，把本次内容叠上去看：\n"
            "红色是纸上已印好的，黑色是打印机要印的——不用试打就能看准不准")
        self.btn_align.clicked.connect(self._check_align)
        shape_row.addWidget(self.btn_align)
        pv_lay.addLayout(shape_row)
        self.shape_hint = QLabel("")
        self.shape_hint.setProperty("muted", "true")
        self.shape_hint.setWordWrap(True)
        pv_lay.addWidget(self.shape_hint)
        # 预览画布：按真实 A4 比例画整张纸，黑字可直接拖着挪位置
        zoom_row = QHBoxLayout()
        zoom_row.addWidget(QLabel("显示："))
        self.zoom_combo = QComboBox()
        for _t, _z in (('适应窗口', 0), ('50%', 18.9), ('75%', 28.3),
                       ('100%（等大）', 37.8), ('150%', 56.7), ('200%', 75.6)):
            self.zoom_combo.addItem(_t, _z)
        self.zoom_combo.setToolTip("100% 就是纸上的实际大小（按 96dpi 折算）")
        self.zoom_combo.currentIndexChanged.connect(self._apply_zoom)
        zoom_row.addWidget(self.zoom_combo)
        self.bg_check = QCheckBox("垫上套头纸")
        self.bg_check.setToolTip("把绑定的套头 PDF 垫在底下，直接看黑字有没有落进预印栏位")
        self.bg_check.stateChanged.connect(self._toggle_bg)
        zoom_row.addWidget(self.bg_check)
        zoom_row.addStretch(1)
        self.drag_hint = QLabel("拖黑字改它的横向位置，拖空白处整张纸一起挪")
        self.drag_hint.setProperty("muted", "true")
        zoom_row.addWidget(self.drag_hint)
        pv_lay.addLayout(zoom_row)

        self.canvas = OverprintCanvas()
        self.canvas.fieldMoved.connect(self._on_field_moved)
        self.canvas.sheetMoved.connect(self._on_sheet_moved)
        self.pv_area = QScrollArea()
        self.pv_area.setWidgetResizable(True)
        self.pv_area.setWidget(self.canvas)
        pv_lay.addWidget(self.pv_area, 1)
        self.pv_note = QLabel("")
        self.pv_note.setProperty("muted", "true")
        self.pv_note.setWordWrap(True)
        pv_lay.addWidget(self.pv_note)
        split.addWidget(pv_box)
        split.setSizes([460, 640])
        root.addWidget(split, 1)

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

        # 支持把 docx 直接拖进窗口：拖模板→添加为模板，拖普通文档→导入内容
        self.setAcceptDrops(True)

        if not self._templates:
            self.status.setText("未找到套打模板。点「添加模板…」导入一份含 {{字段名}} 的 docx。")
        else:
            self._load_fields()

    # ---------- 拖拽 ----------
    @staticmethod
    def _dropped_docx(mime):
        if not mime.hasUrls():
            return None
        for url in mime.urls():
            path = url.toLocalFile()
            if path and path.lower().endswith('.docx') and os.path.isfile(path):
                return path
        return None

    def dragEnterEvent(self, event):
        if self._dropped_docx(event.mimeData()):
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if self._dropped_docx(event.mimeData()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        path = self._dropped_docx(event.mimeData())
        if not path:
            return
        event.acceptProposedAction()
        # 带 {{字段}} 的是套打模板，否则当作要适配的内容文档
        try:
            has_fields = bool(overprint.scan_fields(path))
        except Exception:
            has_fields = False
        if has_fields:
            ret = QMessageBox.question(
                self, "拖入的是套打模板",
                "这份 docx 含 {{字段名}} 占位符，看起来是套打模板。\n\n"
                "添加为模板？选「否」则按内容文档导入。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if ret == QMessageBox.Yes:
                self._add_template_file(path)
                return
        self._load_content_from(path)

    # ---------- 字段 ----------
    HISTORY_MAX = 8

    def _history(self, name):
        raw = settings().value('overprint_hist/{}'.format(name), '') or ''
        return [x for x in str(raw).split('\x1f') if x.strip()]

    def _remember(self, name, value):
        """把这次填的值记进历史，最近用的排最前"""
        value = (value or '').strip()
        if not value or name in _NO_MEMORY:
            return
        hist = [value] + [h for h in self._history(name) if h != value]
        settings().setValue('overprint_hist/{}'.format(name),
                            '\x1f'.join(hist[:self.HISTORY_MAX]))

    def _preflight_ok(self, values):
        """生成前先预检。套打纸是预印的，废一张少一张，
        别等打出来才发现内容压到栏外。"""
        plan = getattr(self, '_last_plan', None)
        if plan is None:
            return True
        try:
            items = overprint.preflight(plan, values, self._offsets)
        except Exception:
            return True
        bad = [m for lv, m in items if lv == 'block']
        warn = [m for lv, m in items if lv == 'warn']
        if not bad and not warn:
            return True
        text = ''
        if bad:
            text += '这几处印出来一定不对：\n  · ' + '\n  · '.join(bad) + '\n\n'
        if warn:
            text += '这几处有风险：\n  · ' + '\n  · '.join(warn) + '\n\n'
        text += '仍然生成吗？'
        ret = QMessageBox.question(
            self, '打印预检', text, QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No if bad else QMessageBox.Yes)
        return ret == QMessageBox.Yes

    def _load_fields(self, *_a):
        path = self.tpl_combo.currentData()
        self._template_path = path
        # 换模板就换一套位置：位置是随模板存的
        self._offsets = overprint.load_offsets(path) if path else {}
        self._offsets_y = overprint.load_offsets_y(path) if path else {}
        self._shift = overprint.load_shift(path) if path else (0.0, 0.0)
        self._pos_dirty = False
        self._update_pos_hint()
        if getattr(self, 'bg_check', None) is not None:
            self.bg_check.setChecked(False)
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
            if name in overprint.TITLE_FIELDS:
                # 标题给多行框：按回车即在该处强制分行，优先于自动梯形回行
                ed = QPlainTextEdit()
                ed.setMinimumHeight(56)
                ed.setMaximumHeight(80)
                ed.setPlaceholderText("按回车可在指定位置手动分行；不分行则按右侧设置自动回行")
                if name not in _NO_MEMORY:
                    ed.setPlainText(s.value('overprint/{}'.format(name), '') or '')
            elif name in _LONG_FIELDS:
                ed = QPlainTextEdit()
                ed.setMinimumHeight(110)
                ed.setPlaceholderText("内容过长会自动缩小字号，仍放不下会提示")
            else:
                # 单行字段做成可编辑下拉：承办部门、经办人这些每次都填同样
                # 几个值，翻一下比重敲一遍快
                ed = QComboBox()
                ed.setEditable(True)
                ed.setInsertPolicy(QComboBox.NoInsert)
                hist = self._history(name)
                ed.addItems(hist)
                ed.setCurrentText('')
                if name not in _NO_MEMORY:
                    ed.setCurrentText(s.value('overprint/{}'.format(name), '') or '')
                ed.setToolTip('可直接输入；点右边箭头能选之前填过的值')
            if isinstance(ed, QPlainTextEdit):
                ed.textChanged.connect(self._schedule_preview)
            elif isinstance(ed, QComboBox):
                ed.currentTextChanged.connect(self._schedule_preview)
            else:
                ed.textChanged.connect(self._schedule_preview)
            self._editors[name] = ed
            form.addRow(name + '：', ed)
        self.scroll.setWidget(host)
        self._refresh_preview()
        self.status.setText("共 {} 个可填字段；留空的字段打印出来就是空白。".format(len(fields)))

    # ---------- 画布交互 ----------
    def _apply_zoom(self, *_a):
        z = self.zoom_combo.currentData() or 0
        self.pv_area.setWidgetResizable(not z)
        self.canvas.set_zoom(z)

    def _toggle_bg(self, *_a):
        if not self.bg_check.isChecked():
            self.canvas.set_background(None)
            return
        path = overprint.load_letterhead(self._template_path or '')
        if not path or not os.path.exists(path):
            self.bg_check.setChecked(False)
            QMessageBox.information(
                self, "还没绑定套头纸",
                "这个模板还没记住套头纸的 PDF。\n"
                "用「套头对位校验…」选一次，之后就会记住。")
            return
        pm = self._bg_cache.get(path)
        if pm is None:
            from scripts import overlay
            ok, why = overlay.can_render()
            if not ok:
                self.bg_check.setChecked(False)
                QMessageBox.information(self, "无法渲染套头", why)
                return
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                from PyQt5.QtGui import QPixmap
                png = overlay.render_page_png(path, 0)
                pm = QPixmap(png)
                try:
                    os.remove(png)
                except OSError:
                    pass
            except Exception as exc:
                self.bg_check.setChecked(False)
                QMessageBox.warning(self, "读取套头失败", str(exc))
                return
            finally:
                QApplication.restoreOverrideCursor()
            self._bg_cache[path] = pm
        self.canvas.set_background(pm)

    def _on_field_moved(self, name, x_cm):
        """拖完一个字段：位置进内存、立刻重排一次，好看到真实结果"""
        if x_cm <= 0.02:
            self._offsets.pop(name, None)
        else:
            self._offsets[name] = round(float(x_cm), 2)
        self._pos_dirty = True
        self._update_pos_hint()
        self._refresh_preview()

    def _on_sheet_moved(self, dx, dy):
        self._shift = (round(float(dx), 2), round(float(dy), 2))
        self._pos_dirty = True
        self._update_pos_hint()

    def _update_pos_hint(self):
        btn = getattr(self, 'btn_save_pos', None)
        if btn is None:
            return
        btn.setEnabled(self._pos_dirty)
        if not self._pos_dirty:
            self.drag_hint.setText("拖黑字改它的横向位置，拖空白处整张纸一起挪")
            return
        bits = []
        if self._offsets:
            bits.append("{} 个字段".format(len(self._offsets)))
        if any(self._shift):
            bits.append("整体 ({:+.2f}, {:+.2f})cm".format(*self._shift))
        self.drag_hint.setText(
            "位置已改（{}）——点「保存位置」才会记住".format("、".join(bits) or "已还原"))

    def _save_positions(self):
        if not self._template_path:
            return
        overprint.save_offsets(self._template_path, self._offsets,
                               shift=self._shift, offsets_y=self._offsets_y)
        self._pos_dirty = False
        self._update_pos_hint()
        self.status.setText("位置已保存到 {}".format(
            overprint.offsets_path(self._template_path)))

    def _reset_positions(self):
        self._offsets = {}
        self._offsets_y = {}
        self._shift = (0.0, 0.0)
        self._pos_dirty = True
        self._update_pos_hint()
        self._refresh_preview()

    # ---------- 预览 ----------
    def _schedule_preview(self, *_a):
        self._pv_timer.start()

    def _refresh_preview(self):
        if not self._template_path or not hasattr(self, 'canvas'):
            return
        try:
            plan = overprint.plan_fill(self._template_path, self._values(),
                                       title_shape=self._title_shape(),
                                       title_lines=self._title_lines(),
                                       offsets=self._offsets,
                                       offsets_y=self._offsets_y)
        except Exception as e:
            self.status.setText('预览失败：{}'.format(e))
            return
        self._last_plan = plan
        self.canvas.set_plan(plan, self._shift)
        msgs = []
        for row in plan['rows']:
            for c in row['cells']:
                if c.get('overflow'):
                    msgs.append('有内容缩到最小仍放不下，建议精简文字')
                    break
        self._refresh_shape_hint(plan)
        shrunk = sum(1 for row in plan['rows'] for c in row['cells'] if c.get('shrunk'))
        if shrunk and not msgs:
            msgs.append('{} 处已自动缩小字号以放进预留格'.format(shrunk))
        self.pv_note.setText('；'.join(dict.fromkeys(msgs)) or
                             '各栏内容均能正常放下')

    def _title_line_texts(self, plan):
        """从预览数据里取标题格的各行文字（预览怎么断，输出就怎么断）"""
        for blk in plan.get('blocks') or []:
            if blk['kind'] != 'table':
                continue
            for row in blk['rows']:
                for c in row['cells']:
                    if not c.get('is_title') or c.get('vmerge_cont'):
                        continue
                    # 只取标题正文：同格里还有白色栏目名「标  题」，
                    # 它是纸上预印的，不属于标题内容
                    txt = ''.join(s['text'] for s in c['segs']
                                  if not s.get('white'))
                    return [l for l in txt.split('\n') if l.strip()]
        return []

    def _edit_offsets(self):
        plan = getattr(self, '_last_plan', None)
        if not self._template_path or not plan:
            return
        fields = plan.get('adjustable') or []
        if not fields:
            QMessageBox.information(
                self, "打印位置微调",
                "这个模板的可填字段都在表格格子里，横向位置由格子本身定死"
                "（紧跟预印的栏目名），没有需要微调的余地。")
            return
        dlg = OffsetDialog(self._template_path, fields,
                           plan.get('field_pos') or {}, self)
        if dlg.exec_() == QDialog.Accepted:
            # 对话框里改的已经落盘了，内存里这一份要跟上，
            # 否则下一次预览还按拖动前的老位置排
            self._offsets = overprint.load_offsets(self._template_path)
            self._offsets_y = overprint.load_offsets_y(self._template_path)
            self._shift = overprint.load_shift(self._template_path)
            self._pos_dirty = False
            self._update_pos_hint()
            self._refresh_preview()
            QMessageBox.information(
                self, "已保存",
                "位置已保存到：\n{}\n\n先生成一份试打，对不准再回来微调。"
                .format(overprint.offsets_path(self._template_path)))

    def _also_build_alignment(self, docx_out, values):
        """生成套打件时顺手产出对位 PDF，放在 docx 旁边。

        纯属附加：失败只记一句提示，绝不影响已经生成好的 docx——
        用户要的是那份能打印的文件，对位件只是方便核对。
        """
        from scripts import overlay
        ok, why = overlay.can_merge()
        if not ok:
            self.status.setText('未生成对位件：{}'.format(why))
            return None
        out = os.path.splitext(docx_out)[0] + '_对位.pdf'
        try:
            overlay.build_alignment_pdf(
                self._template_path, values,
                overprint.load_letterhead(self._template_path), out,
                title_shape=self._title_shape(),
                title_lines=self._title_lines())
            return out
        except Exception as e:
            self.status.setText('未生成对位件：{}'.format(str(e)[:120]))
            return None

    def _new_template(self):
        from app.template_wizard import TemplateWizard
        dlg = TemplateWizard(self)
        if dlg.exec_() == QDialog.Accepted:
            path = getattr(dlg, 'result_path', None)
            if path:
                self._reload_templates(select=path)

    def _check_align(self):
        if not self._template_path:
            return
        from app.align_dialog import AlignDialog
        AlignDialog(self._template_path, self._values(),
                    title_shape=self._title_shape(),
                    title_lines=self._title_lines(), parent=self).exec_()

    def _title_max_lines(self, plan):
        """标题栏按预留高度最多能放几行"""
        for blk in plan.get('blocks') or []:
            if blk['kind'] != 'table':
                continue
            for row in blk['rows']:
                for c in row['cells']:
                    if c.get('is_title') and c.get('max_lines'):
                        return int(c['max_lines'])
        return None

    def _refresh_shape_hint(self, plan):
        """手动分行时把自动回行的两个下拉置灰——否则用户会以为没生效"""
        manual = self._manual_title()
        for cb in (self.shape_combo, self.lines_combo):
            cb.setEnabled(not manual)
        cap = self._title_max_lines(plan)
        # 超过格子高度的行数不给选：标题栏是纸上印死的固定框，
        # 多一行会把整行撑高、下面全部下移，整张单子就与预印栏位错开
        if cap:
            for i in range(self.lines_combo.count()):
                n = self.lines_combo.itemData(i)
                self.lines_combo.model().item(i).setEnabled(n is None or n <= cap)
        lines = self._title_line_texts(plan)
        capnote = ('；此标题栏最多放 {} 行（再多会撑高栏位、整单错位）'
                   .format(cap) if cap else '')
        if manual:
            over = ('　⚠ 手动分了 {} 行，超过栏位能放的 {} 行，会把栏位撑高'
                    .format(len(lines), cap) if cap and len(lines) > cap else '')
            self.shape_hint.setText(
                "标题已按你在标题框里的回车分成 {} 行；清掉回车即可恢复自动回行{}{}"
                .format(len(lines) or 1, capnote, over))
        elif len(lines) > 1:
            self.shape_hint.setText(
                "标题自动分成 {} 行（{}）；想改断点就在标题框里按回车{}".format(
                    len(lines), '、'.join('{} 字'.format(len(l.strip()))
                                          for l in lines), capnote))
        else:
            self.shape_hint.setText(
                "标题一行放得下；想强制分行可指定行数或在标题框里按回车{}".format(capnote))

    def _pv_scale(self, plan):
        """让整幅版面填满预览区：几何和字号同倍放大，比例与行数不变。

        真实版心只有 16.5cm（约 428px），在宽预览区里小得看不清字；
        整体等比放大既看得清，又不影响"哪一行放不下"的判断——
        折行是按真实厘米算好的，放大只是显示倍率。
        """
        avail = max(200, self.preview.viewport().width() - 24)
        want = plan['page']['width_cm'] * _PX_PER_CM     # 按整张纸的宽度铺满
        return max(0.6, min(2.2, avail / want)) if want else 1.0

    def _title_shape(self):
        cb = getattr(self, 'shape_combo', None)
        if cb is None:
            return 'trapezoid_down'
        return cb.itemData(cb.currentIndex()) or 'none'

    def _title_lines(self):
        cb = getattr(self, 'lines_combo', None)
        return cb.itemData(cb.currentIndex()) if cb is not None else None

    def _manual_title(self):
        """标题里有没有用户自己敲的回车"""
        for name in overprint.TITLE_FIELDS:
            ed = self._editors.get(name)
            if ed is not None and isinstance(ed, QPlainTextEdit):
                if '\n' in ed.toPlainText().strip():
                    return True
        return False

    def _values(self):
        out = {}
        for name, ed in self._editors.items():
            if isinstance(ed, QPlainTextEdit):
                v = ed.toPlainText()
            elif isinstance(ed, QComboBox):
                v = ed.currentText()
            else:
                v = ed.text()
            out[name] = v.strip()
        return out

    def _clear_fields(self):
        for ed in self._editors.values():
            if isinstance(ed, QPlainTextEdit):
                ed.setPlainText('')
            elif isinstance(ed, QComboBox):
                ed.setCurrentText('')
            else:
                ed.setText('')
        self._refresh_preview()

    def _import_content(self):
        """从已有 docx 抽取内容填进字段（日期自动拆成年/月/日）"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先选择套打模板")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择要适配的文档", "", "Word 文档 (*.docx);;所有文件 (*.*)")
        if not path:
            return
        self._load_content_from(path)

    def _load_content_from(self, path):
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先选择套打模板")
            return
        self._source_dir = os.path.dirname(os.path.abspath(path))
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
            elif isinstance(ed, QComboBox):
                ed.setCurrentText(val)
            else:
                ed.setText(val)
        self._refresh_preview()
        missing = [n for n in self._editors if not values.get(n)]
        msg = "已识别 {} 个字段".format(len(values))
        if missing:
            msg += "；未识别：{}（请手工补填）".format('、'.join(missing))
        self.status.setText(msg)

    # ---------- 批量 / 套头库 ----------
    def _batch(self):
        if not self._template_path:
            return
        from app.batch_dialog import BatchDialog
        BatchDialog(self._template_path, self).exec_()

    def _letterhead_lib(self):
        """套头库：入库、绑定、自动认。

        自动认用的就是「按扫描件自动对位」那套——谁的红线跟模板的框线配得
        又多又准，谁就是配这个模板的纸。
        """
        if not self._template_path:
            return
        items = overprint.list_letterheads()
        cur = overprint.load_letterhead(self._template_path)
        lines = ['套头库：{}'.format(overprint.letterhead_dir()), '']
        if items:
            for name, path in items:
                lines.append('  {}{}'.format(name,
                                             '（当前模板已绑定）' if path == cur else ''))
        else:
            lines.append('  （空的，先「入库」放几张进来）')
        box = QMessageBox(self)
        box.setWindowTitle('套头库')
        box.setText('\n'.join(lines))
        b_add = box.addButton('入库…', QMessageBox.ActionRole)
        b_match = box.addButton('自动认出配套的', QMessageBox.ActionRole)
        b_bind = box.addButton('选一张绑定…', QMessageBox.ActionRole)
        box.addButton('关闭', QMessageBox.RejectRole)
        box.exec_()
        clicked = box.clickedButton()
        if clicked is b_add:
            path, _ = QFileDialog.getOpenFileName(
                self, '选择套头纸 PDF', '', 'PDF 文件 (*.pdf)')
            if path:
                try:
                    overprint.import_letterhead(path)
                except Exception as exc:
                    QMessageBox.warning(self, '入库失败', str(exc))
                    return
                self.status.setText('已入库：{}'.format(os.path.basename(path)))
        elif clicked is b_match:
            self._match_letterhead()
        elif clicked is b_bind:
            self._bind_letterhead()

    def _match_letterhead(self):
        if not overprint.list_letterheads():
            QMessageBox.information(self, '库是空的', '先用「入库」放几张套头纸进来。')
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            hits = overprint.match_letterhead(self._template_path)
        finally:
            QApplication.restoreOverrideCursor()
        if not hits:
            QMessageBox.information(
                self, '认不出来',
                '库里没有能和这个模板对上的套头纸。\n'
                '（也可能是本机缺 PyMuPDF，量不了线——那就手动绑定。）')
            return
        best, off, pairs = hits[0]
        ret = QMessageBox.question(
            self, '认出来了',
            '最像的是「{}」：配上 {} 条线，偏差 {:.2f}cm。\n\n绑定给当前模板吗？'
            .format(os.path.basename(best), pairs, off),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if ret == QMessageBox.Yes:
            overprint.save_letterhead(self._template_path, best)
            self._bg_cache.pop(best, None)
            self.status.setText('已绑定套头：{}'.format(os.path.basename(best)))

    def _bind_letterhead(self):
        items = overprint.list_letterheads()
        if not items:
            QMessageBox.information(self, '库是空的', '先用「入库」放几张套头纸进来。')
            return
        from PyQt5.QtWidgets import QInputDialog
        names = [n for n, _p in items]
        name, ok = QInputDialog.getItem(self, '选一张绑定', '套头纸：', names, 0, False)
        if not ok:
            return
        path = dict((n, p) for n, p in items)[name]
        overprint.save_letterhead(self._template_path, path)
        self._bg_cache.pop(path, None)
        self.status.setText('已绑定套头：{}'.format(name))

    # ---------- 模板目录 / 修改模板 ----------
    def _open_template_dir(self):
        from PyQt5.QtCore import QUrl
        from PyQt5.QtGui import QDesktopServices
        d = overprint.user_overprint_dir()
        try:
            os.makedirs(d, exist_ok=True)
        except Exception as e:
            QMessageBox.warning(self, "无法打开", str(e))
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(d))
        self.status.setText("套打模板目录：{}".format(d))

    def _edit_template(self):
        """在 Word/WPS 里打开模板修改。

        自带模板随软件安装（打包后在只读的临时目录里），不能直接改，
        先复制一份到用户模板目录，之后修改的就是这份副本。
        """
        path = self._template_path
        if not path or not os.path.exists(path):
            QMessageBox.information(self, "提示", "请先选择套打模板")
            return
        bundled = os.path.normpath(overprint.bundled_overprint_dir())
        is_bundled = os.path.normpath(os.path.dirname(path)) == bundled
        if is_bundled:
            ret = QMessageBox.question(
                self, "自带模板不可直接修改",
                "「{}」是软件自带模板，随软件安装、更新时会被覆盖。\n\n"
                "现在复制一份到你的模板目录，之后修改这份副本？\n"
                "（副本会出现在模板下拉里，自带的那份保持不变）".format(
                    os.path.splitext(os.path.basename(path))[0]),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if ret != QMessageBox.Yes:
                return
            import shutil
            d = overprint.user_overprint_dir()
            os.makedirs(d, exist_ok=True)
            stem, ext = os.path.splitext(os.path.basename(path))
            dest = os.path.join(d, '{}（我的）{}'.format(stem, ext))
            n = 2
            while os.path.exists(dest):
                dest = os.path.join(d, '{}（我的{}）{}'.format(stem, n, ext))
                n += 1
            try:
                shutil.copyfile(path, dest)
            except Exception as e:
                QMessageBox.warning(self, "复制失败", str(e))
                return
            self._reload_templates(select=dest)
            path = dest

        from PyQt5.QtCore import QUrl
        from PyQt5.QtGui import QDesktopServices
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        QMessageBox.information(
            self, "已打开模板",
            "模板已用系统默认程序打开：\n{}\n\n"
            "修改要点：\n"
            "· 纸上已预印的内容（单位名、栏目名、表格线）保持**白色**——"
            "它们只占位、不打印，改成你的真实单位名也不会印出来；\n"
            "· 要填的位置写成 {{{{字段名}}}}，软件会据此生成填写栏；\n"
            "· 不要改动页边距、行高、列宽，否则套打会错位。\n\n"
            "改完保存，回到本窗口重新选一次模板即可生效。".format(path))

    def _reload_templates(self, select=None):
        self._templates = overprint.list_templates()
        self.tpl_combo.blockSignals(True)
        self.tpl_combo.clear()
        for name, p, builtin in self._templates:
            self.tpl_combo.addItem('{}{}'.format(name, '（自带）' if builtin else ''), p)
        idx = self.tpl_combo.findData(select) if select else -1
        self.tpl_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.tpl_combo.blockSignals(False)
        self._load_fields()

    def _import_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择套打模板", "", "Word 文档 (*.docx)")
        if not path:
            return
        self._add_template_file(path)

    def _add_template_file(self, path):
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
        self._reload_templates(select=dest)

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
        # 默认存到导入内容的那个文件夹；没导入过则用上次保存的位置
        base_dir = self._source_dir or settings().value('overprint/last_dir', '') or ''
        if base_dir:
            default = os.path.join(base_dir, default)
        if not self._preflight_ok(values):
            return
        out, _ = QFileDialog.getSaveFileName(self, "保存套打文件", default,
                                             "Word 文档 (*.docx)")
        if not out:
            return
        from PyQt5.QtWidgets import QApplication
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            # 用内存里的位置（可能是刚拖出来还没保存的），生成的就是预览的样子
            n, notes = overprint.fill_form(self._template_path, values, out,
                                           title_shape=self._title_shape(),
                                           title_lines=self._title_lines(),
                                           offsets=self._offsets,
                                           offsets_y=self._offsets_y,
                                           shift=self._shift)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "生成失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()

        s = settings()
        s.setValue('overprint/last_dir', os.path.dirname(os.path.abspath(out)))
        for name, val in values.items():
            if name not in _NO_MEMORY and val:
                s.setValue('overprint/{}'.format(name), val)
            self._remember(name, val)

        # 绑定过套头纸就顺手多出一份对位件，省得再开一次窗口去点
        align_out = None
        if overprint.load_letterhead(self._template_path):
            align_out = self._also_build_alignment(out, values)

        self.result_path = out
        msg = "已生成：\n{}\n\n共填入 {} 个字段。".format(out, n)
        if align_out:
            msg += ("\n\n同时生成了对位件（套头纸 + 本次内容）：\n{}\n"
                    "它只用于核对位置，真正打印仍用上面的 docx 打到预印纸上。"
                    .format(align_out))
        if notes:
            msg += "\n\n注意：\n" + "\n".join('· ' + x for x in notes)
        msg += "\n\n打印时请用预印红头纸，并确认打印机「按实际大小/100%」不缩放。"
        QMessageBox.information(self, "生成完成", msg)
        self.accept()


# ---------------- 版面预览 ----------------

_PX_PER_CM = 26          # 预览缩放：1cm ≈ 26px，A4 宽约 546px
_PT_PER_CM = 28.3465


def _esc(t):
    return (t or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _segs_html(segs, scale):
    """白色（预印在纸上的）用浅灰示意，黑色（真正打印的）用黑色实体。"""
    out = []
    for s in segs:
        txt = s.get('text', '')
        if txt == '\n':
            out.append('<br>')
            continue
        pad = s.get('pad_cm')
        if pad is not None:
            # 制表位空白：按算好的厘米数留白，不当普通空格排
            out.append('<span style="display:inline-block;width:{:.0f}px">'
                       '</span>'.format(max(0.0, pad) * _PX_PER_CM * scale))
            continue
        if not txt:
            continue
        px = max(5.0, (s.get('pt') or 14) / _PT_PER_CM * _PX_PER_CM * scale)
        if s.get('white'):
            style = 'color:#C9C4B8;font-size:{:.1f}px'.format(px)
        else:
            style = 'color:#111;font-weight:600;font-size:{:.1f}px'.format(px)
        out.append('<span style="{}">{}</span>'.format(style, _esc(txt)))
    return ''.join(out)


def render_overprint_html(plan, scale=1.0):
    """把 plan_fill 的结果画成版面示意图。

    灰字 = 纸上已预印的内容（不会打印）；黑字 = 本次真正打印的内容。
    按文档真实块顺序绘制（成文日期在表格之后），格子按真实厘米比例，
    字号按自适应后的实际磅值缩放——"字变得特别小"在预览里肉眼可见。
    表格线按模板的 tblBorders 画：左右外框为 none 就不画，
    否则会凭空多出竖线、与真实版面对不上。

    每一行单独画成一张一行的表：套打单里各行的分栏并不一致——标题行是
    2.67+13.83，承办部门行是 6.77+9.73，领导批示行是整行合并的一格。
    合在一张表里时 Qt 会把这些互相冲突的宽度合成同一套列约束（整行合并
    的那格没有 colspan，被当成第一列的宽度），标题栏就被撑到半幅宽，
    与真实版面完全对不上。逐行成表后各行只受自己的宽度约束。
    表宽用 <table width> **属性**而不是 CSS width：实测 Qt 富文本完全
    无视表格的 CSS 像素宽度（模板表宽应 428px，加了 style 仍被拉满可视区
    渲染成 1190px，格子宽出近三倍，29 字的标题于是挤成一行、
    而 Word 里明明是两行）；width 属性才被真正采纳。列宽仍用百分比，
    它是相对表宽的，表宽对了列宽就对了。

    画的是**整张 A4 纸**（21×29.7cm），不是只画版心：每一行都带上左右
    页边距两栏、上下各加一条页边距空行，纸张四边描出边线。只画版心的话
    预览的宽高比和真实 A4 对不上，用尺子比划位置就无从谈起。
    """
    page = plan['page']
    cw = plan['content_w_cm']
    # 以整页宽为基准，版心与页边距按真实厘米占比分配
    PW = page['width_cm'] * _PX_PER_CM * scale
    W = PW
    pct_l = page['left_cm'] / page['width_cm'] * 100.0
    pct_r = page['right_cm'] / page['width_cm'] * 100.0
    pct_body = 100.0 - pct_l - pct_r
    PAPER = '1px solid #C9C4B8'
    MARGIN_BG = 'background:#F2EFE9;'
    parts = []

    LINE = '1px solid #D9534F'
    NONE = '0'

    def _edge(td_extra, height_px, top=False, bottom=False):
        """一整行页边距空白（同时把纸张上下边线画出来）"""
        return ('<table width="{:.0f}" cellspacing="0" cellpadding="0" '
                'style="border-collapse:collapse;margin:0"><tr>'
                '<td height="{:.0f}" style="height:{:.0f}px;{}'
                'border-left:{};border-right:{};border-top:{};border-bottom:{};'
                '"></td></tr></table>'
                .format(PW, height_px, height_px, td_extra, PAPER, PAPER,
                        PAPER if top else '0', PAPER if bottom else '0'))

    parts.append(_edge(MARGIN_BG, page['top_cm'] * _PX_PER_CM * scale, top=True))

    def _margin_cells(extra_top=NONE, extra_bottom=NONE):
        """左右页边距两栏——纸张左右边线也由它们画"""
        style = ('%s;border-left:%s;border-top:%s;border-bottom:%s'
                 % (MARGIN_BG, PAPER, extra_top, extra_bottom))
        style_r = ('%s;border-right:%s;border-top:%s;border-bottom:%s'
                   % (MARGIN_BG, PAPER, extra_top, extra_bottom))
        return ('<td width="{:.2f}%" style="{}"></td>'.format(pct_l, style),
                '<td width="{:.2f}%" style="{}"></td>'.format(pct_r, style_r))

    for blk in plan.get('blocks') or []:
        if blk['kind'] == 'para':
            # 独立段落也套一张定宽表：div 的 CSS 宽度同样被 Qt 无视，
            # 不约束的话段落会按可视区宽度排，与表格部分对不齐
            ml, mr = _margin_cells()
            parts.append(
                '<table width="{:.0f}" cellspacing="0" cellpadding="0" '
                'style="border-collapse:collapse;margin:0"><tr>{}'
                '<td width="{:.2f}%" style="background:#FFF;text-align:{};'
                'padding:1px 0;white-space:pre-wrap">{}</td>{}</tr></table>'
                .format(PW, ml, pct_body, blk['align'],
                        _segs_html(blk['segs'], scale), mr))
            continue
        b = blk.get('borders') or {}
        rows = blk['rows']
        for ri, row in enumerate(rows):
            h = row['height_cm'] * _PX_PER_CM * scale
            widths = [max(0.01, c.get('width_cm') or 0.01) for c in row['cells']]
            total = sum(widths) or 1.0
            ml, mr = _margin_cells()
            parts.append(
                '<table width="{:.0f}" cellspacing="0" cellpadding="0" '
                'style="border-collapse:collapse;margin:0">'
                '<tr>{}'.format(PW, ml))
            n = len(row['cells'])
            for ci, c in enumerate(row['cells']):
                # 列宽按整页宽折算，版心之外还要留出左右页边距
                w = widths[ci] / total * pct_body
                top = LINE if (ri == 0 and b.get('top') != 'none') or \
                    (ri > 0 and b.get('insideH') != 'none') else NONE
                # 纵向合并的延续格：它和上一行本是同一个格子，
                # 中间不该有横线
                if c.get('vmerge_cont'):
                    top = NONE
                bottom = LINE if (ri == len(rows) - 1 and b.get('bottom') != 'none') \
                    else NONE
                left = LINE if (ci == 0 and b.get('left') != 'none') or \
                    (ci > 0 and b.get('insideV') != 'none') else NONE
                right = LINE if (ci == n - 1 and b.get('right') != 'none') else NONE
                bg = ''
                if c.get('overflow'):
                    bg = 'background:#FDECEA;'
                elif c.get('shrunk'):
                    bg = 'background:#FFF6D8;'
                badge = ''
                if c.get('shrunk'):
                    badge = ('<div style="color:#B8860B;font-size:9px;">'
                             '字号 {}→{}pt{}</div>'.format(
                                 c.get('orig_font_pt'), c.get('font_pt'),
                                 ' · 仍偏长' if c.get('overflow') else ''))
                parts.append(
                    '<td width="{:.2f}%" style="width:{:.2f}%;height:{:.0f}px;'
                    'background:#FFF;{}'
                    'border-top:{};border-bottom:{};border-left:{};border-right:{};'
                    'vertical-align:top;padding:2px 3px;">'
                    '<div style="white-space:pre-wrap;line-height:1.25;">{}</div>{}</td>'
                    .format(w, w, h, bg, top, bottom, left, right,
                            _segs_html(c['segs'], scale), badge))
            parts.append(mr + '</tr></table>')

    parts.append(_edge(MARGIN_BG, max(6.0, page['bottom_cm'] * _PX_PER_CM * scale),
                       bottom=True))
    head = ('<div style="color:#8A8578;font-size:11px;margin:0 0 4px 0;">'
            'A4 纸 {:.0f}×{:.0f}mm　左边距 {:.2f}cm　右边距 {:.2f}cm　'
            '版心宽 {:.2f}cm（预览按真实比例）</div>'.format(
                page['width_cm'] * 10, page['height_cm'] * 10,
                page['left_cm'], page['right_cm'], cw))
    return ('<html><body style="margin:6px;font-family:SimSun,serif;'
            'background:#F3F1EC">' + head + ''.join(parts) + '</body></html>')
