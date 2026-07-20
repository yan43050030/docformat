# -*- coding: utf-8 -*-
"""预设方案页：模板 CRUD + 可视化编辑器 + 即时持久化"""
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog,
                             QFrame, QGridLayout, QHBoxLayout, QInputDialog,
                             QLabel, QMessageBox, QPushButton, QScrollArea,
                             QVBoxLayout, QWidget)

from app.widgets.collapsible import CollapsibleSection
from scripts.formatter import DEFAULT_DETECT_RULES

# ---- 锁定机制：防止滚动页面时滚轮误触修改参数值 ----
_wheel_locked = False


def set_wheel_locked(locked):
    global _wheel_locked
    _wheel_locked = locked


def is_wheel_locked():
    return _wheel_locked


class _LockCombo(QComboBox):
    """带滚轮锁定的下拉框：锁定时滚轮事件向上传递，不改变选中项"""
    def wheelEvent(self, event):
        if _wheel_locked:
            event.ignore()
        else:
            super().wheelEvent(event)


class _LockSpin(QDoubleSpinBox):
    """带滚轮锁定的数值框：锁定时滚轮事件向上传递，不改变数值"""
    def wheelEvent(self, event):
        if _wheel_locked:
            event.ignore()
        else:
            super().wheelEvent(event)

# 识别规则配置：(键, 名称, 白话说明, 常用方案[(方案名, 正则)])
# 方案下拉让不懂正则的用户直接选；选「自定义」才需要写正则。
RULE_FIELDS = [
    ('security', '密级标识',
     '怎样的一行文字算密级。只检查文档开头前 3 个非空行，且整行匹配才算。',
     [('默认：秘密★1年 / 机密★3年 / 绝密', None)]),
    ('docnum', '发文字号',
     '怎样的一行算发文字号。只检查文档开头前 6 个非空行。',
     [('默认：××发〔2026〕12号', None)]),
    ('heading1', '一级标题',
     '段落开头是什么样就算一级标题。',
     [('默认：一、二、三、', None),
      ('法律条文：第一条 第二条', r'^第[一二三四五六七八九十百]+条'),
      ('章节：第一章 第二章', r'^第[一二三四五六七八九十百]+[章节]')]),
    ('heading2', '二级标题',
     '段落开头是什么样就算二级标题。',
     [('默认：（一）（二）', None),
      ('数字点：1. 2. 3.', r'^\d+\.\s*[^\d.\s]')]),
    ('heading3', '三级标题',
     '段落开头是什么样就算三级标题（默认已排除 2026.7.18 这类日期）。',
     [('默认：1. 2. 3.', None),
      ('带括号数字：（1）（2）', r'^（\d+）|^\(\d+\)')]),
    ('heading4', '四级标题',
     '段落开头是什么样就算四级标题。',
     [('默认：（1）（2）', None)]),
    ('attachment', '附件行',
     '怎样的段落算附件说明行（会用附件专用的悬挂缩进排版）。',
     [('默认：附件： / 附件1： / 附件2', None)]),
    ('signature', '落款署名',
     '文末怎样的短行算署名。默认按结尾词判断：以公司/局/委员会等机构词结尾。',
     [('默认：以公司/局/办公室/委员会…结尾', None)]),
    ('date', '落款日期',
     '怎样的一行算日期（会规范为 2026年7月18日 样式并右对齐）。',
     [('默认：2026年7月18日 / 2026.7.18 / 2026-07-18 等', None)]),
    ('closing', '结束语',
     '怎样的段落算结束语（特此通知等，不参与标题识别）。',
     [('默认：特此通知 / 此致 / 敬礼 / 妥否…', None)]),
    ('recipient', '主送机关',
     '怎样的行算主送机关（顶格排版）。默认：以中文＋冒号结尾的短行，如「各部门、各单位：」。',
     [('默认：以冒号结尾的机关列举行', None)]),
]

ELEMENTS = [
    ('security', '密级标识（左上角定密）'), ('docnum', '发文字号'),
    ('title', '标题'), ('recipient', '主送机关'),
    ('heading1', '一级标题'), ('heading2', '二级标题'),
    ('heading3', '三级标题'), ('heading4', '四级标题'),
    ('body', '正文'), ('signature', '署名'), ('date', '日期'),
    ('attachment', '附件'), ('closing', '结尾'),
]

RULE_TYPE_LABELS = {
    'security': '密级标识', 'docnum': '发文字号', 'heading1': '一级标题',
    'heading2': '二级标题', 'heading3': '三级标题', 'heading4': '四级标题',
    'attachment': '附件行', 'signature': '落款署名', 'date': '落款日期',
    'closing': '结束语', 'recipient': '主送机关',
}

# 部分要素除规则外还有位置条件，测试结果里做补充说明
_RULE_POS_NOTES = {
    'security': "（还需位于文档开头前 3 个非空行）",
    'docnum': "（还需位于文档开头前 6 个非空行）",
    'signature': "（还需位于文档末尾附近）",
    'date': "（正文中的日期不受影响，只有文末的日期行按落款日期排版）",
}


def key_pos_note(key):
    return _RULE_POS_NOTES.get(key, '')

CN_FONTS = ['方正小标宋_GBK', '方正小标宋简体', '方正仿宋_GBK', '仿宋_GB2312',
            '仿宋', '黑体', '楷体_GB2312', '楷体', '宋体', '华文中宋',
            '方正书宋_GBK', '方正楷体_GBK', '方正黑体_GBK']
EN_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Cambria', 'Georgia']
FONT_SIZES = [('初号 42', 42), ('小初 36', 36), ('一号 26', 26), ('小一 24', 24),
              ('二号 22', 22), ('小二 18', 18), ('三号 16', 16), ('小三 15', 15),
              ('四号 14', 14), ('小四 12', 12), ('五号 10.5', 10.5), ('小五 9', 9)]
ALIGNS = [('左对齐', 'left'), ('居中', 'center'), ('右对齐', 'right'), ('两端对齐', 'justify')]
PN_STYLES = [('两侧横线 — 1 —', 'dash'), ('纯数字 1', 'plain'),
             ('第 1 页', 'page_text'), ('1 / 10', 'page_total')]
PN_POSITIONS = [('外侧（单右双左）', 'outside'), ('居中', 'center'),
                ('右侧', 'right'), ('左侧', 'left')]


def _font_combo(fonts, editable=True):
    cb = _LockCombo()
    cb.addItems(fonts)
    cb.setEditable(editable)
    return cb


def _size_combo():
    cb = _LockCombo()
    for label, val in FONT_SIZES:
        cb.addItem(label, val)
    return cb


def _spin(minv, maxv, step=1.0, decimals=1):
    sp = _LockSpin()
    sp.setRange(minv, maxv)
    sp.setSingleStep(step)
    sp.setDecimals(decimals)
    return sp


class PresetsPage(QWidget):
    presetsChanged = pyqtSignal()

    def __init__(self, preset_mgr, parent=None):
        super(PresetsPage, self).__init__(parent)
        self.mgr = preset_mgr
        self.current_key = None
        self.preset = None
        self._loading = False
        self._el_widgets = {}
        self._sections = []
        self._build()
        self.reload()

    # ================= UI =================
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 8)
        root.setSpacing(12)

        title = QLabel("预设方案管理")
        title.setProperty("h1", "true")
        sub = QLabel("可视化编辑排版参数，修改即自动保存到本地")
        sub.setProperty("muted", "true")
        root.addWidget(title)
        root.addWidget(sub)

        # ---- 工具条 ----
        bar_card = QFrame()
        bar_card.setProperty("card", "true")
        bar = QHBoxLayout(bar_card)
        bar.setContentsMargins(12, 10, 12, 10)
        bar.setSpacing(8)
        self.combo = QComboBox()
        self.combo.setMinimumWidth(200)
        self.combo.currentIndexChanged.connect(self._on_select)
        bar.addWidget(self.combo, 1)
        for text, slot in [("新建", self._new), ("复制", self._dup),
                           ("重命名", self._rename), ("删除", self._delete),
                           ("导入", self._import), ("导出", self._export)]:
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            if text == "删除":
                btn.setProperty("danger", "true")
                self.delete_btn = btn
            if text == "重命名":
                self.rename_btn = btn
            btn.clicked.connect(slot)
            bar.addWidget(btn)

        bar.addStretch()
        self.btn_lock = QPushButton("🔒 禁止滚轮修改参数")
        self.btn_lock.setCheckable(True)
        self.btn_lock.setChecked(True)
        self.btn_lock.setCursor(Qt.PointingHandCursor)
        self.btn_lock.setToolTip("锁定后鼠标滚轮不会误改参数值；点击下拉框或用键盘输入仍可正常修改")
        self.btn_lock.toggled.connect(self._on_lock_toggled)
        bar.addWidget(self.btn_lock)
        set_wheel_locked(True)

        root.addWidget(bar_card)

        self.builtin_hint = QLabel("内置预设参数只读（名称可通过「重命名」修改），点击「复制」可生成参数可编辑的自定义模板")
        self.builtin_hint.setProperty("muted", "true")
        root.addWidget(self.builtin_hint)

        # ---- 编辑器（滚动） ----
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.editor_host = QWidget()
        self.editor_lay = QVBoxLayout(self.editor_host)
        self.editor_lay.setContentsMargins(0, 0, 6, 12)
        self.editor_lay.setSpacing(10)
        scroll.setWidget(self.editor_host)
        root.addWidget(scroll, 1)

        self._build_page_section()
        self._build_element_sections()
        self._build_rules_section()
        self._build_table_section()
        self._build_advanced_section()
        self.editor_lay.addStretch(1)

    def _build_rules_section(self):
        import re as _re
        from PyQt5.QtWidgets import QLineEdit
        sec = CollapsibleSection("识别规则（哪些文字算标题/密级/署名…）")
        g = QGridLayout()
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(6)
        self._rule_edits = {}     # key -> 自定义正则输入框
        self._rule_combos = {}    # key -> 方案下拉

        intro = QLabel(
            "每类要素按下拉框选择识别方式即可，选「自定义」才需要填写规则。\n"
            "自定义规则用正则表达式描述「段落开头长什么样」，常用写法："
            "^ 表示行开头，$ 表示行结尾，\\d 表示任意数字，[一二三] 表示这几个字中的任意一个，"
            "+ 表示出现一次或多次，| 表示「或」。例如 ^第[一二三四五六七八九十百]+条 匹配「第十三条」。")
        intro.setProperty("muted", "true")
        intro.setWordWrap(True)
        # 允许鼠标选中复制（^ $ \d 等符号手打不便）
        intro.setTextInteractionFlags(Qt.TextSelectableByMouse)
        g.addWidget(intro, 0, 0, 1, 3)

        for row, (key, label, tip, options) in enumerate(RULE_FIELDS, start=1):
            lab = QLabel(label)
            lab.setToolTip(tip)

            combo = _LockCombo()
            for opt_label, opt_pattern in options:
                combo.addItem(opt_label, opt_pattern)   # None = 默认规则
            combo.addItem('自定义…', '__custom__')
            combo.setToolTip(tip)

            edit = QLineEdit()
            edit.setPlaceholderText('自定义正则，留空即默认：{}'.format(DEFAULT_DETECT_RULES[key]))
            edit.setToolTip(tip)
            edit.setVisible(False)

            def _validate(text, e=edit):
                try:
                    if text.strip():
                        _re.compile(text)
                    e.setStyleSheet("")
                    return True
                except _re.error:
                    e.setStyleSheet("border: 1px solid #C62828;")
                    return False

            def _on_combo(_idx, k=key, c=combo, e=edit):
                if self._loading:
                    return
                val = c.currentData()
                e.setVisible(val == '__custom__')
                if val == '__custom__':
                    # 预填当前生效的规则（此前选中的方案或默认规则），
                    # 在现有规则基础上改，比从空白写正则容易得多
                    if not e.text().strip():
                        e.blockSignals(True)
                        e.setText(DEFAULT_DETECT_RULES[k])
                        e.setStyleSheet("")
                        e.blockSignals(False)
                else:
                    # 选中的方案直接写入（None=默认→清空）
                    e.blockSignals(True)
                    e.setText(val or '')
                    e.setStyleSheet("")
                    e.blockSignals(False)
                    self._save_from_widgets()
                self._update_rule_tester()

            def _on_edited(text, e=edit):
                if _validate(text):
                    self._save_from_widgets()
                self._update_rule_tester()

            combo.currentIndexChanged.connect(_on_combo)
            edit.textChanged.connect(_on_edited)

            g.addWidget(lab, row, 0)
            g.addWidget(combo, row, 1)
            g.addWidget(edit, row, 2)
            self._rule_edits[key] = edit
            self._rule_combos[key] = combo

        # ---- 实时测试：粘贴一行文字，看会被识别成什么 ----
        test_row = len(RULE_FIELDS) + 1
        test_lab = QLabel("规则测试")
        self.rule_test_edit = QLineEdit()
        self.rule_test_edit.setPlaceholderText("在这里粘贴文档中的一行文字，立即查看它会被识别成什么")
        self.rule_test_edit.textChanged.connect(self._update_rule_tester)
        self.rule_test_result = QLabel("")
        self.rule_test_result.setProperty("muted", "true")
        self.rule_test_result.setWordWrap(True)
        self.rule_test_result.setTextInteractionFlags(Qt.TextSelectableByMouse)
        g.addWidget(test_lab, test_row, 0)
        g.addWidget(self.rule_test_edit, test_row, 1, 1, 2)
        g.addWidget(self.rule_test_result, test_row + 1, 1, 1, 2)

        g.setColumnStretch(1, 2)
        g.setColumnStretch(2, 3)
        sec.set_body_layout(g)
        self._sections.append(sec)
        self.editor_lay.addWidget(sec)

    def _current_rules(self):
        rules = {}
        for key, edit in self._rule_edits.items():
            val = edit.text().strip()
            if val:
                rules[key] = val
        return rules

    def _update_rule_tester(self, *_args):
        text = self.rule_test_edit.text().strip() if hasattr(self, 'rule_test_edit') else ''
        if not text:
            self.rule_test_result.setText("")
            return
        from scripts.formatter import _compile_rules
        compiled = _compile_rules(self._current_rules())
        # 按引擎优先级顺序测试（与 detect_para_type 中的顺序一致）
        order = ['security', 'docnum', 'heading1', 'heading2', 'heading3',
                 'heading4', 'recipient', 'attachment', 'closing', 'date', 'signature']
        matched = []
        for key in order:
            rule = compiled.get(key)
            if rule is None:
                continue
            hit = rule.search(text) if key == 'signature' else rule.match(text)
            if hit:
                matched.append(key)
        if matched:
            first = RULE_TYPE_LABELS.get(matched[0], matched[0])
            others = [RULE_TYPE_LABELS.get(k, k) for k in matched[1:]]
            msg = "✓ 这行文字符合「{}」的识别规则{}".format(first, key_pos_note(matched[0]))
            if others:
                msg += "；也符合：{}（实际按引擎优先级和段落位置综合判断）".format('、'.join(others))
            self.rule_test_result.setText(msg)
        else:
            self.rule_test_result.setText(
                "○ 不符合任何要素规则，将结合位置识别为正文、标题或落款等")

    def _build_page_section(self):
        sec = CollapsibleSection("页面与页码", expanded=True)
        g = QGridLayout()
        g.setHorizontalSpacing(12)
        g.setVerticalSpacing(8)

        self.margin_spins = {}
        for col, (key, label) in enumerate([('top', '上边距'), ('bottom', '下边距'),
                                            ('left', '左边距'), ('right', '右边距')]):
            g.addWidget(QLabel(label + ' (cm)'), 0, col)
            sp = _spin(0.5, 10, 0.1)
            sp.valueChanged.connect(self._save_from_widgets)
            self.margin_spins[key] = sp
            g.addWidget(sp, 1, col)

        self.pn_enabled = QCheckBox("启用页码")
        self.pn_enabled.stateChanged.connect(self._save_from_widgets)
        g.addWidget(self.pn_enabled, 2, 0, 1, 2)

        g.addWidget(QLabel("页码字体"), 3, 0)
        self.pn_font = _font_combo(CN_FONTS)
        self.pn_font.currentTextChanged.connect(self._save_from_widgets)
        g.addWidget(self.pn_font, 4, 0)

        g.addWidget(QLabel("页码字号"), 3, 1)
        self.pn_size = _size_combo()
        self.pn_size.currentIndexChanged.connect(self._save_from_widgets)
        g.addWidget(self.pn_size, 4, 1)

        g.addWidget(QLabel("页码样式"), 3, 2)
        self.pn_style = _LockCombo()
        for label, val in PN_STYLES:
            self.pn_style.addItem(label, val)
        self.pn_style.currentIndexChanged.connect(self._save_from_widgets)
        g.addWidget(self.pn_style, 4, 2)

        g.addWidget(QLabel("页码位置"), 3, 3)
        self.pn_pos = _LockCombo()
        for label, val in PN_POSITIONS:
            self.pn_pos.addItem(label, val)
        self.pn_pos.currentIndexChanged.connect(self._save_from_widgets)
        g.addWidget(self.pn_pos, 4, 3)

        sec.set_body_layout(g)
        self._sections.append(sec)
        self.editor_lay.addWidget(sec)

    def _build_element_sections(self):
        for key, label in ELEMENTS:
            sec = CollapsibleSection(label, expanded=(key in ('title', 'body')))
            g = QGridLayout()
            g.setHorizontalSpacing(12)
            g.setVerticalSpacing(6)

            w = {}
            g.addWidget(QLabel("中文字体"), 0, 0)
            w['font_cn'] = _font_combo(CN_FONTS)
            g.addWidget(w['font_cn'], 1, 0)

            g.addWidget(QLabel("英文字体"), 0, 1)
            w['font_en'] = _font_combo(EN_FONTS)
            g.addWidget(w['font_en'], 1, 1)

            g.addWidget(QLabel("字号"), 0, 2)
            w['size'] = _size_combo()
            g.addWidget(w['size'], 1, 2)

            g.addWidget(QLabel("对齐"), 0, 3)
            w['align'] = _LockCombo()
            for al, av in ALIGNS:
                w['align'].addItem(al, av)
            g.addWidget(w['align'], 1, 3)

            g.addWidget(QLabel("首行缩进 (pt)"), 2, 0)
            w['indent'] = _spin(0, 100, 1, 0)
            g.addWidget(w['indent'], 3, 0)

            g.addWidget(QLabel("段前 (pt)"), 2, 1)
            w['space_before'] = _spin(0, 50, 1, 0)
            g.addWidget(w['space_before'], 3, 1)

            g.addWidget(QLabel("段后 (pt)"), 2, 2)
            w['space_after'] = _spin(0, 50, 1, 0)
            g.addWidget(w['space_after'], 3, 2)

            if key == 'body':
                g.addWidget(QLabel("行距 (pt)"), 2, 3)
                w['line_spacing'] = _spin(0, 100, 1, 0)
                w['line_spacing'].setSpecialValueText("默认")
                g.addWidget(w['line_spacing'], 3, 3)

            w['bold'] = QCheckBox("加粗")
            g.addWidget(w['bold'], 4, 0)

            for widget in w.values():
                if isinstance(widget, QComboBox):
                    widget.currentIndexChanged.connect(self._save_from_widgets)
                    if widget.isEditable():
                        widget.currentTextChanged.connect(self._save_from_widgets)
                elif isinstance(widget, QDoubleSpinBox):
                    widget.valueChanged.connect(self._save_from_widgets)
                elif isinstance(widget, QCheckBox):
                    widget.stateChanged.connect(self._save_from_widgets)

            self._el_widgets[key] = w
            sec.set_body_layout(g)
            self._sections.append(sec)
            self.editor_lay.addWidget(sec)

    def _build_table_section(self):
        sec = CollapsibleSection("表格优化")
        g = QGridLayout()
        g.setHorizontalSpacing(12)
        self.tb_optimize = QCheckBox("启用表格优化（边框/列宽/行高标准化）")
        self.tb_auto_col = QCheckBox("按内容智能分配列宽")
        self.tb_blank_after = QCheckBox("表格后保留空行")
        g.addWidget(self.tb_optimize, 0, 0, 1, 2)
        g.addWidget(self.tb_auto_col, 1, 0, 1, 2)
        g.addWidget(self.tb_blank_after, 2, 0, 1, 2)
        g.addWidget(QLabel("边框粗细 (pt)"), 3, 0)
        self.tb_border = _spin(0.25, 3, 0.25, 2)
        g.addWidget(self.tb_border, 4, 0)
        g.addWidget(QLabel("行高 (cm)"), 3, 1)
        self.tb_row_h = _spin(0.3, 3, 0.1)
        g.addWidget(self.tb_row_h, 4, 1)
        for wdg in [self.tb_optimize, self.tb_auto_col, self.tb_blank_after]:
            wdg.stateChanged.connect(self._save_from_widgets)
        for wdg in [self.tb_border, self.tb_row_h]:
            wdg.valueChanged.connect(self._save_from_widgets)
        sec.set_body_layout(g)
        self._sections.append(sec)
        self.editor_lay.addWidget(sec)

    def _build_advanced_section(self):
        sec = CollapsibleSection("高级选项")
        g = QGridLayout()
        self.adv_first_bold = QCheckBox("正文段首句加粗（如「一是……」）")
        self.adv_bold_serial = QCheckBox("正文序号引导词加粗")
        self.adv_deep_clean = QCheckBox("深度清洗（清除颜色/下划线/斜体等杂样式）")
        g.addWidget(self.adv_first_bold, 0, 0)
        g.addWidget(self.adv_bold_serial, 1, 0)
        g.addWidget(self.adv_deep_clean, 2, 0)
        for wdg in [self.adv_first_bold, self.adv_bold_serial, self.adv_deep_clean]:
            wdg.stateChanged.connect(self._save_from_widgets)
        sec.set_body_layout(g)
        self._sections.append(sec)
        self.editor_lay.addWidget(sec)

    # ================= 数据绑定 =================
    def reload(self):
        self._loading = True
        current = self.mgr.active_key
        self.combo.blockSignals(True)
        self.combo.clear()
        for key, name, is_builtin in self.mgr.list_all():
            self.combo.addItem('{}{}'.format(name, '（内置）' if is_builtin else ''), key)
        idx = self.combo.findData(current)
        self.combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo.blockSignals(False)
        self._loading = False
        self._load_preset(self.combo.currentData())

    def _on_select(self, _idx):
        if self._loading:
            return
        key = self.combo.currentData()
        if key:
            self.mgr.set_active(key)
            self._load_preset(key)
            self.presetsChanged.emit()

    def _load_preset(self, key):
        if not key:
            return
        self.current_key = key
        self.preset = self.mgr.get(key)
        is_builtin = self.mgr.is_builtin(key)
        self._loading = True

        p = self.preset
        page = p.get('page', {})
        for k, sp in self.margin_spins.items():
            sp.setValue(float(page.get(k, 2.5)))

        self.pn_enabled.setChecked(bool(p.get('page_number', False)))
        self.pn_font.setCurrentText(p.get('page_number_font', '宋体'))
        self._set_combo_data(self.pn_size, p.get('page_number_size', 14))
        self._set_combo_data(self.pn_style, p.get('page_number_style', 'dash'))
        self._set_combo_data(self.pn_pos, p.get('page_number_position', 'outside'))

        for key2, w in self._el_widgets.items():
            el = p.get(key2, {})
            w['font_cn'].setCurrentText(el.get('font_cn', '仿宋_GB2312'))
            w['font_en'].setCurrentText(el.get('font_en', 'Times New Roman'))
            self._set_combo_data(w['size'], el.get('size', 16))
            self._set_combo_data(w['align'], el.get('align', 'left'))
            w['indent'].setValue(float(el.get('indent', 0)))
            w['space_before'].setValue(float(el.get('space_before', 0) or 0))
            w['space_after'].setValue(float(el.get('space_after', 0) or 0))
            if 'line_spacing' in w:
                ls = el.get('line_spacing')
                w['line_spacing'].setValue(float(ls) if ls else 0)
            w['bold'].setChecked(bool(el.get('bold', False)))

        tb = p.get('table', {})
        self.tb_optimize.setChecked(bool(tb.get('optimize', True)))
        self.tb_auto_col.setChecked(bool(tb.get('auto_col_width', True)))
        self.tb_blank_after.setChecked(bool(tb.get('after_table_blank_line', True)))
        self.tb_border.setValue(float(tb.get('border_size_pt', 0.5)))
        self.tb_row_h.setValue(float(tb.get('row_height_cm', 0.7)))

        self.adv_first_bold.setChecked(bool(p.get('first_line_bold', False)))
        self.adv_bold_serial.setChecked(bool(p.get('bold_serial', True)))
        self.adv_deep_clean.setChecked(bool(p.get('deep_clean', False)))

        rules = p.get('detect_rules', {}) or {}
        for key, edit in self._rule_edits.items():
            combo = self._rule_combos[key]
            pattern = rules.get(key, '')
            edit.setText(pattern)
            edit.setStyleSheet("")
            if not pattern:
                combo.setCurrentIndex(0)          # 默认规则
                edit.setVisible(False)
            else:
                idx = combo.findData(pattern)      # 命中某个内置方案
                if idx >= 0:
                    combo.setCurrentIndex(idx)
                    edit.setVisible(False)
                else:
                    combo.setCurrentIndex(combo.findData('__custom__'))
                    edit.setVisible(True)
        if hasattr(self, 'rule_test_edit'):
            self._update_rule_tester()

        self._loading = False
        for sec in self._sections:
            sec.set_editable(not is_builtin)
        self.builtin_hint.setVisible(is_builtin)
        self.delete_btn.setEnabled(not is_builtin)
        self.rename_btn.setEnabled(True)   # 内置模板也允许改名（参数仍只读）

    @staticmethod
    def _set_combo_data(combo, value):
        idx = combo.findData(value)
        if idx < 0 and combo.count():
            idx = 0
        combo.setCurrentIndex(idx)

    def _save_from_widgets(self, *_args):
        if self._loading or self.preset is None or self.mgr.is_builtin(self.current_key):
            return
        p = self.preset
        p['page'] = {k: sp.value() for k, sp in self.margin_spins.items()}
        p['page_number'] = self.pn_enabled.isChecked()
        p['page_number_font'] = self.pn_font.currentText()
        p['page_number_size'] = self.pn_size.currentData()
        p['page_number_style'] = self.pn_style.currentData()
        p['page_number_position'] = self.pn_pos.currentData()

        for key, w in self._el_widgets.items():
            el = p.setdefault(key, {})
            el['font_cn'] = w['font_cn'].currentText()
            el['font_en'] = w['font_en'].currentText()
            el['size'] = w['size'].currentData()
            el['align'] = w['align'].currentData()
            el['indent'] = int(w['indent'].value())
            el['space_before'] = int(w['space_before'].value())
            el['space_after'] = int(w['space_after'].value())
            el['bold'] = w['bold'].isChecked()
            if 'line_spacing' in w:
                v = w['line_spacing'].value()
                el['line_spacing'] = int(v) if v > 0 else None

        p['table'] = {
            'optimize': self.tb_optimize.isChecked(),
            'auto_col_width': self.tb_auto_col.isChecked(),
            'after_table_blank_line': self.tb_blank_after.isChecked(),
            'border_size_pt': self.tb_border.value(),
            'row_height_cm': self.tb_row_h.value(),
        }
        p['first_line_bold'] = self.adv_first_bold.isChecked()
        p['bold_serial'] = self.adv_bold_serial.isChecked()
        p['deep_clean'] = self.adv_deep_clean.isChecked()

        rules = {}
        for key, edit in self._rule_edits.items():
            val = edit.text().strip()
            if val:
                rules[key] = val
        if rules:
            p['detect_rules'] = rules
        else:
            p.pop('detect_rules', None)

        self.mgr.update(self.current_key, p)
        self.presetsChanged.emit()

    def _on_lock_toggled(self, checked):
        set_wheel_locked(checked)
        if checked:
            self.btn_lock.setText("🔒 禁止滚轮修改参数")
            self.btn_lock.setToolTip("锁定后鼠标滚轮不会误改参数值；点击下拉框或用键盘输入仍可正常修改")
        else:
            self.btn_lock.setText("🔓 滚轮可修改参数")
            self.btn_lock.setToolTip("滚轮可修改参数值，翻阅页面时请注意勿误触")

    # ================= 工具条动作 =================
    def _new(self):
        name, ok = QInputDialog.getText(self, "新建模板", "模板名称：",
                                        text="自定义格式 {}".format(len(self.mgr.user) + 1))
        if ok and name.strip():
            self.mgr.create(name.strip())
            self.reload()
            self.presetsChanged.emit()

    def _dup(self):
        self.mgr.duplicate(self.current_key)
        self.reload()
        self.presetsChanged.emit()

    def _rename(self):
        name, ok = QInputDialog.getText(self, "重命名", "新名称：",
                                        text=self.preset.get('name', ''))
        if ok and name.strip():
            self.mgr.rename(self.current_key, name.strip())
            self.reload()
            self.presetsChanged.emit()

    def _delete(self):
        if self.mgr.is_builtin(self.current_key):
            return
        ret = QMessageBox.question(self, "删除模板",
                                   "确定删除「{}」？此操作不可恢复。".format(self.preset.get('name', '')))
        if ret == QMessageBox.Yes:
            self.mgr.delete(self.current_key)
            self.reload()
            self.presetsChanged.emit()

    def _import(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入模板", "", "JSON (*.json)")
        if not path:
            return
        try:
            imported = self.mgr.import_from(path)
            if imported:
                QMessageBox.information(self, "导入成功", "已导入 {} 个模板".format(len(imported)))
                self.reload()
                self.presetsChanged.emit()
            else:
                QMessageBox.warning(self, "导入失败", "文件中没有有效的模板数据")
        except Exception as e:
            QMessageBox.warning(self, "导入失败", "文件格式不正确：{}".format(e))

    def _export(self):
        name = self.preset.get('name', 'preset') if self.preset else 'preset'
        path, _ = QFileDialog.getSaveFileName(self, "导出模板", name + ".json", "JSON (*.json)")
        if path:
            self.mgr.export_to(self.current_key, path)
            QMessageBox.information(self, "导出成功", "模板已导出到：\n" + path)
