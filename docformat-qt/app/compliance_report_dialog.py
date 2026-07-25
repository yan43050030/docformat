# -*- coding: utf-8 -*-
"""公文合规检查结果对话框（可交互修正）

结构：
- 偏差按「页面/内容」与「段落·各类型」分组，每组可整组勾选；
- 可自动修正的偏差 → 勾选框，勾中才改；
- 不可自动修正的（结构缺失、序号层次）→ 只提示，标注需手动处理；
- 合格项默认折叠，点一下可展开核对。
点「应用所选修改并另存」后，仅对勾选项动手，原文件不动，结果另存。
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QCheckBox, QDialog, QFrame, QHBoxLayout, QLabel,
                             QPushButton, QScrollArea, QVBoxLayout, QWidget)

from scripts.compliance import TYPE_LABELS, fix_label


def _group_of(finding):
    """把 finding 归到一个展示分组：段落项按类型分组，其余归「页面与内容」。"""
    fk = finding.get('fix_key') or ''
    if fk.startswith('para:'):
        parts = fk.split(':')
        if len(parts) == 3:
            return TYPE_LABELS.get(parts[1], parts[1])
    item = finding.get('item', '')
    if '·' in item:            # 合格的段落项形如「正文·字体」
        return item.split('·')[0]
    return '页面与内容'


class ComplianceReportDialog(QDialog):
    """exec_() 返回 Accepted 表示用户点了「应用所选修改」，且至少选了一项。"""

    def __init__(self, results, parent=None):
        super(ComplianceReportDialog, self).__init__(parent)
        self.setWindowTitle("公文合规检查结果")
        self.resize(780, 660)
        self._results = results
        self._boxes = {}          # {result_index: {fix_key: QCheckBox}}
        self._fixable_total = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(8)

        tip = QLabel("检查标准来自你当前选中的预设，段落项已逐段按识别类型核对。"
                     "勾选你认可、希望自动修正的偏差，其余保持不动——"
                     "修正结果会另存为新文件，原文件不改。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        scroll.setWidget(host)
        body = QVBoxLayout(host)
        body.setContentsMargins(2, 2, 2, 2)
        body.setSpacing(12)

        for ri, res in enumerate(results):
            body.addWidget(self._build_file_block(ri, res))
        body.addStretch(1)
        root.addWidget(scroll, 1)

        btns = QHBoxLayout()
        self.select_all = QCheckBox("全选可修正项")
        self.select_all.stateChanged.connect(self._toggle_all)
        if self._fixable_total == 0:
            self.select_all.setEnabled(False)
        btns.addWidget(self.select_all)
        btns.addStretch(1)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        self.apply_btn = QPushButton("应用所选修改并另存")
        self.apply_btn.setProperty("primary", "true")
        self.apply_btn.setCursor(Qt.PointingHandCursor)
        self.apply_btn.setEnabled(False)
        self.apply_btn.clicked.connect(self._on_apply)
        btns.addWidget(close_btn)
        btns.addWidget(self.apply_btn)
        root.addLayout(btns)

    # ---------- 构建 ----------
    def _build_file_block(self, ri, res):
        card = QFrame()
        card.setProperty("card", "true")
        v = QVBoxLayout(card)
        v.setContentsMargins(14, 12, 14, 12)
        v.setSpacing(6)

        findings = res.get('findings', [])
        warns = [f for f in findings if f['level'] == 'warn']
        oks = [f for f in findings if f['level'] != 'warn']

        head = QLabel('◆ {}'.format(res.get('display', '')))
        head.setProperty("sectionTitle", "true")
        v.addWidget(head)

        meta = QLabel('对照预设：{}　·　{}'.format(
            res.get('preset_name', '') or '当前预设',
            '存在 {} 项偏差'.format(len(warns)) if warns else '未发现偏差 ✓'))
        meta.setProperty("muted", "true")
        v.addWidget(meta)

        fixable = res.get('fix_input') is not None
        self._boxes[ri] = {}
        if not fixable and warns:
            note = QLabel("· 此文件非 .docx 格式，暂不支持自动修正，"
                          "请先在 Word/WPS 里另存为 .docx")
            note.setProperty("muted", "true")
            note.setWordWrap(True)
            v.addWidget(note)

        # 偏差按分组展示，「页面与内容」排在最前
        groups = {}
        for f in warns:
            groups.setdefault(_group_of(f), []).append(f)
        order = sorted(groups.keys(), key=lambda g: (g != '页面与内容', g))
        for gname in order:
            v.addWidget(self._build_group(ri, gname, groups[gname], fixable))

        if oks:
            v.addWidget(self._build_ok_block(oks))
        return card

    def _build_group(self, ri, gname, items, fixable):
        box = QFrame()
        gv = QVBoxLayout(box)
        gv.setContentsMargins(0, 4, 0, 0)
        gv.setSpacing(3)

        fixables = [f for f in items if f.get('fix_key') and fixable]
        header = QHBoxLayout()
        gl = QLabel('{}（{} 项偏差）'.format(gname, len(items)))
        gl.setProperty("sectionTitle", "true")
        header.addWidget(gl)
        header.addStretch(1)
        if fixables:
            pick = QPushButton("全选本组")
            pick.setProperty("flat", "true")
            pick.setCursor(Qt.PointingHandCursor)
            header.addWidget(pick)
        gv.addLayout(header)

        group_boxes = []
        for f in items:
            fk = f.get('fix_key')
            text = '【{}】{}'.format(f['item'], f['detail'])
            if fk and fixable:
                cb = QCheckBox('✗ {}　→ 可自动{}'.format(text, fix_label(fk)))
                cb.stateChanged.connect(self._refresh_apply)
                self._boxes[ri][fk] = cb
                self._fixable_total += 1
                group_boxes.append(cb)
                gv.addWidget(cb)
            else:
                row = QLabel('✗ {}{}'.format(text, '　（需手动处理）' if not fk else ''))
                row.setWordWrap(True)
                gv.addWidget(row)

        if fixables and group_boxes:
            def _pick_all(_c=False, _boxes=group_boxes):
                target = not all(b.isChecked() for b in _boxes)
                for b in _boxes:
                    b.setChecked(target)
            pick.clicked.connect(_pick_all)
        return box

    def _build_ok_block(self, oks):
        box = QFrame()
        bv = QVBoxLayout(box)
        bv.setContentsMargins(0, 6, 0, 0)
        bv.setSpacing(3)

        toggle = QPushButton('✓ {} 项已符合预设（点击展开核对）'.format(len(oks)))
        toggle.setProperty("flat", "true")
        toggle.setCursor(Qt.PointingHandCursor)
        toggle.setCheckable(True)
        bv.addWidget(toggle, 0, Qt.AlignLeft)

        detail = QWidget()
        dv = QVBoxLayout(detail)
        dv.setContentsMargins(12, 2, 0, 0)
        dv.setSpacing(2)
        for f in oks:
            mark = '✓' if f['level'] == 'ok' else '·'
            row = QLabel('{} 【{}】{}'.format(mark, f['item'], f['detail']))
            row.setProperty("muted", "true")
            row.setWordWrap(True)
            dv.addWidget(row)
        detail.setVisible(False)
        bv.addWidget(detail)
        toggle.toggled.connect(detail.setVisible)
        return box

    # ---------- 交互 ----------
    def _toggle_all(self, state):
        on = state == Qt.Checked
        for keys in self._boxes.values():
            for cb in keys.values():
                cb.blockSignals(True)
                cb.setChecked(on)
                cb.blockSignals(False)
        self._refresh_apply()

    def _refresh_apply(self, *_a):
        self.apply_btn.setEnabled(any(
            cb.isChecked() for keys in self._boxes.values() for cb in keys.values()))

    def _on_apply(self):
        if any(cb.isChecked() for keys in self._boxes.values() for cb in keys.values()):
            self.accept()

    def selections(self):
        """返回 [{'fix_input','display','preset','fix_keys':[...]}]，仅含有勾选的文件。"""
        out = []
        for ri, res in enumerate(self._results):
            keys = [k for k, cb in self._boxes.get(ri, {}).items() if cb.isChecked()]
            if keys and res.get('fix_input'):
                out.append({
                    'fix_input': res['fix_input'],
                    'display': res.get('display', ''),
                    'preset': res.get('preset', {}),
                    'fix_keys': keys,
                })
        return out
