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
                             QPushButton, QScrollArea, QSplitter,
                             QTextBrowser, QVBoxLayout, QWidget)

from scripts.compliance import (ALIGN_LABELS, TYPE_LABELS, fix_label,
                                preview_spec_after)
# 复用排版预览的中西文分段渲染：Qt 富文本不做逐字体回退，
# 中英混排时数字会落到中文字体，必须手动切分才与真实 docx 一致
from app.preview_dialog import _segment_font_html

_ALIGN_CSS = {'left': 'left', 'center': 'center', 'right': 'right',
              'justify': 'justify', 'distribute': 'justify'}


def _render_wording_html(entries, side):
    """用语预览：只列会被改的段落，把词标出来。

    左右两侧的标记位置不同——原文里标的是**错词**，改后标的是**新词**，
    两者长度可能不一样（"3个"→"三个"），所以各按各自的坐标标。
    """
    html = ['<html><body style="font-family:SimSun;font-size:11pt;'
            'line-height:1.7;margin:10px">']
    for e in entries:
        text = e['before'] if side == 'before' else e['after']
        # marks 里 (起, 止) 是改后文本的坐标；原文侧要按错词自己的长度倒推
        spans = []
        if side == 'after':
            spans = [(a, b) for a, b, _o, _n in e['marks']]
        else:
            shift = 0
            for a, b, old, new in e['marks']:
                s0 = a - shift
                spans.append((s0, s0 + len(old)))
                shift += len(new) - len(old)
        color = '#FFD9D9' if side == 'before' else '#D6F5DC'
        out, cur = [], 0
        for a, b in spans:
            a, b = max(0, a), min(len(text), b)
            if a < cur:
                continue
            out.append(_esc(text[cur:a]))
            out.append('<span style="background:{}">{}</span>'.format(
                color, _esc(text[a:b])))
            cur = b
        out.append(_esc(text[cur:]))
        html.append('<p><a name="p{}"></a><span style="color:#999">第{}段　</span>{}</p>'
                    .format(e['index'], e['index'] + 1, ''.join(out)))
    html.append('</body></html>')
    return ''.join(html)


def _render_preview_html(entries, fix_keys, side):
    """渲染合规预览。side='before' 用实际格式，'after' 按已认可项修正后的格式。

    偏差处用底色标出；'after' 侧只有被认可的项才会变，未认可的保持现状。
    """
    approved = set(fix_keys)
    parts = []
    for e in entries:
        spec = e['actual'] if side == 'before' else preview_spec_after(e, approved)
        style = []
        size = spec.get('size') or 16
        style.append('font-size:{}pt'.format(size))
        style.append('text-align:{}'.format(
            _ALIGN_CSS.get(spec.get('align') or 'justify', 'justify')))
        ind = spec.get('indent') or 0
        if ind:
            style.append('text-indent:{}pt'.format(ind))
        ls = spec.get('line_spacing')
        if ls:
            style.append('line-height:{}pt'.format(ls))
        if spec.get('bold'):
            style.append('font-weight:bold')
        sb = spec.get('space_before') or 0
        sa = spec.get('space_after') or 0
        style.append('margin:{}pt 0 {}pt 0'.format(sb, sa))
        font = spec.get('font')
        font_en = spec.get('font_en') or 'Times New Roman'
        # 正文按中西文分段套字体（数字/英文走 font_en），段落级不设 font-family
        inner = _segment_font_html(e['text'], font_en, font or '宋体')

        # 改动标记：before 侧标出所有偏差；after 侧标出本次会被改的
        if side == 'before':
            mark = bool(e['bad'])
        else:
            mark = any('para:{}:{}'.format(e['ptype'], a) in approved for a in e['bad'])
        if mark:
            style.append('background-color:#FFF6D8')
        cls = ' class="chg"' if mark else ''
        tag = TYPE_LABELS.get(e['ptype'], e['ptype'])
        desc = '{} {}pt{} {}'.format(
            font or '未设置', round(size, 1) if size else '?',
            ' 粗' if spec.get('bold') else '',
            ALIGN_LABELS.get(spec.get('align'), '未设置'))
        parts.append(
            '<p{} style="{}"><a name="p{}"></a>'
            '<span class="tag">{} · 第{}段</span>'
            '<span class="meta">{}</span><br>{}</p>'.format(
                cls, '; '.join(style), e['index'], tag, e['index'], desc, inner))
    return ('<html><head><style>'
            'body {{ font-family:"SimSun",serif; font-size:12pt; margin:10px; }}'
            'p {{ white-space:pre-wrap; padding:2px 4px; }}'
            'p.chg {{ background:#FFF6D8; border-left:3px solid #E0A800; }}'
            '.tag {{ font-size:8pt; color:#666; background:#F0EDE6; '
            'border:1px solid #D8D2C4; border-radius:3px; padding:0 4px; '
            'margin-right:6px; }}'
            '.meta {{ font-size:8pt; color:#999; }}'
            '</style></head><body>{}</body></html>').format(''.join(parts))


def _esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


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
        self._wording = bool(results) and all(
            r.get('kind') == 'wording' for r in results)
        self.setWindowTitle("公文用语检查结果" if self._wording else "公文合规检查结果")
        self.resize(780, 660)
        self._results = results
        self._boxes = {}          # {result_index: {fix_key: QCheckBox}}
        self._sec_pick = {}       # {result_index: (密级下拉, 期限下拉)}
        self._fixable_total = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(8)

        tip = QLabel(
            "上方勾选你认可、希望改掉的问题，下方即时显示改哪个词、改成什么；"
            "未勾选的一个字都不动。点问题条目可跳到对应段落。"
            "改动一律写成 <b>Word 修订</b>，在 Word/WPS 里能逐条看、逐条接受或拒绝，"
            "结果另存为新文件，原文件不改。"
            if self._wording else
            "上方勾选你认可、希望自动修正的偏差，下方即时显示改哪儿、改成什么样；"
            "未勾选的保持不动。点问题条目可跳到对应段落。"
            "修正结果另存为新文件，原文件不改。")
        tip.setProperty("muted", "true")
        tip.setWordWrap(True)
        root.addWidget(tip)

        # 问题清单与对比预览同屏：勾选时立刻看到改哪儿、改成什么样，
        # 不用在标签页之间来回切
        self.main_split = QSplitter(Qt.Vertical)

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
        self.main_split.addWidget(scroll)
        self.main_split.addWidget(self._build_preview_tab())
        self.main_split.setStretchFactor(0, 3)
        self.main_split.setStretchFactor(1, 4)
        self.main_split.setSizes([300, 380])
        root.addWidget(self.main_split, 1)

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

        self._render_preview()      # 打开即显示现状，无需先切换

    # ---------- 对比预览 ----------
    def _build_preview_tab(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(4, 8, 4, 4)
        v.setSpacing(6)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("预览文件："))
        from PyQt5.QtWidgets import QComboBox
        self.pv_combo = QComboBox()
        for ri, res in enumerate(self._results):
            self.pv_combo.addItem(res.get('display', ''), ri)
        self.pv_combo.currentIndexChanged.connect(self._render_preview)
        bar.addWidget(self.pv_combo, 1)
        self.pv_note = QLabel("")
        self.pv_note.setProperty("muted", "true")
        bar.addWidget(self.pv_note)
        v.addLayout(bar)

        head = QHBoxLayout()
        hl = QLabel("原文（红底=有问题的词）" if self._wording
                    else "现状（黄底=存在偏差）")
        hl.setProperty("sectionTitle", "true")
        hr = QLabel("改后（绿底=改成的词）" if self._wording
                    else "修正后（黄底=本次会被改）")
        hr.setProperty("sectionTitle", "true")
        head.addWidget(hl, 1)
        head.addWidget(hr, 1)
        v.addLayout(head)

        split = QSplitter(Qt.Horizontal)
        self.pv_before = QTextBrowser()
        self.pv_after = QTextBrowser()
        self._pv_lock = False
        self.pv_before.verticalScrollBar().valueChanged.connect(
            lambda _x: self._sync_pv(self.pv_before, self.pv_after))
        self.pv_after.verticalScrollBar().valueChanged.connect(
            lambda _x: self._sync_pv(self.pv_after, self.pv_before))
        split.addWidget(self.pv_before)
        split.addWidget(self.pv_after)
        split.setSizes([400, 400])
        v.addWidget(split, 1)
        return page

    def _sync_pv(self, src, dst):
        if self._pv_lock:
            return
        self._pv_lock = True
        try:
            sb, db = src.verticalScrollBar(), dst.verticalScrollBar()
            ratio = sb.value() / float(sb.maximum()) if sb.maximum() else 0
            db.setValue(int(ratio * db.maximum()))
        finally:
            self._pv_lock = False

    def locate(self, ri, locations):
        """点问题条目 → 预览跳到该问题所在段落"""
        if not locations:
            return
        if hasattr(self, 'pv_combo'):
            idx = self.pv_combo.findData(ri)
            if idx >= 0 and idx != self.pv_combo.currentIndex():
                self.pv_combo.setCurrentIndex(idx)
        anchor = 'p{}'.format(locations[0])
        for view in (self.pv_before, self.pv_after):
            view.scrollToAnchor(anchor)

    def _render_preview(self, *_a):
        if not hasattr(self, 'pv_combo'):
            return
        ri = self.pv_combo.currentData()
        if ri is None:
            return
        res = self._results[ri]
        if self._wording:
            self._render_wording_preview(ri, res)
            return
        entries = res.get('preview') or []
        if not entries:
            empty = ('<html><body style="font-family:SimSun;margin:14px;color:#888">'
                     '此文件无可预览的段落</body></html>')
            self.pv_before.setHtml(empty)
            self.pv_after.setHtml(empty)
            self.pv_note.setText('')
            return
        keys = [k for k, cb in self._boxes.get(ri, {}).items() if cb.isChecked()]
        para_keys = [k for k in keys if k.startswith('para:')]
        pos_b = self.pv_before.verticalScrollBar().value()
        self.pv_before.setHtml(_render_preview_html(entries, para_keys, 'before'))
        self.pv_after.setHtml(_render_preview_html(entries, para_keys, 'after'))
        self.pv_before.verticalScrollBar().setValue(pos_b)
        n_bad = sum(1 for e in entries if e['bad'])
        n_fix = sum(1 for e in entries
                    if any('para:{}:{}'.format(e['ptype'], a) in set(para_keys)
                           for a in e['bad']))
        doc_keys = [k for k in keys if not k.startswith('para:')]
        extra = '；另有 {} 项页面/内容级修正（不在此预览）'.format(len(doc_keys)) if doc_keys else ''
        self.pv_note.setText('{} 段有偏差，本次将修正 {} 段{}'.format(n_bad, n_fix, extra))

    def _render_wording_preview(self, ri, res):
        """用语检查的预览：左边原文标出错词，右边改后标出新词。

        只列**会被改的段落**——全文照抄一遍，真正改动的两三个词反而淹了。
        """
        keys = [k for k, cb in self._boxes.get(ri, {}).items() if cb.isChecked()]
        path = res.get('fix_input')
        entries = []
        if path and keys:
            try:
                from scripts.compliance import wording_preview
                entries = wording_preview(path, res.get('preset') or {}, keys)
            except Exception as exc:
                self.pv_note.setText('预览失败：{}'.format(exc))
        if not entries:
            why = ('勾选上面的问题，这里显示会改哪个词' if path
                   else '此文件非 .docx，无法预览改动')
            empty = ('<html><body style="font-family:SimSun;margin:14px;'
                     'color:#888">{}</body></html>'.format(why))
            self.pv_before.setHtml(empty)
            self.pv_after.setHtml(empty)
            self.pv_note.setText('')
            return
        self.pv_before.setHtml(_render_wording_html(entries, 'before'))
        self.pv_after.setHtml(_render_wording_html(entries, 'after'))
        n = sum(len(e['marks']) for e in entries)
        self.pv_note.setText('{} 段共 {} 处会改（以修订方式写入）'.format(
            len(entries), n))

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
            locs = f.get('locations') or []
            if fk and fixable:
                cb = QCheckBox('✗ {}　→ 可自动{}'.format(text, fix_label(fk)))
                cb.stateChanged.connect(self._refresh_apply)
                if locs:
                    cb.setToolTip('点击可跳到第 {} 段'.format(locs[0]))
                    cb.clicked.connect(
                        lambda _c=False, _ri=ri, _l=locs: self.locate(_ri, _l))
                self._boxes[ri][fk] = cb
                self._fixable_total += 1
                group_boxes.append(cb)
                gv.addWidget(cb)
                if fk == 'security:insert':
                    # 密级和保密期限必须由人选定——标错密级比不标更麻烦。
                    # 软件只负责把选定的内容排到版头该在的位置
                    gv.addWidget(self._security_picker(ri, cb))
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

    def _security_picker(self, ri, cb):
        """插入密级时的密级 / 保密期限选择条。

        这一项不像"把字号改成三号"那样只有一个正确答案——密级定多少级、
        保多久，是定密责任人的判断，软件不能替他填。所以给的是选择条，
        选好了才让勾。
        """
        from PyQt5.QtWidgets import QComboBox
        from scripts.security_mark import LEVELS, PERIODS
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(26, 0, 0, 4)
        h.setSpacing(6)
        tip = QLabel('定为：')
        tip.setProperty("muted", "true")
        h.addWidget(tip)
        lv = QComboBox()
        lv.addItems(LEVELS)
        lv.setCurrentIndex(LEVELS.index('秘密'))
        h.addWidget(lv)
        star = QLabel('★')
        star.setProperty("muted", "true")
        h.addWidget(star)
        pd = QComboBox()
        pd.addItems(PERIODS)
        pd.setEditable(True)
        h.addWidget(pd)
        note = QLabel('密级与保密期限由定密责任人确定，软件只负责排到版头正确位置')
        note.setProperty("muted", "true")
        note.setWordWrap(True)
        h.addWidget(note, 1)
        row.setEnabled(cb.isChecked())
        cb.stateChanged.connect(lambda _s, _r=row, _c=cb: _r.setEnabled(_c.isChecked()))
        self._sec_pick[ri] = (lv, pd)
        return row

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
        # 勾选变化即时反映到下方对比预览（同屏，无需切换）
        if hasattr(self, 'pv_before'):
            self._render_preview()

    def _on_apply(self):
        if any(cb.isChecked() for keys in self._boxes.values() for cb in keys.values()):
            self.accept()

    def selections(self):
        """返回 [{'fix_input','display','preset','fix_keys':[...]}]，仅含有勾选的文件。"""
        out = []
        for ri, res in enumerate(self._results):
            keys = [k for k, cb in self._boxes.get(ri, {}).items() if cb.isChecked()]
            # 插入密级时把用户选定的密级和期限写进 key 里带下去，
            # 中途谁也不许替他改
            if 'security:insert' in keys and ri in self._sec_pick:
                lv, pd = self._sec_pick[ri]
                period = pd.currentText().strip()
                keys = ['security:insert:{}{}'.format(
                    lv.currentText(), '★' + period if period else '')
                    if k == 'security:insert' else k for k in keys]
            if keys and res.get('fix_input'):
                out.append({
                    'fix_input': res['fix_input'],
                    'display': res.get('display', ''),
                    'preset': res.get('preset', {}),
                    'fix_keys': keys,
                })
        return out
