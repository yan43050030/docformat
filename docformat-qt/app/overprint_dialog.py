# -*- coding: utf-8 -*-
"""套打填写：选套打模板 → 填字段 → 生成可直接打到预印纸上的文件"""
import os

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QComboBox, QDialog, QFileDialog, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                             QPlainTextEdit, QPushButton, QScrollArea,
                             QSplitter, QTextBrowser, QVBoxLayout, QWidget)

from app.theme import settings
from scripts import overprint

# 这些字段内容通常较长，用多行输入框
_LONG_FIELDS = {'拟办意见', '领导批示', '备注', '主要内容', '说明'}
# 这些字段每次都变，不做记忆
_NO_MEMORY = {'标题', '拟办意见', '领导批示', '备注', '年', '月', '日'}


class OverprintDialog(QDialog):
    def __init__(self, parent=None):
        super(OverprintDialog, self).__init__(parent)
        self.setWindowTitle("套打填写 — 打到预印红头纸上")
        self.resize(1120, 720)
        self._editors = {}
        self._template_path = None
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
        add_btn = QPushButton("添加模板…")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setToolTip("导入自己的套打模板 docx（含 {{字段名}} 占位符）")
        add_btn.clicked.connect(self._import_template)
        row.addWidget(add_btn)
        root.addLayout(row)

        src_row = QHBoxLayout()
        pick = QPushButton("从 docx 导入内容…")
        pick.setCursor(Qt.PointingHandCursor)
        pick.setToolTip("选一份已有的送审单/草稿 docx，自动识别各部分内容填进下面的字段；\n"
                        "日期会拆成年/月/日分别落位，识别不到的可手工补填")
        pick.clicked.connect(self._import_content)
        src_row.addWidget(pick)
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
        self.preview = QTextBrowser()
        pv_lay.addWidget(self.preview, 1)
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

        if not self._templates:
            self.status.setText("未找到套打模板。点「添加模板…」导入一份含 {{字段名}} 的 docx。")
        else:
            self._load_fields()

    # ---------- 字段 ----------
    def _load_fields(self, *_a):
        path = self.tpl_combo.currentData()
        self._template_path = path
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
            if name in _LONG_FIELDS:
                ed = QPlainTextEdit()
                ed.setMinimumHeight(110)
                ed.setPlaceholderText("内容过长会自动缩小字号，仍放不下会提示")
            else:
                ed = QLineEdit()
                if name not in _NO_MEMORY:
                    ed.setText(s.value('overprint/{}'.format(name), '') or '')
            if isinstance(ed, QPlainTextEdit):
                ed.textChanged.connect(self._schedule_preview)
            else:
                ed.textChanged.connect(self._schedule_preview)
            self._editors[name] = ed
            form.addRow(name + '：', ed)
        self.scroll.setWidget(host)
        self._refresh_preview()
        self.status.setText("共 {} 个可填字段；留空的字段打印出来就是空白。".format(len(fields)))

    # ---------- 预览 ----------
    def _schedule_preview(self, *_a):
        self._pv_timer.start()

    def _refresh_preview(self):
        if not self._template_path or not hasattr(self, 'preview'):
            return
        try:
            plan = overprint.plan_fill(self._template_path, self._values())
        except Exception as e:
            self.preview.setHtml('<body style="color:#888;font-family:SimSun">'
                                 '预览失败：{}</body>'.format(e))
            return
        pos = self.preview.verticalScrollBar().value()
        self.preview.setHtml(render_overprint_html(plan))
        self.preview.verticalScrollBar().setValue(pos)
        msgs = []
        for row in plan['rows']:
            for c in row['cells']:
                if c.get('overflow'):
                    msgs.append('有内容缩到最小仍放不下，建议精简文字')
                    break
        shrunk = sum(1 for row in plan['rows'] for c in row['cells'] if c.get('shrunk'))
        if shrunk and not msgs:
            msgs.append('{} 处已自动缩小字号以放进预留格'.format(shrunk))
        self.pv_note.setText('；'.join(dict.fromkeys(msgs)) or
                             '各栏内容均能正常放下')

    def _values(self):
        out = {}
        for name, ed in self._editors.items():
            out[name] = (ed.toPlainText() if isinstance(ed, QPlainTextEdit)
                         else ed.text()).strip()
        return out

    def _clear_fields(self):
        for ed in self._editors.values():
            if isinstance(ed, QPlainTextEdit):
                ed.setPlainText('')
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
            else:
                ed.setText(val)
        self._refresh_preview()
        missing = [n for n in self._editors if not values.get(n)]
        msg = "已识别 {} 个字段".format(len(values))
        if missing:
            msg += "；未识别：{}（请手工补填）".format('、'.join(missing))
        self.status.setText(msg)

    def _import_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择套打模板", "", "Word 文档 (*.docx)")
        if not path:
            return
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
        self._templates = overprint.list_templates()
        self.tpl_combo.blockSignals(True)
        self.tpl_combo.clear()
        for name, p, builtin in self._templates:
            self.tpl_combo.addItem('{}{}'.format(name, '（自带）' if builtin else ''), p)
        idx = self.tpl_combo.findData(dest)
        self.tpl_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.tpl_combo.blockSignals(False)
        self._load_fields()

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
        out, _ = QFileDialog.getSaveFileName(self, "保存套打文件", default,
                                             "Word 文档 (*.docx)")
        if not out:
            return
        from PyQt5.QtWidgets import QApplication
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            n, notes = overprint.fill_form(self._template_path, values, out)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "生成失败", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()

        s = settings()
        for name, val in values.items():
            if name not in _NO_MEMORY and val:
                s.setValue('overprint/{}'.format(name), val)

        self.result_path = out
        msg = "已生成：\n{}\n\n共填入 {} 个字段。".format(out, n)
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
    格子按真实厘米比例绘制，字号按自适应后的实际磅值缩放——
    所以"字变得特别小"在预览里就是肉眼可见的小。
    """
    page = plan['page']
    cw = plan['content_w_cm']
    parts = []
    parts.append(
        '<div style="width:{:.0f}px;background:#FFF;border:1px solid #D8D2C4;'
        'padding:{:.0f}px {:.0f}px;">'.format(
            cw * _PX_PER_CM * scale,
            page['top_cm'] * _PX_PER_CM * scale * 0.35,
            2))

    for p in plan['paras']:
        parts.append('<div style="text-align:{};margin:2px 0;white-space:pre-wrap">{}</div>'
                     .format(p['align'], _segs_html(p['segs'], scale)))

    parts.append('<table cellspacing="0" cellpadding="0" '
                 'style="width:{:.0f}px;border-collapse:collapse;margin-top:4px">'
                 .format(cw * _PX_PER_CM * scale))
    for row in plan['rows']:
        h = row['height_cm'] * _PX_PER_CM * scale
        parts.append('<tr>')
        for c in row['cells']:
            w = c['width_cm'] * _PX_PER_CM * scale
            bg = ''
            if c.get('overflow'):
                bg = 'background:#FDECEA;'          # 放不下：红底警示
            elif c.get('shrunk'):
                bg = 'background:#FFF6D8;'          # 已缩小：黄底提示
            badge = ''
            if c.get('shrunk'):
                badge = ('<div style="color:#B8860B;font-size:9px;">字号 {}→{}pt{}</div>'
                         .format(c.get('orig_font_pt'), c.get('font_pt'),
                                 ' · 仍偏长' if c.get('overflow') else ''))
            parts.append(
                '<td style="width:{:.0f}px;height:{:.0f}px;{}border:1px solid #E0A0A0;'
                'vertical-align:top;padding:2px 3px;overflow:hidden;">'
                '<div style="white-space:pre-wrap;line-height:1.25;">{}</div>{}</td>'
                .format(w, h, bg, _segs_html(c['segs'], scale), badge))
        parts.append('</tr>')
    parts.append('</table></div>')
    return ('<html><body style="margin:6px;font-family:SimSun,serif;background:#F3F1EC">'
            + ''.join(parts) + '</body></html>')
