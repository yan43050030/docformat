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
        self._last_plan = plan
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
            else:
                ed.setText(val)
        self._refresh_preview()
        missing = [n for n in self._editors if not values.get(n)]
        msg = "已识别 {} 个字段".format(len(values))
        if missing:
            msg += "；未识别：{}（请手工补填）".format('、'.join(missing))
        self.status.setText(msg)

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
        s.setValue('overprint/last_dir', os.path.dirname(os.path.abspath(out)))
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
    按文档真实块顺序绘制（成文日期在表格之后），格子按真实厘米比例，
    字号按自适应后的实际磅值缩放——"字变得特别小"在预览里肉眼可见。
    表格线按模板的 tblBorders 画：左右外框为 none 就不画，
    否则会凭空多出竖线、与真实版面对不上。

    每一行单独画成一张一行的表：套打单里各行的分栏并不一致——标题行是
    2.67+13.83，承办部门行是 6.77+9.73，领导批示行是整行合并的一格。
    合在一张表里时 Qt 会把这些互相冲突的宽度合成同一套列约束（整行合并
    的那格没有 colspan，被当成第一列的宽度），标题栏就被撑到半幅宽，
    与真实版面完全对不上。逐行成表后各行只受自己的宽度约束。
    宽度用百分比而非像素：Qt 富文本对像素宽度只当"建议值"，
    内容稍长就会自行重新分配。
    """
    page = plan['page']
    cw = plan['content_w_cm']
    W = cw * _PX_PER_CM * scale
    parts = ['<div style="width:{:.0f}px;background:#FFF;'
             'border:1px solid #D8D2C4;padding:6px 0;">'.format(W)]

    LINE = '1px solid #D9534F'
    NONE = '0'

    for blk in plan.get('blocks') or []:
        if blk['kind'] == 'para':
            parts.append(
                '<div style="text-align:{};margin:2px 0;white-space:pre-wrap">{}</div>'
                .format(blk['align'], _segs_html(blk['segs'], scale)))
            continue
        b = blk.get('borders') or {}
        rows = blk['rows']
        for ri, row in enumerate(rows):
            h = row['height_cm'] * _PX_PER_CM * scale
            widths = [max(0.01, c.get('width_cm') or 0.01) for c in row['cells']]
            total = sum(widths) or 1.0
            parts.append(
                '<table cellspacing="0" cellpadding="0" '
                'style="width:{:.0f}px;border-collapse:collapse;margin:0">'
                '<tr>'.format(W))
            n = len(row['cells'])
            for ci, c in enumerate(row['cells']):
                w = widths[ci] / total * 100.0
                top = LINE if (ri == 0 and b.get('top') != 'none') or \
                    (ri > 0 and b.get('insideH') != 'none') else NONE
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
                    '<td width="{:.2f}%" style="width:{:.2f}%;height:{:.0f}px;{}'
                    'border-top:{};border-bottom:{};border-left:{};border-right:{};'
                    'vertical-align:top;padding:2px 3px;">'
                    '<div style="white-space:pre-wrap;line-height:1.25;">{}</div>{}</td>'
                    .format(w, w, h, bg, top, bottom, left, right,
                            _segs_html(c['segs'], scale), badge))
            parts.append('</tr></table>')
    parts.append('</div>')
    return ('<html><body style="margin:6px;font-family:SimSun,serif;background:#F3F1EC">'
            + ''.join(parts) + '</body></html>')
