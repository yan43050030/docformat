# -*- coding: utf-8 -*-
"""首页：文件选择/拖拽 + 处理模式 + 预设 + 开始处理"""
import os

from PyQt5.QtCore import Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QFrame,
                             QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                             QPlainTextEdit, QProgressBar, QPushButton,
                             QScrollArea, QVBoxLayout, QWidget)

from app.widgets.drop_zone import DropZone, ALLOWED_EXTS
from app.widgets.file_list import FileList
from app.worker import (MODE_AI_PASTE, MODE_DIAGNOSE, MODE_FULL,
                        MODE_PUNCTUATION, MODE_TOC_AUTO, MODE_TOC_MANUAL,
                        AiPasteWorker, ProcessWorker)

MODES = [
    (MODE_FULL, '智能一键处理', '标点修复 + 排版规范 + 样式清洗，一步到位'),
    (MODE_DIAGNOSE, '格式诊断', '仅分析文档问题，不修改文件内容'),
    (MODE_PUNCTUATION, '标点修复', '仅修复中英文标点混用，保留原有段落格式'),
    (MODE_AI_PASTE, 'AI 粘贴生成', '粘贴 AI 生成的文本或 Markdown，自动生成规范公文'),
    (MODE_TOC_AUTO, '生成自动目录（域）', '在文首插入 Word 目录域，Word/WPS 打开后右键更新域即可自动生成页码'),
    (MODE_TOC_MANUAL, '生成手动目录页', '扫描标题层级，在文首插入目录文本（页码需手动填入）'),
]


def make_card():
    card = QFrame()
    card.setProperty("card", "true")
    return card


class ModeCard(QFrame):
    """可点击的处理模式卡片（标题 + 一句话说明，选中高亮描边）"""
    clicked = pyqtSignal(str)

    def __init__(self, mid, title, desc, parent=None):
        super(ModeCard, self).__init__(parent)
        self.mid = mid
        self.setProperty("modeCard", "true")
        self.setProperty("selected", "false")
        self.setProperty("modeId", mid)
        self.setCursor(Qt.PointingHandCursor)
        v = QVBoxLayout(self)
        v.setContentsMargins(12, 10, 12, 10)
        v.setSpacing(3)
        t = QLabel(title)
        t.setProperty("modeCardTitle", "true")
        d = QLabel(desc)
        d.setProperty("muted", "true")
        d.setWordWrap(True)
        v.addWidget(t)
        v.addWidget(d)

    def set_selected(self, on):
        self.setProperty("selected", "true" if on else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.mid)


class HomePage(QWidget):
    logMessage = pyqtSignal(str, str)
    presetChanged = pyqtSignal(str)

    def __init__(self, preset_mgr, parent=None):
        super(HomePage, self).__init__(parent)
        self.mgr = preset_mgr
        self.files = []
        self.worker = None
        self._outputs = []          # 本轮成功输出的文件
        self._type_overrides = {}   # 预览中手动指定的段落类型 {路径: {序号: 类型}}
        self._seal = False          # 是否加盖公章落款布局
        self.font_check_enabled = True   # 处理前检查排版字体是否安装（测试时可关闭）
        self._build()
        self.reload_presets()
        self._restore_prefs()

    # ---------- UI ----------
    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        scroll.setWidget(host)
        outer.addWidget(scroll)

        root = QVBoxLayout(host)
        root.setContentsMargins(28, 24, 28, 16)
        root.setSpacing(14)

        title = QLabel("公文格式处理")
        title.setProperty("h1", "true")
        sub = QLabel("一键规范化 Word 文档排版，遵循 GB/T 9704-2012 标准")
        sub.setProperty("muted", "true")
        root.addWidget(title)
        root.addWidget(sub)

        # --- 文件卡片 ---
        self.file_card = make_card()
        fc = QVBoxLayout(self.file_card)
        fc.setContentsMargins(14, 14, 14, 14)
        fc.setSpacing(10)

        self.drop_zone = DropZone()
        self.drop_zone.filesDropped.connect(self.add_files)
        self.drop_zone.clicked.connect(self.pick_files)

        self.file_bar = QWidget()
        fb = QHBoxLayout(self.file_bar)
        fb.setContentsMargins(0, 0, 0, 0)
        self.file_count_label = QLabel("")
        self.file_count_label.setProperty("sectionTitle", "true")
        add_btn = QPushButton("+ 添加更多文件")
        add_btn.setProperty("flat", "true")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.clicked.connect(self.pick_files)
        clear_btn = QPushButton("清空")
        clear_btn.setProperty("flat", "true")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self.clear_files)
        fb.addWidget(self.file_count_label)
        fb.addStretch(1)
        fb.addWidget(add_btn)
        fb.addWidget(clear_btn)

        self.file_list = FileList()
        self.file_list.fileRemoved.connect(self.remove_file)
        self.file_bar.setVisible(False)
        self.file_list.setVisible(False)

        fc.addWidget(self.drop_zone)
        fc.addWidget(self.file_bar)
        fc.addWidget(self.file_list)
        root.addWidget(self.file_card)

        # --- AI 粘贴卡片（与文件卡片互斥显示）---
        self.paste_card = make_card()
        pc = QVBoxLayout(self.paste_card)
        pc.setContentsMargins(14, 14, 14, 14)
        pc.setSpacing(8)
        paste_title = QLabel("粘贴 AI 生成的文本 / Markdown")
        paste_title.setProperty("sectionTitle", "true")
        self.paste_edit = QPlainTextEdit()
        self.paste_edit.setPlaceholderText(
            "把 AI 写好的公文草稿粘贴到这里（支持 Markdown 语法，会自动清洗 #、**、- 等标记）\n"
            "点击「开始处理」后选择保存位置，即可生成排版规范的 docx 文件")
        self.paste_edit.setMinimumHeight(180)
        pc.addWidget(paste_title)
        pc.addWidget(self.paste_edit)
        self.paste_card.setVisible(False)
        root.addWidget(self.paste_card)

        # --- 模式 + 预设卡片 ---
        cfg_card = make_card()
        cols = QHBoxLayout(cfg_card)
        cols.setContentsMargins(16, 14, 16, 14)
        cols.setSpacing(24)

        # 左列：处理模式（独立竖排，不与右列共享行高）
        left_col = QVBoxLayout()
        left_col.setSpacing(2)
        mode_title = QLabel("处理模式")
        mode_title.setProperty("sectionTitle", "true")
        left_col.addWidget(mode_title)
        left_col.addSpacing(6)

        from PyQt5.QtWidgets import QGridLayout
        self._mode = MODE_FULL
        self._mode_cards = {}
        mode_grid = QGridLayout()
        mode_grid.setHorizontalSpacing(10)
        mode_grid.setVerticalSpacing(10)
        for i, (mid, label, desc) in enumerate(MODES):
            card = ModeCard(mid, label, desc)
            card.clicked.connect(self.set_mode)
            self._mode_cards[mid] = card
            mode_grid.addWidget(card, i // 2, i % 2)
        self._mode_cards[MODE_FULL].set_selected(True)
        left_col.addLayout(mode_grid)
        left_col.addStretch(1)

        # 右列：预设 + 后缀 + 徽章（独立竖排）
        right_col = QVBoxLayout()
        right_col.setSpacing(8)
        right_title = QLabel("排版预设方案")
        right_title.setProperty("sectionTitle", "true")
        right_col.addWidget(right_title)

        self.preset_combo = QComboBox()
        self.preset_combo.setMinimumWidth(180)
        self.preset_combo.currentIndexChanged.connect(self._on_preset_selected)
        right_col.addWidget(self.preset_combo)

        suffix_row = QHBoxLayout()
        suffix_row.setSpacing(6)
        suffix_row.addWidget(QLabel("输出后缀"))
        self.suffix_edit = QLineEdit("_processed")
        self.suffix_edit.setMaximumWidth(130)
        suffix_row.addWidget(self.suffix_edit)
        suffix_row.addWidget(QLabel(".docx"))
        suffix_row.addStretch(1)
        right_col.addLayout(suffix_row)

        self.revision_check = QCheckBox("以修订模式输出（改动在 Word 审阅中可见、可接受/拒绝）")
        right_col.addWidget(self.revision_check)

        self.autoopen_check = QCheckBox("处理完成后自动打开输出文件")
        right_col.addWidget(self.autoopen_check)

        # 预设信息徽章：三行紧凑显示关键排版参数
        self.badge_size = QLabel()
        self.badge_size.setProperty("badge", "true")
        self.badge_margin = QLabel()
        self.badge_margin.setProperty("badge", "true")
        self.badge_body = QLabel()
        self.badge_body.setProperty("badgeAccent", "true")
        self.badge_title = QLabel()
        self.badge_title.setProperty("badge", "true")
        self.badge_page_num = QLabel()
        self.badge_page_num.setProperty("badge", "true")
        self.badge_seal = QLabel()
        self.badge_seal.setProperty("badge", "true")
        for _row_widgets in [
            [self.badge_size, self.badge_margin],
            [self.badge_body, self.badge_title],
            [self.badge_page_num, self.badge_seal],
        ]:
            row_lay = QHBoxLayout()
            row_lay.setSpacing(6)
            for w in _row_widgets:
                row_lay.addWidget(w)
            row_lay.addStretch(1)
            right_col.addSpacing(2)
            right_col.addLayout(row_lay)
        right_col.addStretch(1)

        cols.addLayout(left_col, 5)
        cols.addLayout(right_col, 4)
        root.addWidget(cfg_card)

        # --- 操作行 ---
        action_row = QWidget()
        ar = QHBoxLayout(action_row)
        ar.setContentsMargins(0, 0, 0, 0)
        self.process_btn = QPushButton("▶ 开始处理")
        self.process_btn.setProperty("primary", "true")
        self.process_btn.setCursor(Qt.PointingHandCursor)
        self.process_btn.clicked.connect(self.start_process)
        self.preview_btn = QPushButton("预览对比")
        self.preview_btn.setCursor(Qt.PointingHandCursor)
        self.preview_btn.setEnabled(False)
        self.preview_btn.clicked.connect(self.open_preview)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self.cancel_process)
        self.open_out_btn = QPushButton("打开输出位置")
        self.open_out_btn.setVisible(False)
        self.open_out_btn.setCursor(Qt.PointingHandCursor)
        self.open_out_btn.clicked.connect(self.open_output_location)
        self.status_label = QLabel("")
        self.status_label.setProperty("muted", "true")
        ar.addWidget(self.process_btn)
        ar.addWidget(self.preview_btn)
        ar.addWidget(self.cancel_btn)
        ar.addWidget(self.open_out_btn)
        ar.addSpacing(10)
        ar.addWidget(self.status_label)
        ar.addStretch(1)
        root.addWidget(action_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        hint = QLabel("处理后的文件将在原文件旁自动生成，原文件不会被覆盖")
        hint.setProperty("muted", "true")
        hint.setAlignment(Qt.AlignCenter)
        root.addWidget(hint)
        root.addStretch(1)

    # ---------- 使用习惯记忆 ----------
    def _restore_prefs(self):
        from app.theme import settings
        s = settings()
        self.suffix_edit.setText(s.value('home/suffix', '_processed') or '_processed')
        self.revision_check.setChecked(s.value('home/revision', False, type=bool))
        self.autoopen_check.setChecked(s.value('home/auto_open', False, type=bool))
        self.autoopen_check.stateChanged.connect(self._save_prefs)
        mode = s.value('home/mode', MODE_FULL)
        if mode in self._mode_cards and mode != self._mode:
            self.set_mode(mode)
        self.suffix_edit.editingFinished.connect(self._save_prefs)
        self.revision_check.stateChanged.connect(self._save_prefs)

    def _save_prefs(self, *_a):
        from app.theme import settings
        s = settings()
        s.setValue('home/suffix', self.suffix_edit.text().strip() or '_processed')
        s.setValue('home/revision', self.revision_check.isChecked())
        s.setValue('home/auto_open', self.autoopen_check.isChecked())

    # ---------- 预设 ----------
    def reload_presets(self):
        current = self.mgr.active_key
        self.preset_combo.blockSignals(True)
        self.preset_combo.clear()
        for key, name, is_builtin in self.mgr.list_all():
            label = '{}{}'.format(name, '（内置）' if is_builtin else '')
            self.preset_combo.addItem(label, key)
        idx = self.preset_combo.findData(current)
        self.preset_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.preset_combo.blockSignals(False)
        self._update_badges()

    def _on_preset_selected(self, _idx):
        key = self.preset_combo.currentData()
        if key:
            self.mgr.set_active(key)
            self._update_badges()
            self.presetChanged.emit(key)

    def _update_badges(self):
        p = self.mgr.get(self.mgr.active_key)
        page = p.get('page', {})
        body = p.get('body', {})
        title = p.get('title', {})
        sz = p.get('page_size', 'A4')
        grid = p.get('grid', {})
        pn = p.get('page_number', False)
        seal = p.get('gb_seal', False)

        # 行1: 页面规格 + 页边距
        ginfo = '{}行{}字'.format(grid.get('lines_per_page', ''), grid.get('chars_per_line', '')) if grid else ''
        self.badge_size.setText('{} {}'.format(sz, ginfo).strip())
        self.badge_margin.setText('边距 {}/{}/{}/{} cm'.format(
            page.get('top', '-'), page.get('bottom', '-'),
            page.get('left', '-'), page.get('right', '-')))
        # 行2: 正文字体 + 标题字体（正文用强调色徽章）
        body_ls = body.get('line_spacing')
        self.badge_body.setText('正文 {}{} {}pt 行距{}'.format(
            body.get('font_cn', '?'), ' 加粗' if body.get('bold') else '',
            body.get('size', '?'),
            '{}pt'.format(body_ls) if body_ls else '默认',
        ))
        self.badge_title.setText('标题 {} {}pt{}'.format(
            title.get('font_cn', '?'), title.get('size', '?'),
            ' 加粗' if title.get('bold') else '',
        ))
        # 行3: 页码 + 落款
        pn_txt = ''
        if pn:
            pn_txt = '页码 {} {}pt {} {}'.format(
                p.get('page_number_font', '宋体'),
                p.get('page_number_size', 14),
                {'dash': '—1—', 'plain': '1', 'page_text': '第1页', 'page_total': '1/10'}.get(
                    p.get('page_number_style', 'dash'), '—1—'),
                {'outside': '外侧', 'center': '居中'}.get(p.get('page_number_position', 'center'), '')
            )
        else:
            pn_txt = '页码 关闭'
        self.badge_page_num.setText(pn_txt)
        self.badge_seal.setText('落款 {}'.format('加盖公章' if seal else '无公章'))

    # ---------- 模式 ----------
    def current_mode(self):
        return self._mode

    def set_mode(self, mid):
        if mid not in self._mode_cards:
            return
        self._mode = mid
        for m, card in self._mode_cards.items():
            card.set_selected(m == mid)
        from app.theme import settings
        settings().setValue('home/mode', mid)
        self._on_mode_changed()

    def _on_mode_changed(self, _btn=None):
        is_paste = self.current_mode() == MODE_AI_PASTE
        self.file_card.setVisible(not is_paste)
        self.paste_card.setVisible(is_paste)
        self._update_action_state()

    # ---------- 文件 ----------
    def pick_files(self):
        flt = "文档 (*.docx *.doc *.wps *.txt *.md);;所有文件 (*.*)"
        paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", flt)
        if paths:
            self.add_files(paths)

    def add_files(self, paths):
        added = 0
        for p in paths:
            if p.lower().endswith(ALLOWED_EXTS) and p not in self.files:
                self.files.append(p)
                added += 1
        if added:
            self._refresh_file_ui()
            self.logMessage.emit('info', '已添加 {} 个文件'.format(added))

    def remove_file(self, index):
        if 0 <= index < len(self.files):
            self.files.pop(index)
            self._refresh_file_ui()

    def clear_files(self):
        # 多个文件时先确认，避免误触（清空按钮就在每行移除按钮正上方）
        if len(self.files) > 1:
            ret = QMessageBox.question(
                self, "清空列表",
                "确定移除列表中的全部 {} 个文件吗？（不会删除文件本身）".format(len(self.files)),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return
        self.files = []
        self._refresh_file_ui()

    def _refresh_file_ui(self):
        has = bool(self.files)
        self.file_bar.setVisible(has)
        self.file_list.setVisible(has)
        self.file_count_label.setText('{} 个文件已就绪'.format(len(self.files)))
        self.drop_zone.setVisible(not has)
        if has:
            self.file_list.set_files(self.files)
        self._update_action_state()

    def _update_action_state(self):
        mode = self.current_mode()
        if mode == MODE_AI_PASTE:
            self.process_btn.setEnabled(True)
            self.preview_btn.setEnabled(False)
        else:
            self.process_btn.setEnabled(bool(self.files))
            # 预览展示的是排版模拟效果，只对"智能一键处理"模式有意义
            self.preview_btn.setEnabled(bool(self.files) and mode == MODE_FULL)

    def open_preview(self):
        if not self.files:
            return
        from app.preview_dialog import PreviewDialog
        preset = self.mgr.get(self.mgr.active_key)
        dlg = PreviewDialog(self.files, preset, self)
        if dlg.exec_() == PreviewDialog.Accepted:
            self._type_overrides = dlg.get_overrides()
            self._seal = dlg.seal_check.isChecked()
            try:
                self.start_process()
            finally:
                self._type_overrides = {}
                self._seal = False

    # ---------- 字体检查 ----------
    def _missing_fonts(self):
        """用 Qt 字体库检查当前预设需要的中文字体是否已安装（三平台通用）"""
        from PyQt5.QtGui import QFontDatabase
        preset = self.mgr.get(self.mgr.active_key)
        needed = set()
        for key, fmt in preset.items():
            if isinstance(fmt, dict) and fmt.get('font_cn'):
                needed.add(fmt['font_cn'])
        if preset.get('page_number_font'):
            needed.add(preset['page_number_font'])
        try:
            installed = set(QFontDatabase().families())
        except Exception:
            return []
        return sorted(f for f in needed if f not in installed)

    def _confirm_fonts(self):
        """缺字体时提示用户；返回 False 表示用户选择取消处理"""
        if not self.font_check_enabled:
            return True
        missing = self._missing_fonts()
        if not missing:
            return True
        self.logMessage.emit('warning', '本机未安装排版所需字体：' + '、'.join(missing))
        ret = QMessageBox.question(
            self, "缺少排版字体",
            "本机未安装当前预设需要的以下字体：\n\n    {}\n\n"
            "仍可继续处理（字体名会正确写入文档），但在本机用 Word/WPS 打开时"
            "会显示为替代字体；在装有这些字体的电脑上打开则显示正常。\n\n"
            "是否继续处理？".format('\n    '.join(missing)),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        return ret == QMessageBox.Yes

    # ---------- 处理 ----------
    def start_process(self):
        if self.worker is not None and self.worker.isRunning():
            return
        mode = self.current_mode()
        preset_name, custom = self.mgr.engine_args(self.mgr.active_key)
        if self._seal:
            import copy
            if custom is None:
                custom = copy.deepcopy(self.mgr.get(self.mgr.active_key))
                preset_name = 'custom'
            custom['gb_seal'] = True
        if mode in (MODE_FULL, MODE_AI_PASTE) and not self._confirm_fonts():
            return

        if mode == MODE_AI_PASTE:
            text = self.paste_edit.toPlainText().strip()
            if not text:
                QMessageBox.information(self, "提示", "请先粘贴要生成公文的文本内容")
                return
            out, _ = QFileDialog.getSaveFileName(self, "保存生成的公文", "公文.docx", "Word 文档 (*.docx)")
            if not out:
                return
            self.worker = AiPasteWorker(text, out, preset_name, custom, self)
            self.worker.logMessage.connect(self.logMessage)
            self.worker.finishedWith.connect(self._on_ai_done)
            self._set_busy(True, indeterminate=True)
            self.worker.start()
            return

        if not self.files:
            return
        suffix = self.suffix_edit.text().strip() or '_processed'
        self._outputs = []
        self.open_out_btn.setVisible(False)
        self.file_list.reset_statuses()
        self.worker = ProcessWorker(
            self.files, mode, preset_name, custom, suffix,
            revision_mode=self.revision_check.isChecked(),
            type_overrides=self._type_overrides,
            parent=self)
        self.worker.logMessage.connect(self.logMessage)
        self.worker.progressChanged.connect(self.progress.setValue)
        self.worker.diagnoseReady.connect(self._show_diagnose)
        self.worker.fileStarted.connect(self._on_file_started)
        self.worker.fileFinished.connect(self._on_file_finished)
        self.worker.fileFailed.connect(self._on_file_failed)
        self.worker.allFinished.connect(self._on_all_done)
        self._set_busy(True)
        self.worker.start()

    def _on_file_started(self, path):
        self.file_list.set_status(path, 'processing')

    def _on_file_finished(self, path, output):
        self.file_list.set_status(path, 'ok', output or '')
        if output:
            self._outputs.append(output)

    def _on_file_failed(self, path, error):
        self.file_list.set_status(path, 'fail', error)

    def open_output_location(self):
        if not self._outputs:
            return
        target = os.path.dirname(self._outputs[-1]) or '.'
        QDesktopServices.openUrl(QUrl.fromLocalFile(target))

    def _auto_open_outputs(self):
        """处理完成后按设置自动打开：单文件开文件本身，多文件开所在文件夹"""
        if not self.autoopen_check.isChecked() or not self._outputs:
            return
        if len(self._outputs) == 1:
            QDesktopServices.openUrl(QUrl.fromLocalFile(self._outputs[0]))
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(
                os.path.dirname(self._outputs[-1]) or '.'))

    def cancel_process(self):
        if self.worker is not None and isinstance(self.worker, ProcessWorker):
            self.worker.cancel()
            self.cancel_btn.setEnabled(False)

    def _set_busy(self, busy, indeterminate=False):
        self.process_btn.setEnabled(not busy)
        self.process_btn.setText("处理中..." if busy else "▶ 开始处理")
        self.cancel_btn.setVisible(busy and not indeterminate)
        self.cancel_btn.setEnabled(True)
        self.progress.setVisible(busy)
        if indeterminate:
            self.progress.setRange(0, 0)
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
        self.status_label.setText("正在处理，请稍候..." if busy else "")

    def _on_all_done(self, ok, fail):
        self._set_busy(False)
        self.status_label.setText("完成：成功 {} 个，失败 {} 个".format(ok, fail))
        self.open_out_btn.setVisible(bool(self._outputs))
        self.logMessage.emit('info', '全部处理完成：成功 {}，失败 {}'.format(ok, fail))
        self.worker = None
        self._update_action_state()
        self._auto_open_outputs()

    def _on_ai_done(self, success, payload):
        self._set_busy(False)
        self.worker = None
        if success:
            self.status_label.setText("已生成: {}".format(os.path.basename(payload)))
            self._outputs = [payload]
            self.open_out_btn.setVisible(True)
            self._auto_open_outputs()
            QMessageBox.information(self, "生成成功", "公文已生成：\n{}".format(payload))
        else:
            QMessageBox.warning(self, "生成失败", payload)
        self._update_action_state()

    def _show_diagnose(self, report):
        self.logMessage.emit('info', '诊断报告：\n' + report)
        from app.report_dialog import ReportDialog
        dlg = ReportDialog(report, self)
        if dlg.exec_() == ReportDialog.Accepted:
            # 一键转入智能修复：切换模式并立即处理同一批文件
            self.set_mode(MODE_FULL)
            self.start_process()
