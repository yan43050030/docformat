# -*- coding: utf-8 -*-
"""主窗口：侧边栏导航 + 六页堆叠 + 底部状态栏"""
import os

from PyQt5.QtCore import QSize, Qt, QTimer, QUrl
from PyQt5.QtGui import QDesktopServices, QKeySequence
from PyQt5.QtWidgets import (QButtonGroup, QFrame, QHBoxLayout, QLabel,
                             QMainWindow, QMessageBox, QPushButton,
                             QShortcut, QStackedWidget, QVBoxLayout, QWidget)

from app.theme import THEMES, settings
from app.widgets.nav_icons import NAV_ICON_KINDS, make_icon

from app.presets import PresetManager
from app.theme import build_qss, current_theme_id
from app.pages.home_page import HomePage
from app.pages.presets_page import PresetsPage
from app.pages.theme_page import ThemePage
from app.pages.log_page import LogPage
from app.pages.template_draft_page import TemplateDraftPage
from app.pages.template_maker_page import TemplateMakerPage

VERSION = '3.1.0'
NAV_ITEMS = [('格式处理', 0), ('版式方案', 1), ('文书起草', 2), ('文书模板制作', 3),
             ('主题', 4), ('日志', 5)]


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setWindowTitle("DocFormat Pro v{} — 公文格式自动排版工具".format(VERSION))
        self.resize(1120, 780)
        self.setMinimumSize(900, 620)

        self.mgr = PresetManager()

        central = QWidget()
        central.setObjectName("Root")
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        # ---- 侧边栏 ----
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(180)
        sb = QVBoxLayout(sidebar)
        sb.setContentsMargins(14, 18, 14, 14)
        sb.setSpacing(4)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(4)
        brand1 = QLabel("DocFormat")
        brand1.setObjectName("Brand")
        brand2 = QLabel("Pro")
        brand2.setObjectName("BrandAccent")
        brand_row.addWidget(brand1)
        brand_row.addWidget(brand2)
        brand_row.addStretch(1)
        sb.addLayout(brand_row)
        sb.addSpacing(18)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self._nav_buttons = []
        for label, idx in NAV_ITEMS:
            btn = QPushButton(label)
            btn.setProperty("navBtn", "true")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setIconSize(QSize(18, 18))
            btn.setToolTip("{}（Ctrl+{}）".format(label, idx + 1))
            if idx == 0:
                btn.setChecked(True)
            self.nav_group.addButton(btn, idx)
            # 麒麟 V10 可能是 Qt 5.12，避免使用 5.15 才有的 idClicked
            btn.clicked.connect(lambda _=False, i=idx: self._switch_page(i))
            sb.addWidget(btn)
            self._nav_buttons.append(btn)

        sb.addStretch(1)

        help_btn = QPushButton("使用说明 (F1)")
        help_btn.setProperty("flat", "true")
        help_btn.setCursor(Qt.PointingHandCursor)
        help_btn.clicked.connect(self.show_help)
        sb.addWidget(help_btn)

        self.btn_diagnostic = QPushButton("系统诊断")
        self.btn_diagnostic.setCursor(Qt.PointingHandCursor)
        self.btn_diagnostic.setToolTip("收集系统信息，方便排查运行问题")
        self.btn_diagnostic.clicked.connect(self._show_diagnostic)
        sb.addWidget(self.btn_diagnostic)
        sb.addSpacing(6)

        ver = QLabel("版本 {} · GB/T 9704-2012".format(VERSION))
        ver.setObjectName("Version")
        ver.setWordWrap(True)
        sb.addWidget(ver)

        # ---- 页面 ----
        self.stack = QStackedWidget()
        self.home_page = HomePage(self.mgr)
        self.presets_page = PresetsPage(self.mgr)
        self.template_draft_page = TemplateDraftPage(self.mgr)
        self.template_maker_page = TemplateMakerPage()
        self.theme_page = ThemePage()
        self.log_page = LogPage()
        for page in [self.home_page, self.presets_page,
                     self.template_draft_page, self.template_maker_page,
                     self.theme_page, self.log_page]:
            self.stack.addWidget(page)

        body.addWidget(sidebar)
        body.addWidget(self.stack, 1)

        # ---- 状态栏 ----
        status = QFrame()
        status.setObjectName("StatusBar")
        status.setFixedHeight(30)
        st = QHBoxLayout(status)
        st.setContentsMargins(14, 0, 14, 0)
        self.status_preset = QLabel()
        self.status_preset.setProperty("muted", "true")
        st.addWidget(self.status_preset)
        st.addStretch(1)
        self.update_label = QLabel("")
        self.update_label.setProperty("muted", "true")
        self.update_label.setVisible(False)
        self.update_label.setCursor(Qt.PointingHandCursor)
        st.addWidget(self.update_label)
        right = QLabel("原文件不会被覆盖")
        right.setProperty("muted", "true")
        st.addWidget(right)

        outer.addLayout(body, 1)
        outer.addWidget(status)
        self.setCentralWidget(central)

        # ---- 信号接线 ----
        self.home_page.logMessage.connect(self.log_page.append)
        self.home_page.presetChanged.connect(self._on_preset_changed)
        self.presets_page.presetsChanged.connect(self._on_presets_mutated)
        self.theme_page.themeSelected.connect(self.apply_theme)

        self.apply_theme(current_theme_id())
        self._refresh_status()
        self.log_page.append('info', 'DocFormat Pro 已启动')
        self._check_deps()

        # ---- 快捷键 ----
        self._shortcuts = []
        for _label, _idx in NAV_ITEMS:
            sc = QShortcut(QKeySequence('Ctrl+{}'.format(_idx + 1)), self)
            sc.activated.connect(lambda i=_idx: self.nav_to(i))
            self._shortcuts.append(sc)
        sc_open = QShortcut(QKeySequence('Ctrl+O'), self)
        sc_open.activated.connect(self._shortcut_open_files)
        sc_run = QShortcut(QKeySequence('Ctrl+Return'), self)
        sc_run.activated.connect(self._shortcut_process)
        sc_help = QShortcut(QKeySequence(Qt.Key_F1), self)
        sc_help.activated.connect(self.show_help)
        self._shortcuts += [sc_open, sc_run, sc_help]

        # ---- 全窗口拖拽：任何页面拖入文件都进入格式处理 ----
        self.setAcceptDrops(True)

        # ---- 首次启动引导 + 后台版本检查 ----
        QTimer.singleShot(300, self._maybe_onboard)
        self._update_checker = None
        QTimer.singleShot(2000, self._start_update_check)

    # ---------- 快捷键 / 导航 ----------
    def nav_to(self, idx):
        btn = self.nav_group.button(idx)
        if btn is not None:
            btn.setChecked(True)
        self._switch_page(idx)

    def _shortcut_open_files(self):
        self.nav_to(0)
        self.home_page.pick_files()

    def _shortcut_process(self):
        if self.stack.currentIndex() == 0:
            self.home_page.start_process()

    def show_help(self):
        from app.help_dialog import HelpDialog
        HelpDialog(self).exec_()

    # ---------- 全窗口拖拽 ----------
    def dragEnterEvent(self, event):
        from app.widgets.drop_zone import filter_paths
        if event.mimeData().hasUrls() and filter_paths(event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        from app.widgets.drop_zone import filter_paths
        paths = filter_paths(event.mimeData().urls())
        if paths:
            self.nav_to(0)
            self.home_page.add_files(paths)
            event.acceptProposedAction()

    # ---------- 首次引导 / 更新检查 ----------
    def _maybe_onboard(self):
        if os.environ.get('QT_QPA_PLATFORM') == 'offscreen':
            return   # 自动化测试环境跳过弹窗
        s = settings()
        if s.value('onboarded_v3', False, type=bool):
            return
        s.setValue('onboarded_v3', True)
        from app.onboarding_dialog import OnboardingDialog
        dlg = OnboardingDialog(self)
        dlg.exec_()
        if dlg.chosen_page is not None:
            self.nav_to(dlg.chosen_page)

    def _start_update_check(self):
        if os.environ.get('QT_QPA_PLATFORM') == 'offscreen':
            return
        try:
            from app.update_check import UpdateChecker
            self._update_checker = UpdateChecker(VERSION, self)
            self._update_checker.newVersion.connect(self._on_new_version)
            self._update_checker.start()
        except Exception:
            pass

    def _on_new_version(self, tag, url):
        self.update_label.setText('发现新版本 {}，点击下载'.format(tag))
        self.update_label.setVisible(True)
        self._update_url = url
        self.update_label.mousePressEvent = (
            lambda _e: QDesktopServices.openUrl(QUrl(self._update_url)))
        self.log_page.append('info', '检测到新版本 {}：{}'.format(tag, url))

    def _check_deps(self):
        from app.template_common import check_dependencies
        results = check_dependencies()
        for status, mod, detail in results:
            level = 'warning' if status == 'warn' else 'info'
            self.log_page.append(level, '依赖检测: {}'.format(detail))

    def _show_diagnostic(self):
        from app.diagnostic import show_diagnostic_dialog
        show_diagnostic_dialog(self)

    def _switch_page(self, idx):
        self.stack.setCurrentIndex(idx)
        if idx == 1:
            self.presets_page.reload()
        elif idx == 0:
            self.home_page.reload_presets()
        elif idx == 2:
            self.template_draft_page._load_template_list()

    def apply_theme(self, tid):
        self.setStyleSheet(build_qss(tid))
        c = THEMES.get(tid) or list(THEMES.values())[0]
        for btn, kind in zip(getattr(self, '_nav_buttons', []), NAV_ICON_KINDS):
            btn.setIcon(make_icon(kind, c['ink_light'], c['accent_fg']))
        drop_zone = getattr(getattr(self, 'home_page', None), 'drop_zone', None)
        if drop_zone is not None and hasattr(drop_zone, 'set_icon_color'):
            drop_zone.set_icon_color(c['ink_muted'])

    def _on_preset_changed(self, _key):
        self._refresh_status()

    def _on_presets_mutated(self):
        self.home_page.reload_presets()
        self._refresh_status()

    def _refresh_status(self):
        p = self.mgr.get(self.mgr.active_key)
        self.status_preset.setText("当前预设: {}".format(p.get('name', '')))

    def closeEvent(self, event):
        """处理线程运行中直接关窗会崩溃（QThread destroyed），先确认并停止"""
        worker = getattr(self.home_page, 'worker', None)
        if worker is not None and worker.isRunning():
            ret = QMessageBox.question(
                self, "正在处理",
                "还有文件正在处理，确定退出吗？\n未完成的文件不会生成输出。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                event.ignore()
                return
            if hasattr(worker, 'cancel'):
                worker.cancel()
            worker.wait(5000)
        event.accept()
