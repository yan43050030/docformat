# -*- coding: utf-8 -*-
"""主窗口：侧边栏导航 + 六页堆叠 + 底部状态栏"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QButtonGroup, QFrame, QHBoxLayout, QLabel,
                             QMainWindow, QMessageBox, QPushButton,
                             QStackedWidget, QVBoxLayout, QWidget)

from app.presets import PresetManager
from app.theme import build_qss, current_theme_id
from app.pages.home_page import HomePage
from app.pages.presets_page import PresetsPage
from app.pages.theme_page import ThemePage
from app.pages.log_page import LogPage
from app.pages.template_draft_page import TemplateDraftPage
from app.pages.template_maker_page import TemplateMakerPage

VERSION = '2.4.0'
NAV_ITEMS = [('处理', 0), ('预设方案', 1), ('主题', 2), ('日志', 3),
             ('模板起草', 4), ('模板制作', 5)]


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
        for label, idx in NAV_ITEMS:
            btn = QPushButton(label)
            btn.setProperty("navBtn", "true")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            if idx == 0:
                btn.setChecked(True)
            self.nav_group.addButton(btn, idx)
            # 麒麟 V10 可能是 Qt 5.12，避免使用 5.15 才有的 idClicked
            btn.clicked.connect(lambda _=False, i=idx: self._switch_page(i))
            sb.addWidget(btn)

        sb.addStretch(1)

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
        self.theme_page = ThemePage()
        self.log_page = LogPage()
        self.template_draft_page = TemplateDraftPage(self.mgr)
        self.template_maker_page = TemplateMakerPage()
        for page in [self.home_page, self.presets_page, self.theme_page, self.log_page,
                     self.template_draft_page, self.template_maker_page]:
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
        elif idx == 4:
            self.template_draft_page._load_template_list()

    def apply_theme(self, tid):
        self.setStyleSheet(build_qss(tid))

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
