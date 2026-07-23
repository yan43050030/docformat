# -*- coding: utf-8 -*-
"""
系统诊断信息收集 + 崩溃报告对话框

可独立使用：即使 app 其他模块全部损坏，本模块也能工作。
对话框内置"复制全部"按钮，用户一键复制后发给开发者。
"""
import os
import sys
import platform
import traceback


def collect_system_info():
    """收集系统诊断信息，返回 {标题: 值} 有序字典。纯数据采集，无 UI 依赖。"""
    # 用 list of tuples 保证顺序
    info = []

    # ── 操作系统 ──
    info.append(("操作系统", "{} {}".format(platform.system(), platform.release())))
    info.append(("系统版本", _linux_distro()))
    info.append(("内核版本", platform.version()))
    info.append(("架构", platform.machine()))
    info.append(("主机名", platform.node()))

    # ── glibc（Linux 关键项） ──
    glibc_ver = _glibc_version()
    if glibc_ver:
        info.append(("glibc 版本", glibc_ver))

    # ── 图形环境 ──
    display = os.environ.get("DISPLAY", "")
    wayland = os.environ.get("WAYLAND_DISPLAY", "")
    if display:
        info.append(("图形环境", "X11 (DISPLAY={})".format(display)))
    elif wayland:
        info.append(("图形环境", "Wayland (WAYLAND_DISPLAY={})".format(wayland)))
    else:
        info.append(("图形环境", "未检测到（{} / {} 均为空）".format(
            "$DISPLAY" if not display else "✓",
            "$WAYLAND_DISPLAY" if not wayland else "✓")))

    desktop = os.environ.get("XDG_CURRENT_DESKTOP") or os.environ.get("DESKTOP_SESSION") or ""
    if desktop:
        info.append(("桌面环境", desktop))

    # ── Python ──
    info.append(("Python 版本", sys.version.split()[0]))
    info.append(("Python 路径", sys.executable))
    info.append(("Python 架构", "64-bit" if sys.maxsize > 2**32 else "32-bit"))

    # ── 运行模式 ──
    if getattr(sys, 'frozen', False):
        info.append(("运行模式", "PyInstaller 打包 (onefile)"))
        meipass = getattr(sys, '_MEIPASS', '')
        if meipass:
            info.append(("解压目录", meipass))
    else:
        info.append(("运行模式", "源码运行"))

    # ── 关键包版本 ──
    for mod_name, pkg_name in [("PyQt5", "PyQt5"), ("docx", "python-docx"),
                                 ("lxml", "lxml"), ("Qt", "PyQt5-Qt")]:
        ver = _package_version(mod_name)
        if ver:
            info.append(("{} 版本".format(pkg_name), ver))
        else:
            info.append(("{} 版本".format(pkg_name), "✗ 未安装"))

    # ── 系统运行时库（Linux 关键项，缺失会导致 Qt 启动失败）──
    if sys.platform != "win32":
        for label, status in _check_system_libs():
            info.append(("系统库: {}".format(label), status))

    # ── 格式转换工具 ──
    if sys.platform != "win32":
        try:
            from app.converter_linux import find_soffice, find_wps
            soffice = find_soffice()
            wps = find_wps()
            info.append(("LibreOffice", soffice if soffice else "未找到"))
            info.append(("WPS Office", wps if wps else "未找到"))
        except Exception:
            pass

    # ── 当前工作目录 ── (用户名段脱敏，避免泄露人名/涉密目录)
    try:
        from app.redact import mask_home
    except Exception:
        def mask_home(p):
            return p
    info.append(("工作目录", mask_home(os.getcwd())))
    info.append(("用户目录", mask_home(os.path.expanduser("~"))))

    # ── 磁盘空间 ──
    disk_info = _disk_usage(os.path.expanduser("~"))
    if disk_info:
        info.append(("HOME 磁盘空间", disk_info))

    # ── locale ──
    info.append(("LANG", os.environ.get("LANG", "未设置")))
    info.append(("LC_ALL", os.environ.get("LC_ALL", "未设置")))

    # ── 内存 ──
    mem = _memory_info()
    if mem:
        info.append(("系统内存", mem))

    return info


def format_diagnostic_text(extra_lines=None):
    """将诊断信息格式化为可复制的文本块"""
    lines = []
    lines.append("=" * 62)
    lines.append("  DocFormat Pro 诊断报告")
    lines.append("=" * 62)
    lines.append("")

    for label, value in collect_system_info():
        lines.append("  {}: {}".format(label, value))

    if extra_lines:
        lines.append("")
        lines.append("-" * 62)
        for line in extra_lines:
            lines.append("  {}".format(line))

    lines.append("")
    lines.append("=" * 62)
    return "\n".join(lines)


def _linux_distro():
    """Linux 发行版信息"""
    for f in ["/etc/os-release", "/etc/lsb-release"]:
        try:
            with open(f) as fh:
                content = fh.read()
            for line in content.splitlines():
                if line.startswith("PRETTY_NAME="):
                    return line.split("=", 1)[1].strip('"')
                if line.startswith("DISTRIB_DESCRIPTION="):
                    return line.split("=", 1)[1].strip('"')
        except Exception:
            pass
    return "未知"


def _glibc_version():
    """获取 glibc 版本（Linux）"""
    import ctypes
    try:
        libc = ctypes.CDLL("libc.so.6")
        buf = ctypes.create_string_buffer(256)
        conf = ctypes.c_int(0)
        # confstr(_CS_GNU_LIBC_VERSION, buf, sizeof(buf))
        res = ctypes.CDLL("libc.so.6", use_errno=True).confstr(
            ctypes.c_int(2), buf, ctypes.c_size_t(256))
        if res > 0:
            return buf.value.decode("utf-8", errors="replace")
    except Exception:
        pass

    # 回退：通过 popen 调 ldd
    import subprocess
    try:
        out = subprocess.check_output(
            ["ldd", "--version"], stderr=subprocess.STDOUT, timeout=5
        ).decode("utf-8", errors="replace")
        for line in out.splitlines():
            if "ldd" in line or "libc" in line or "GLIBC" in line.upper():
                return line.strip()
    except Exception:
        pass
    return ""


def _package_version(module_name):
    """获取 Python 包的版本号"""
    try:
        mod = __import__(module_name)
        # 尝试 __version__ / VERSION
        for attr in ("__version__", "VERSION", "version"):
            v = getattr(mod, attr, None)
            if v:
                if isinstance(v, tuple):
                    return ".".join(str(x) for x in v)
                return str(v)
        # PyQt5 不行就试 QtCore
        if module_name == "PyQt5":
            from PyQt5.QtCore import QT_VERSION_STR
            return "Qt {}".format(QT_VERSION_STR)
        return "(已安装，版本未知)"
    except ImportError:
        return ""
    except Exception:
        return "(检测失败)"


def _disk_usage(path):
    """返回路径所在磁盘的使用情况"""
    try:
        import shutil
        u = shutil.disk_usage(path)
        total_gb = u.total / (1024**3)
        free_gb = u.free / (1024**3)
        return "{:.1f} GB 可用 / {:.1f} GB 总量".format(free_gb, total_gb)
    except Exception:
        return ""


def _memory_info():
    """获取系统内存信息"""
    # Linux
    try:
        with open("/proc/meminfo") as f:
            for line in f:
                if "MemTotal" in line:
                    kb = int(line.split(":")[1].strip().split()[0])
                    gb = kb / (1024**2)
                    return "{:.1f} GB".format(gb)
    except Exception:
        pass
    # Windows
    try:
        import ctypes
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        mem = MEMORYSTATUSEX()
        mem.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(mem))
        gb = mem.ullTotalPhys / (1024**3)
        return "{:.1f} GB".format(gb)
    except Exception:
        pass
    return ""


def _check_system_libs():
    """检测关键系统运行时库（Linux）。

    返回 [(库名, 状态)] 列表，状态为 "✓" / "✗ 缺失" / "⚠ 不确定"。
    这些库是 PyInstaller 打包的 Qt5 应用启动所必需的。
    """
    if sys.platform == "win32":
        return []

    import ctypes

    # 关键库及其说明（不含 glibc 本身，因为 glibc 是 ABI 边界无法捆绑）
    CRITICAL_LIBS = [
        # Qt5 核心 — 缺失任何一个 GUI 都拉不起来
        ("libQt5Widgets.so.5", "Qt5 Widgets"),
        ("libQt5Gui.so.5", "Qt5 GUI"),
        ("libQt5Core.so.5", "Qt5 Core"),
        ("libQt5PrintSupport.so.5", "Qt5 打印"),

        # XCB — Qt 在 Linux 上的显示后端
        ("libxcb.so.1", "X11 协议核心"),
        ("libxcb-util.so.1", "XCB 工具"),
        ("libxcb-icccm.so.4", "XCB ICCCM"),
        ("libxcb-image.so.0", "XCB 图像"),
        ("libxcb-keysyms.so.1", "XCB 按键"),
        ("libxcb-randr.so.0", "XCB 分辨率"),
        ("libxcb-render.so.0", "XCB 渲染"),
        ("libxcb-shape.so.0", "XCB 形状"),
        ("libxcb-shm.so.0", "XCB 共享内存"),
        ("libxcb-sync.so.1", "XCB 同步"),
        ("libxcb-xfixes.so.0", "XCB 修复扩展"),
        ("libxcb-xinerama.so.0", "XCB 多屏"),
        ("libxcb-xkb.so.1", "XCB 键盘"),

        # 字体 — 中文渲染依赖
        ("libfontconfig.so.1", "字体配置"),
        ("libfreetype.so.6", "字体渲染"),

        # 图形
        ("libGL.so.1", "OpenGL"),
        ("libEGL.so.1", "EGL"),

        # GLib — Qt 事件循环依赖
        ("libglib-2.0.so.0", "GLib"),
        ("libgobject-2.0.so.0", "GObject"),

        # C++ 运行时
        ("libstdc++.so.6", "C++ 标准库"),

        # D-Bus — Linux 桌面通信
        ("libdbus-1.so.3", "D-Bus"),
    ]

    results = []
    for soname, desc in CRITICAL_LIBS:
        try:
            ctypes.CDLL(soname)
            results.append(("{} ({})".format(desc, soname), "✓"))
        except OSError:
            results.append(("{} ({})".format(desc, soname), "✗ 缺失"))
        except Exception:
            results.append(("{} ({})".format(desc, soname), "⚠ 不确定"))

    return results


# ================================================================
# 崩溃报告对话框（纯 PyQt5，不依赖 app 其他模块）
# ================================================================

def show_crash_dialog(exc_type, exc_value, exc_tb):
    """在崩溃时显示诊断对话框。可独立调用，即使 app 完全损坏也能工作。"""
    tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)

    if _try_show_qt_crash_dialog(tb_lines):
        return

    # 如果连 Qt 都起不来，至少打印到终端 / 写文件
    _fallback_output(tb_lines)


def _try_show_qt_crash_dialog(tb_lines):
    """尝试用 PyQt5 弹出崩溃报告窗口，失败返回 False"""
    try:
        from PyQt5.QtWidgets import (
            QApplication, QDialog, QVBoxLayout, QHBoxLayout,
            QTextEdit, QPushButton, QLabel, QMessageBox,
        )
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QFont

        # 确保有 QApplication 实例
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv[:1] if len(sys.argv) > 0 else ["diagnostic"])

        dlg = QDialog()
        dlg.setWindowTitle("DocFormat Pro — 诊断报告")
        dlg.resize(720, 540)
        dlg.setMinimumSize(600, 420)
        # 独立配色：不依赖外部主题 QSS，在任何系统上都清晰可读
        dlg.setStyleSheet("""
            QDialog { background: #FFFFFF; }
            QLabel { color: #333333; font-size: 13px; }
            QPushButton { border: 1px solid #BBB; border-radius: 4px;
                padding: 6px 16px; color: #333333; background: #F5F5F5; }
            QPushButton:hover { background: #E8E8E8; }
        """)

        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)

        # 标题区
        title = QLabel("程序运行遇到问题")
        title.setStyleSheet("QLabel { font-size: 16px; font-weight: bold; color: #C0392B; }")
        layout.addWidget(title)

        desc = QLabel(
            "请点击下方「复制全部信息」按钮，将诊断内容发送给开发者。\n"
            "这会帮助我们快速定位问题。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("QLabel { color: #555555; }")
        layout.addWidget(desc)

        # 诊断文本区
        editor = QTextEdit()
        editor.setReadOnly(True)
        editor.setAcceptRichText(False)
        editor.setFont(QFont("Courier New, monospace", 10))
        editor.setStyleSheet(
            "QTextEdit { background: #FAFAFA; color: #222222; "
            "border: 1px solid #CCC; border-radius: 4px; padding: 8px; }"
        )

        # 拼装内容：错误信息 + 系统信息
        lines = []
        lines.append("【错误信息】")
        lines.append("")
        for line in tb_lines:
            lines.append(line.rstrip("\n"))
        lines.append("")
        lines.append("【系统信息】")
        lines.append("")
        lines.append(format_diagnostic_text())
        editor.setPlainText("\n".join(lines))
        layout.addWidget(editor, 1)

        # 按钮区
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        copy_btn = QPushButton("复制全部信息")
        copy_btn.setMinimumHeight(36)
        copy_btn.setStyleSheet(
            "QPushButton { background: #2980B9; color: white; border: none; "
            "border-radius: 4px; padding: 8px 24px; font-size: 14px; font-weight: bold; }"
            "QPushButton:hover { background: #3498DB; }"
        )
        def do_copy():
            editor.selectAll()
            editor.copy()
            cursor = editor.textCursor()
            cursor.clearSelection()
            editor.setTextCursor(cursor)
            QMessageBox.information(dlg, "已复制", "诊断信息已复制到剪贴板，请粘贴发送给开发者。")
        copy_btn.clicked.connect(do_copy)
        btn_layout.addWidget(copy_btn)

        close_btn = QPushButton("关闭")
        close_btn.setMinimumHeight(36)
        close_btn.clicked.connect(dlg.reject)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        dlg.exec_()
        return True
    except Exception:
        return False


def _fallback_output(tb_lines):
    """Qt 都起不来时的兜底输出"""
    text = "\n".join(tb_lines) + "\n\n" + format_diagnostic_text()

    # 尝试写文件到桌面或 HOME
    for base in [os.path.expanduser("~/Desktop"), os.path.expanduser("~"), "/tmp"]:
        try:
            fpath = os.path.join(base, "DocFormatPro_诊断报告.txt")
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(text)
            print("\n诊断报告已写入: {}".format(fpath), file=sys.stderr)
            print("请将此文件内容发送给开发者。", file=sys.stderr)
            break
        except Exception:
            continue

    # 也打印到 stderr
    print("\n" + "=" * 60, file=sys.stderr)
    print(text, file=sys.stderr)
    print("=" * 60, file=sys.stderr)


# ================================================================
# 主动诊断功能（在软件正常运行时可手动触发的诊断）
# ================================================================

def show_diagnostic_dialog(parent=None):
    """手动触发的诊断窗口，显示系统信息，可复制。从软件内部调用。"""
    try:
        from PyQt5.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout,
            QTextEdit, QPushButton, QLabel, QMessageBox,
        )
        from PyQt5.QtGui import QFont

        dlg = QDialog(parent)
        dlg.setWindowTitle("系统诊断信息")
        dlg.resize(700, 500)
        dlg.setMinimumSize(560, 380)
        # 独立配色：不依赖外部主题
        dlg.setStyleSheet("""
            QDialog { background: #FFFFFF; }
            QLabel { color: #333333; font-size: 13px; }
            QPushButton { border: 1px solid #BBB; border-radius: 4px;
                padding: 6px 16px; color: #333333; background: #F5F5F5; }
            QPushButton:hover { background: #E8E8E8; }
        """)

        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)

        title = QLabel("系统诊断信息")
        title.setStyleSheet("QLabel { font-size: 16px; font-weight: bold; color: #2C3E50; }")
        layout.addWidget(title)

        hint = QLabel(
            "以下信息描述了当前系统的运行环境。\n"
            "如软件运行异常，请复制以下内容反馈给开发者。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("QLabel { color: #555555; }")
        layout.addWidget(hint)

        editor = QTextEdit()
        editor.setReadOnly(True)
        editor.setAcceptRichText(False)
        editor.setFont(QFont("Courier New, monospace", 10))
        editor.setStyleSheet(
            "QTextEdit { background: #FAFAFA; color: #222222; "
            "border: 1px solid #CCC; border-radius: 4px; padding: 8px; }"
        )
        editor.setPlainText(format_diagnostic_text())
        layout.addWidget(editor, 1)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        copy_btn = QPushButton("复制全部信息")
        copy_btn.setMinimumHeight(36)
        copy_btn.setStyleSheet(
            "QPushButton { background: #2980B9; color: white; border: none; "
            "border-radius: 4px; padding: 8px 24px; font-size: 14px; font-weight: bold; }"
            "QPushButton:hover { background: #3498DB; }"
        )
        def do_copy():
            editor.selectAll()
            editor.copy()
            cursor = editor.textCursor()
            cursor.clearSelection()
            editor.setTextCursor(cursor)
            QMessageBox.information(dlg, "已复制", "诊断信息已复制到剪贴板。")
        copy_btn.clicked.connect(do_copy)
        btn_layout.addWidget(copy_btn)

        close_btn = QPushButton("关闭")
        close_btn.setMinimumHeight(36)
        close_btn.clicked.connect(dlg.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        dlg.exec_()
    except Exception:
        pass
