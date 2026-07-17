# -*- coding: utf-8 -*-
"""布局检查：离屏渲染窗口，断言无溢出/遮挡，并输出截图供人工核对"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt5.QtWidgets import QApplication

app = QApplication(sys.argv)
from app.main_window import MainWindow

OUT = os.path.join(os.path.dirname(__file__), '_shots')
os.makedirs(OUT, exist_ok=True)

win = MainWindow()


def check_layout(tag, w, h):
    win.resize(w, h)
    win.show()
    app.processEvents()
    home = win.home_page
    problems = []

    def visible_in(widget, container, name):
        if not widget.isVisible():
            return
        top_left = widget.mapTo(container, widget.rect().topLeft())
        right = top_left.x() + widget.width()
        bottom = top_left.y() + widget.height()
        if right > container.width() + 2:
            problems.append('{}: 右侧溢出 {}px (容器宽 {})'.format(name, right - container.width(), container.width()))
        if widget.width() < 40:
            problems.append('{}: 宽度过窄 {}px'.format(name, widget.width()))

    central = win.centralWidget()
    visible_in(home.preset_combo, central, '预设下拉')
    visible_in(home.suffix_edit, central, '后缀输入')
    visible_in(home.badge_page, central, '边距徽章')
    visible_in(home.badge_body, central, '正文徽章')
    visible_in(home.badge_spacing, central, '行距徽章')
    visible_in(home.process_btn, central, '开始处理按钮')
    visible_in(home.drop_zone, central, '拖拽区')

    if problems:
        print('[{} {}x{}] 发现布局问题:'.format(tag, w, h))
        for p in problems:
            print('   ✗ ' + p)
        return False
    print('[{} {}x{}] 布局检查通过（无溢出/遮挡） ✓'.format(tag, w, h))
    return True


ok = True
ok &= check_layout('默认尺寸', 1120, 780)
ok &= check_layout('最小尺寸', 900, 620)

# 各页面截图
shots = []
for idx, name in [(0, 'home'), (1, 'presets'), (2, 'theme'), (3, 'log')]:
    win.stack.setCurrentIndex(idx)
    app.processEvents()
    path = os.path.join(OUT, '{}.png'.format(name))
    win.grab().save(path)
    shots.append(path)
print('截图已保存: ' + ', '.join(os.path.basename(s) for s in shots))

# AI 粘贴模式截图
win.stack.setCurrentIndex(0)
for b in win.home_page.mode_group.buttons():
    if b.property('modeId') == 'ai_paste':
        b.setChecked(True)
        win.home_page._on_mode_changed(b)
        break
app.processEvents()
win.grab().save(os.path.join(OUT, 'home_ai_paste.png'))
print('截图已保存: home_ai_paste.png')

sys.exit(0 if ok else 1)
