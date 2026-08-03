# -*- coding: utf-8 -*-
"""套打页：把整套套打功能摆到侧边栏一级入口上。

以前套打挂在首页"转换与工具"那行小按钮里，和"转为 docx"并排——可它早已
不是一个小工具了：模板库、可视化编辑、批量套打、套头库、打印预检、扫描
自动对位、直接打印，是一整套东西。埋在那儿等于没有。

页面本身不重写界面：直接把 OverprintDialog 以 embedded 形态嵌进来。
弹窗和页面用同一份代码，改一处两处都对。
"""
from PyQt5.QtWidgets import QLabel, QVBoxLayout, QWidget

# 本页在侧边栏里的位置，首页那个入口靠它把人送过来
NAV_INDEX = 2


class OverprintPage(QWidget):
    def __init__(self, parent=None):
        super(OverprintPage, self).__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 16)
        root.setSpacing(10)

        title = QLabel("套打填写")
        title.setProperty("h1", "true")
        sub = QLabel("把内容打到已经印好红头的纸上——纸上原有的内容在模板里存成"
                     "白字，占准位置但不显影，打印机只印黑字。")
        sub.setProperty("muted", "true")
        sub.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(sub)

        from app.overprint_dialog import OverprintDialog
        self.panel = OverprintDialog(self, embedded=True)
        # 以普通控件的身份长在页面里，不再是独立窗口
        self.panel.setWindowFlags(self.panel.windowFlags())
        self.panel.setSizeGripEnabled(False)
        root.addWidget(self.panel, 1)
