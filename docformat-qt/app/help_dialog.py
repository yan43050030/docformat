# -*- coding: utf-8 -*-
"""内置使用说明（F1 / 侧边栏"使用说明"打开）"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QHBoxLayout, QPushButton, QTextBrowser, QVBoxLayout

_HELP_HTML = """
<h2>DocFormat Pro 使用说明</h2>

<h3>各页面做什么</h3>
<table cellpadding="4">
<tr><td><b>格式处理</b></td><td>把现有 Word/文本文件一键排成规范公文（主战场）</td></tr>
<tr><td><b>版式方案</b></td><td>查看/编辑排版参数与识别规则，管理模板</td></tr>
<tr><td><b>文书起草</b></td><td>选文书模板 → 填空 → 直接产出排好版的公文</td></tr>
<tr><td><b>文书模板制作</b></td><td>把一份历史公文挖空成可复用的起草模板</td></tr>
</table>

<h3>格式处理的四种模式</h3>
<table cellpadding="4">
<tr><td><b>智能一键处理</b></td><td>标点修复 + 排版规范 + 样式清洗，一步到位</td></tr>
<tr><td><b>格式诊断</b></td><td>只检查不修改，输出问题报告（可导出、可一键转修复）</td></tr>
<tr><td><b>标点修复</b></td><td>只规范标点，不动版式</td></tr>
<tr><td><b>AI 粘贴生成</b></td><td>粘贴 AI 写的文本/Markdown，直接产出规范 docx</td></tr>
</table>

<h3>键盘快捷键</h3>
<table cellpadding="4">
<tr><td><b>Ctrl+1 ~ Ctrl+6</b></td><td>切换页面</td></tr>
<tr><td><b>Ctrl+O</b></td><td>选择文件</td></tr>
<tr><td><b>Ctrl+回车</b></td><td>开始处理</td></tr>
<tr><td><b>F1</b></td><td>打开本说明</td></tr>
</table>

<h3>常见问题</h3>
<p><b>提示缺少字体？</b> 输出文档的字体名是正确的，拿到装有方正字体的电脑上
打开即正常；本机想正常显示需安装对应字体。</p>
<p><b>.doc/.wps 处理失败？</b> Windows 需要本机装有 WPS 或 Word；
麒麟/UOS 需要 LibreOffice（<code>sudo apt install libreoffice-writer</code>）。</p>
<p><b>某段落识别错了？</b> 用「预览对比」，点击右侧段落前的类型标签手动指定；
或在「版式方案 → 识别规则」里调整规则（有实时测试框）。</p>
<p><b>输出文件在哪？</b> 与原文件同目录，文件名加后缀；处理完成后点
「打开输出位置」直达。原文件永远不会被覆盖。</p>
"""


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super(HelpDialog, self).__init__(parent)
        self.setWindowTitle("使用说明")
        self.resize(680, 640)
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        view = QTextBrowser()
        view.setOpenExternalLinks(True)
        view.setHtml(_HELP_HTML)
        root.addWidget(view, 1)
        btns = QHBoxLayout()
        btns.addStretch(1)
        close_btn = QPushButton("关闭")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btns.addWidget(close_btn)
        root.addLayout(btns)
