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
<tr><td><b>套打填写</b></td><td>把内容打到已印好红头的纸上；含模板可视化编辑、批量套打、套头库、直接打印</td></tr>
<tr><td><b>文书起草</b></td><td>选文书模板 → 填空 → 直接产出排好版的公文</td></tr>
<tr><td><b>文书模板制作</b></td><td>把一份历史公文挖空成可复用的起草模板</td></tr>
</table>

<h3>格式处理的几种模式</h3>
<table cellpadding="4">
<tr><td><b>智能一键处理</b></td><td>标点修复 + 排版规范 + 样式清洗，一步到位</td></tr>
<tr><td><b>公文合规检查</b></td><td>查<b>版式</b>：逐段对照预设核对字体字号行距边距，勾选认可的偏差精准修正</td></tr>
<tr><td><b>公文用语检查</b></td><td>查<b>文字</b>：错别字、数字用法、文种与结语搭配；只查有明文规定的，文风不管。
改动一律写成 Word 修订，逐条看过认可才落笔</td></tr>
<tr><td><b>格式清洗</b></td><td>清掉看不见的脏格式，专治排版怪问题</td></tr>
<tr><td><b>标点修复</b></td><td>只规范标点，不动版式</td></tr>
<tr><td><b>生成目录</b></td><td>按标题层级生成目录，可选自动域或静态页</td></tr>
<tr><td><b>AI 粘贴生成</b></td><td>粘贴 AI 写的文本/Markdown，直接产出规范 docx</td></tr>
</table>

<h3>键盘快捷键</h3>
<table cellpadding="4">
<tr><td><b>Ctrl+1 ~ Ctrl+7</b></td><td>切换页面</td></tr>
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
<p><b>改名归档太烦？</b> 选好文件后点「转换与工具 → 归档命名」：
文号、标题、日期、密级软件都已经认过，按命名式一次改好，还能把每份的信息
追加到 CSV 登记台账。默认是<b>复制</b>，原件留在原地；重名自动加 (2) 不覆盖。
注意台账里是明文的标题和密级，存放位置请按其中最高密级斟酌。</p>
<p><b>怕漏标密级？</b> 用「公文合规检查」，里面有<b>密级标注</b>一组（默认开）：
有份号却没密级、正文写着"注意保密"却没密级行、密级缺保密期限或写法不规范、
密级位置不对，都会报出来，并可一键补标。
补标时<b>密级和保密期限由你选</b>——标错密级比不标更麻烦，软件只负责排到版头正确位置。
多个文件一起拖进来即可批量筛查。</p>
<p><b>套打怎么调准？</b> 在「套打填写」页点「打印…」→「对位测试页」，
打在<b>普通白纸</b>上：每栏标着它会印到距纸左边多少厘米，预印栏目名也用浅灰显出来。
和真实预印纸对光一叠，哪一栏偏了一眼看得见，再用「打印位置微调」或
「可视化编辑」改到位——白纸比预印纸便宜。</p>
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
