# DocFormat Pro (Qt)

公文格式自动排版工具（GB/T 9704-2012），PyQt5 桌面应用。
支持 **Windows** 与 **麒麟 / 统信 UOS 等国产系统**（含飞腾/鲲鹏 ARM 架构）。

排版引擎复用自 [docformat-gui](https://github.com/KaguraNanaga/docformat-gui)（MIT），本项目在其之上提供：
Qt 原生文件拖拽（麒麟 ARM 可用）、Linux 下 .doc/.wps 输入（LibreOffice 转换链）、
信创内网系统源部署、可视化模板编辑与持久化、AI 粘贴生成公文、四套界面主题。

## 功能

- **智能一键处理**：标点修复 → 排版规范（页边距/字体/字号/缩进/行距/页码/表格）→ 样式清洗
- **排版前后对比预览**：左右分栏显示原文与模拟排版效果（含每段识别类型标注），确认无误后才真正写文件；
  **点击段落类型标签可手动指定**该段是标题/正文/附件等，手动调整在实际排版时生效
- **格式诊断**：只检查不修改（含表格单元格），报告可导出 TXT，可一键转入智能修复
- **标点修复**：仅规范中英文标点，不动版式；英文撇号（it's）、跨段引号、英文单词间空格均受保护
- **发文字号**：自动识别"××发〔2026〕12号"并按仿宋居中排版，字体字号可在模板中调整
- **修订模式输出**：可选把所有改动写成 Word 修订（审阅中可见、可逐条接受/拒绝）
- **AI 粘贴生成**：粘贴 AI 写的文本/Markdown，自动清洗标记并产出规范公文 docx
- **密级标识**：自动识别文首"秘密★1年 / 机密★3年 / 绝密"等定密行（仅文档前 3 个非空段、整行匹配才认定），按三号黑体顶格排版，可在模板中自定义
- **识别规则可自定义**：密级/发文字号/各级标题/附件/署名/日期/结束语/主送机关共 11 类要素的识别规则
  均可在"预设方案 → 识别规则"中修改。常用方案（如法律条文"第一条"）直接下拉选择，无需写正则；
  进阶用户可选"自定义"填正则，并用**实时测试框**粘贴一行文字立即验证识别结果，非法正则自动回退
- **模板系统**：内置四套模板，默认为**公文格式（图解版·22行28字）**——A4，上3.8/下3.3/左右2.8cm
  边距，方正小标宋_GBK 二号标题、方正黑体/楷体/仿宋_GBK 三号加粗各级标题与正文，行距固定 28 磅，
  文档网格每页 22 行 × 每行 28 字，页码四号半角宋体，落款按"成文日期右空 2 字、
  署名与日期首字错 2 字"自动对位；另含通用公文（GB/T 9704）/学术/法律三套；
  自定义模板可视化编辑、即时保存、JSON 导入导出
- **批量处理**：多文件/整个文件夹拖入，逐文件状态标记（✓/✗）与日志，完成后一键打开输出位置；
  输出在原文件旁加后缀，不覆盖原文件，重名时自动追加序号；批量转换 .doc/.wps 时 Office/WPS 实例整批复用
- **高分屏适配**：125% / 150% 等非整数缩放按真实比例渲染（PassThrough）

## Windows 使用

```bat
pip install -r requirements.txt
双击 启动.bat        （或 python main.py）
```

打包独立 exe（免 Python 环境分发）：

```bat
pip install pyinstaller
packaging\build_windows.bat     → dist\DocFormatPro.exe
```

.doc/.wps 输入依赖本机 WPS 或 Microsoft Word（自动检测，WPS 优先）。

## 麒麟 / UOS 使用（推荐源码部署）

```bash
bash packaging/install_kylin.sh
```

脚本会通过**系统源** `apt install python3-pyqt5 python3-docx`（内网可用，无 glibc 兼容问题），
并安装桌面图标，之后双击图标即可启动。

- .doc/.wps 输入需要 LibreOffice：`sudo apt install libreoffice-writer`（仅处理 .docx 则无需安装）
- 二进制打包（可选）：在目标架构机器上执行 `bash packaging/build_linux.sh`

## 目录结构

```
main.py            入口
app/               Qt 界面（主窗口/四页面/拖拽控件/主题/模板管理/处理线程）
scripts/           排版引擎（来自 docformat-gui，MIT）
packaging/         Windows/Linux 打包与麒麟安装脚本
smoke_test.py      引擎冒烟测试（8 项）
gui_test.py        GUI 自动化测试（12 项，offscreen 运行）
```

用户模板存储位置：Windows `%APPDATA%\DocFormatPro\templates.json`，Linux `~/.config/DocFormatPro/templates.json`。

## 测试

```bash
python smoke_test.py    # 引擎：三模式 + 自定义模板 + AI 生成
python gui_test.py      # GUI：拖拽事件 + 处理链路 + 模板持久化 + 主题
```

## 许可

MIT，引擎部分版权归原项目所有，见 [LICENSE](LICENSE) 与 [LICENSE.docformat-gui](LICENSE.docformat-gui)。
