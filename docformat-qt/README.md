# DocFormat Pro (Qt)

公文格式自动排版工具（GB/T 9704-2012），PyQt5 桌面应用。
支持 **Windows** 与 **麒麟 / 统信 UOS 等国产系统**（含飞腾/鲲鹏 ARM 架构）。

排版引擎复用自 [docformat-gui](https://github.com/KaguraNanaga/docformat-gui)（MIT），本项目在其之上提供：
Qt 原生文件拖拽（麒麟 ARM 可用）、Linux 下 .doc/.wps 输入（LibreOffice 转换链）、
信创内网系统源部署、可视化模板编辑与持久化、AI 粘贴生成公文、四套界面主题。

## 功能

- **智能一键处理**：标点修复 → 排版规范（页边距/字体/字号/缩进/行距/页码/表格）→ 样式清洗
- **排版前后对比预览**：左右分栏显示原文与模拟排版效果（含每段识别类型标注），确认无误后才真正写文件
- **格式诊断**：只检查不修改，输出问题报告
- **标点修复**：仅规范中英文标点，不动版式
- **AI 粘贴生成**：粘贴 AI 写的文本/Markdown，自动清洗标记并产出规范公文 docx
- **密级标识**：自动识别文首"秘密★1年 / 机密★3年 / 绝密"等定密行（仅文档前 3 个非空段、整行匹配才认定），按三号黑体顶格排版，可在模板中自定义
- **识别规则可自定义**：密级与一至四级标题的识别正则均可在"预设方案 → 识别规则"中修改（如法律条文改为 `^第[一二三四五六七八九十百]+条`），留空即用默认规则，非法正则自动回退
- **模板系统**：内置公文/学术/法律三套模板；自定义模板可视化编辑、即时保存、JSON 导入导出
- **批量处理**：多文件/整个文件夹拖入，逐文件进度与日志；输出在原文件旁加后缀，不覆盖原文件
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
smoke_test.py      引擎冒烟测试（6 项）
gui_test.py        GUI 自动化测试（9 项，offscreen 运行）
```

用户模板存储位置：Windows `%APPDATA%\DocFormatPro\templates.json`，Linux `~/.config/DocFormatPro/templates.json`。

## 测试

```bash
python smoke_test.py    # 引擎：三模式 + 自定义模板 + AI 生成
python gui_test.py      # GUI：拖拽事件 + 处理链路 + 模板持久化 + 主题
```

## 许可

MIT，引擎部分版权归原项目所有，见 [LICENSE](LICENSE) 与 [LICENSE.docformat-gui](LICENSE.docformat-gui)。
