# 公文排版工具（DocFormat）

公文格式自动排版工具集，遵循 GB/T 9704-2012《党政机关公文格式》。

一键完成：标点修复（英文标点→中文规范标点）、版式规范（页边距/字体/字号/首行缩进/行距/页码/表格）、样式清洗，并支持自定义排版模板、批量处理与文件拖拽。

## 仓库结构：两个项目

### `docformat-qt/` — 主项目（推荐使用）✅

**Python + PyQt5 桌面应用，功能完整可用。**

- 跨平台：Windows 7+ 与 **麒麟 / 统信 UOS 等国产系统**（含飞腾/鲲鹏 ARM，PyQt5 走系统源安装，信创内网可部署）
- 四种模式：智能一键处理 / 格式诊断 / 标点修复 / **AI 粘贴生成**（粘贴 AI 写的 Markdown 直接产出规范公文）
- **排版前后对比预览**：左右分栏 + 段落类型标注，确认后才执行排版
- 自定义排版模板：可视化编辑全部参数（含密级/标题**识别规则正则**），修改即保存，支持 JSON 导入导出
- 密级标识（左上角定密，如"秘密★1年"）自动识别与排版
- Qt 原生文件拖拽（支持整个文件夹拖入）；高分屏 125%/150% 缩放清晰渲染
- Word/WPS 文件：`.docx` 直接处理；`.doc`/`.wps` 在 Windows 走 WPS/Word COM 转换，在 Linux 走 LibreOffice 转换链
- Windows 可打包为单文件 `DocFormatPro.exe`（见 `packaging/build_windows.bat`）
- 自动化测试：`smoke_test.py`（引擎 6 项）、`gui_test.py`（界面 11 项）、`layout_check.py`（布局）

使用方法见 [docformat-qt/README.md](docformat-qt/README.md)。

### `docformat-pro/` — 早期 UI 原型（仅作界面参考）⚠️

**Tauri 2 + React 19 的界面原型，核心功能未实现，不可实际使用。**

- 完成度：界面/交互/主题完整；但排版引擎为空实现、"开始处理"为模拟进度、模板不持久化、Tauri 壳未编译成功
- 保留原因：docformat-qt 的界面设计（侧边栏四页布局、纸质感配色）复刻自此原型，留档供后续 Web/Tauri 路线参考

## 致谢与引擎来源

排版核心引擎（`docformat-qt/scripts/` 目录：formatter / analyzer / punctuation / converter）来自开源项目：

**[KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui)**（MIT License）

本仓库在其引擎之上重写了 GUI（tkinter → PyQt5），并新增：麒麟 ARM 可靠拖拽、Linux 下 .doc/.wps 输入转换链、信创系统源部署方案、可视化模板编辑器、AI 粘贴生成、多主题界面。原始许可证见 [docformat-qt/LICENSE.docformat-gui](docformat-qt/LICENSE.docformat-gui)。

## 快速开始（Windows）

```bat
cd docformat-qt
pip install -r requirements.txt
启动.bat                          :: 源码运行
packaging\build_windows.bat       :: 或打包独立 exe → dist\DocFormatPro.exe
```

## 快速开始（麒麟 / UOS）

```bash
cd docformat-qt
bash packaging/install_kylin.sh   # 系统源装依赖 + 桌面图标
```

## 许可

MIT — 详见 [LICENSE](docformat-qt/LICENSE)。
