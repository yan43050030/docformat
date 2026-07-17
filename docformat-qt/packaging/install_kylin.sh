#!/usr/bin/env bash
# 麒麟 / 统信 UOS 源码部署脚本（推荐方式，规避 glibc 兼容问题）
# 用法: bash install_kylin.sh
set -e
cd "$(dirname "$0")/.."
APP_DIR="$(pwd)"

echo "=== DocFormat Pro 麒麟/UOS 安装 ==="

# 1. 系统源安装依赖（信创内网可用，无需 pip 联网）
if command -v apt >/dev/null 2>&1; then
    echo "[1/3] 通过 apt 安装依赖 (python3-pyqt5 / python3-docx)..."
    sudo apt install -y python3-pyqt5 python3-docx || {
        echo "python3-docx 不在源内时，回退 pip 安装："
        sudo apt install -y python3-pip
        pip3 install --user python-docx
    }
elif command -v yum >/dev/null 2>&1; then
    echo "[1/3] 通过 yum 安装依赖..."
    sudo yum install -y python3-qt5 || true
    pip3 install --user python-docx
else
    echo "未识别的包管理器，请手动安装 python3-pyqt5 与 python-docx"
fi

# 2. 可选：LibreOffice（处理 .doc/.wps 输入需要）
if ! command -v soffice >/dev/null 2>&1 && ! command -v libreoffice >/dev/null 2>&1; then
    echo "[提示] 未检测到 LibreOffice：.docx 处理不受影响；"
    echo "       如需处理 .doc/.wps 输入，请执行: sudo apt install libreoffice-writer"
fi

# 3. 安装桌面图标（实现"点击启动 GUI"）
echo "[2/3] 安装桌面入口..."
DESKTOP_DIR="$HOME/.local/share/applications"
mkdir -p "$DESKTOP_DIR"
sed "s|@APP_DIR@|$APP_DIR|g" packaging/docformat.desktop > "$DESKTOP_DIR/docformat-pro.desktop"
chmod +x "$DESKTOP_DIR/docformat-pro.desktop"
# 麒麟 UKUI 桌面通常也读取 ~/桌面
for D in "$HOME/桌面" "$HOME/Desktop"; do
    if [ -d "$D" ]; then
        cp "$DESKTOP_DIR/docformat-pro.desktop" "$D/" && chmod +x "$D/docformat-pro.desktop"
    fi
done

echo "[3/3] 验证启动..."
python3 -c "import PyQt5.QtWidgets, docx; print('依赖检查通过')"

echo ""
echo "安装完成！启动方式："
echo "  · 双击桌面「DocFormat Pro」图标"
echo "  · 或命令行: python3 $APP_DIR/main.py"
