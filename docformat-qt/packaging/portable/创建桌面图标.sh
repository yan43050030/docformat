#!/bin/sh
# 给当前用户装一个桌面图标 —— 全程不需要管理员密码。
#
# 只往这两个地方写，都在自己家目录里：
#   ~/.local/share/applications  （开始菜单/应用列表）
#   ~/桌面 或 ~/Desktop          （桌面图标，麒麟 UKUI 两个名字都可能有）
# 想撤销就把这两处的 docformat-pro.desktop 删掉。
set -e
DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

if [ ! -f "$DIR/DocFormatPro" ]; then
    echo "找不到主程序，请在解压出来的目录里运行本脚本。" >&2
    exit 1
fi
chmod u+rx "$DIR/DocFormatPro" "$DIR/启动.sh" 2>/dev/null || true

APPS="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
mkdir -p "$APPS"
ENTRY="$APPS/docformat-pro.desktop"
ICON="$DIR/icon.png"
[ -f "$ICON" ] || ICON=

cat > "$ENTRY" << EOF
[Desktop Entry]
Type=Application
Name=DocFormat Pro
Name[zh_CN]=DocFormat Pro 公文排版
Comment=公文格式自动排版工具 (GB/T 9704-2012)
Exec="$DIR/启动.sh"
Path=$DIR
Icon=$ICON
Terminal=false
Categories=Office;WordProcessor;
StartupNotify=true
EOF
chmod u+rx "$ENTRY"

for D in "$HOME/桌面" "$HOME/Desktop"; do
    if [ -d "$D" ]; then
        cp "$ENTRY" "$D/docformat-pro.desktop" && chmod u+rx "$D/docformat-pro.desktop"
        # 麒麟/GNOME 新版要求桌面上的快捷方式被标记为"受信任"才肯直接启动
        command -v gio >/dev/null 2>&1 && \
            gio set "$D/docformat-pro.desktop" metadata::trusted true 2>/dev/null || true
    fi
done

command -v update-desktop-database >/dev/null 2>&1 && \
    update-desktop-database "$APPS" 2>/dev/null || true

echo "桌面图标已创建。"
echo "  应用列表：$ENTRY"
echo "  程序目录：$DIR"
echo "注意：这个图标指向当前位置，之后**不要移动或改名**这个文件夹，"
echo "      否则要重新运行一次本脚本。"
