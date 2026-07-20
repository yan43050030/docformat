#!/usr/bin/env bash
# ============================================================================
# DocFormat Pro — 离线运行时依赖提取脚本
# ============================================================================
# 用途：在构建容器（Debian 10 / glibc 2.28）内运行，提取 PyInstaller
#       产物依赖的所有系统 .so 文件，打包为离线依赖应急包。
#
# 用法：bash bundle_runtime_libs.sh <DocFormatPro二进制> [输出目录]
#
# 输出：DocFormatPro-deps-<arch>.tar.gz
#   内含：
#     runtime/          — 所有需要的 .so 文件
#     install_deps_offline.sh — 目标机器上的安装脚本
#
# 这个包只需发布一次，后续软件更新不需要重新发布依赖包。
# ============================================================================
set -e

BIN="${1:?请指定 DocFormatPro 二进制路径}"
OUT_DIR="${2:-.}"
ARCH=$(uname -m)
PKG_NAME="DocFormatPro-deps-$ARCH"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== DocFormat Pro 离线依赖提取 ==="
echo "二进制: $BIN"
echo "架构: $ARCH"
echo "输出: $OUT_DIR/$PKG_NAME.tar.gz"
echo ""

# ── 1. 收集所有动态库依赖 ──
echo "[1/4] 扫描动态库依赖..."

# 用 ldd 列出所有链接的 so，提取路径
# 排除虚拟库（linux-vdso, ld-linux-*），保留真实路径
DEPS_FILE=$(mktemp)
ldd "$BIN" 2>&1 | grep -oP '/[^ ]+\.so[^ ]*' | sort -u > "$DEPS_FILE"

# 递归解析间接依赖（部分 so 的依赖可能不在第一层 ldd 输出中）
# 做两轮扫描确保完整
for round in 1 2; do
    count_before=$(wc -l < "$DEPS_FILE")
    while IFS= read -r so; do
        ldd "$so" 2>/dev/null | grep -oP '/[^ ]+\.so[^ ]*' | sort -u >> "$DEPS_FILE" || true
    done < "$DEPS_FILE"
    sort -u -o "$DEPS_FILE" "$DEPS_FILE"
    count_after=$(wc -l < "$DEPS_FILE")
    echo "  第 $round 轮: $count_before → $count_after 个库"
done

TOTAL=$(wc -l < "$DEPS_FILE")
echo "  共发现 $TOTAL 个依赖库"

# ── 2. 复制 .so 文件到 staging ──
echo "[2/4] 复制库文件..."

STAGING="$OUT_DIR/$PKG_NAME"
rm -rf "$STAGING"
mkdir -p "$STAGING/runtime"

COPIED=0
while IFS= read -r so; do
    if [ -f "$so" ]; then
        # 跟随符号链接复制实际文件
        real=$(readlink -f "$so" 2>/dev/null || echo "$so")
        if [ -f "$real" ]; then
            cp -L "$real" "$STAGING/runtime/" 2>/dev/null && ((COPIED++)) || true
        fi
    fi
done < "$DEPS_FILE"

echo "  已复制 $COPIED 个库文件"

rm -f "$DEPS_FILE"

# ── 3. 生成安装脚本（嵌入在包内） ──
echo "[3/4] 生成离线安装脚本..."

cat > "$STAGING/install_deps_offline.sh" << 'INSTALL_SCRIPT'
#!/usr/bin/env bash
# ============================================================================
# DocFormat Pro — 离线依赖安装脚本
# ============================================================================
# 用法：bash install_deps_offline.sh [DocFormatPro 所在目录]
#
# 默认安装到当前目录下的 runtime/，并创建 DocFormatPro.sh 启动包装器。
# 如果指定了目录参数，则在那个目录安装（依赖必须和二进制在同一父目录）。
#
# 此脚本解决麒麟/统信离线系统缺少 Qt5/xcb/fontconfig 等 .so 的问题。
# ============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET="${1:-$SCRIPT_DIR}"

RUNTIME_SRC="$SCRIPT_DIR/runtime"
RUNTIME_DST="$TARGET/runtime"

if [ ! -d "$RUNTIME_SRC" ]; then
    echo "错误：未找到 runtime/ 目录，请确保在依赖包解压目录内运行此脚本。"
    exit 1
fi

echo "=== DocFormat Pro 离线依赖安装 ==="
echo "源: $RUNTIME_SRC"
echo "目标: $RUNTIME_DST"
echo ""

# 如果目标已存在且不同，询问是否覆盖
if [ -d "$RUNTIME_DST" ] && [ "$RUNTIME_SRC" != "$RUNTIME_DST" ]; then
    echo "目标已存在 runtime/ 目录，将覆盖。"
    rm -rf "$RUNTIME_DST"
fi

if [ "$RUNTIME_SRC" != "$RUNTIME_DST" ]; then
    cp -r "$RUNTIME_SRC" "$RUNTIME_DST"
    echo "已复制 $(ls "$RUNTIME_DST" | wc -l) 个运行时库到 $RUNTIME_DST"
else
    echo "runtime/ 已在目标位置，跳过复制。"
fi

# 创建启动包装脚本
BIN_NAME="DocFormatPro"
BIN_PATH="$TARGET/$BIN_NAME"

WRAPPER="$TARGET/${BIN_NAME}.sh"
cat > "$WRAPPER" << WRAPPER_EOF
#!/usr/bin/env bash
# DocFormat Pro 启动包装器 — 使用离线依赖库
HERE="\$(cd "\$(dirname "\$0")" && pwd)"
export LD_LIBRARY_PATH="\$HERE/runtime:\${LD_LIBRARY_PATH}"
exec "\$HERE/$BIN_NAME" "\$@"
WRAPPER_EOF

chmod +x "$WRAPPER"

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  安装完成！"
echo ""
echo "  启动方式："
echo "    $WRAPPER"
echo ""
echo "  也可以创建桌面图标指向此包装脚本："
echo "    cp packaging/docformat.desktop ~/.local/share/applications/"
echo "    sed -i 's|Exec=.*|Exec=$WRAPPER|' ~/.local/share/applications/docformat.desktop"
echo ""
echo "  如需完全免依赖运行，请确保 DocFormatPro 二进制也在同一目录。"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
INSTALL_SCRIPT

chmod +x "$STAGING/install_deps_offline.sh"

# ── 4. 打包 ──
echo "[4/4] 打包..."

tar -czf "$OUT_DIR/$PKG_NAME.tar.gz" -C "$OUT_DIR" "$PKG_NAME"

# 显示结果
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  离线依赖包已生成"
echo "  文件: $OUT_DIR/$PKG_NAME.tar.gz"
echo "  大小: $(du -h "$OUT_DIR/$PKG_NAME.tar.gz" | cut -f1)"
echo "  包含: $(ls "$STAGING/runtime" | wc -l) 个 .so 文件"
echo ""
echo "  使用方式（在离线麒麟/UOS 上）："
echo "    1. tar xzf $PKG_NAME.tar.gz"
echo "    2. bash $PKG_NAME/install_deps_offline.sh /path/to/DocFormatPro/"
echo "    3. 用 DocFormatPro.sh 启动"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# 清理 staging（保留 tar.gz）
rm -rf "$STAGING"
