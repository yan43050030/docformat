#!/usr/bin/env bash
# ============================================================================
# DocFormat Pro — 国产 Linux 系统兼容性诊断脚本
# ============================================================================
# 用法：
#   1. 拷贝到目标机器上
#   2. chmod +x diagnose.sh && bash diagnose.sh [可选的二进制路径]
#   3. 将输出反馈给开发者
#
# 检测内容：
#   · 系统版本 / 内核 / 架构 / glibc / 图形环境
#   · 二进制模式：ldd 查缺失 so、试跑看错误
#   · 源码模式：Python 版本、PyQt5 / python-docx / lxml 是否就绪
#   · LibreOffice（.doc/.wps 转换依赖）
#   · 权限 / 磁盘空间 / 中文 locale
# ============================================================================

set -e

BINARY="${1:-}"

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m'       # No Color
BOLD='\033[1m'

pass_msg()  { echo -e "  ${GREEN}✓${NC} $1"; }
warn_msg()  { echo -e "  ${YELLOW}⚠${NC} $1"; }
fail_msg()  { echo -e "  ${RED}✗${NC} $1"; }
info_msg()  { echo -e "  ${CYAN}→${NC} $1"; }
section()   { echo ""; echo -e "${BOLD}━━━ $1 ━━━${NC}"; }

failures=0

# ============================================================
section "1. 系统基本信息"
# ============================================================

echo -n "主机名: "; hostname 2>/dev/null || echo "(未知)"

# OS 信息
if [ -f /etc/os-release ]; then
    . /etc/os-release
    echo "操作系统: $PRETTY_NAME"
    info_msg "ID=${ID}, VERSION_ID=${VERSION_ID}"
elif [ -f /etc/lsb-release ]; then
    . /etc/lsb-release
    echo "操作系统: ${DISTRIB_DESCRIPTION:-未知}"
else
    warn_msg "无法确定操作系统版本（缺少 /etc/os-release）"
fi

# 内核
echo "内核版本: $(uname -r)"

# 架构
ARCH=$(uname -m)
echo "CPU 架构: $ARCH"
case "$ARCH" in
    x86_64)  pass_msg "架构：x86_64 — 主流 Intel/AMD/兆芯/海光" ;;
    aarch64) pass_msg "架构：ARM64 — 鲲鹏 920 / 飞腾 2000" ;;
    loongarch64) warn_msg "架构：LoongArch64 — 龙芯（当前无预编译包，需源码安装）" ;;
    *)       warn_msg "架构：$ARCH — 非标准架构，需确认兼容性" ;;
esac

# glibc —— 这是 PyInstaller 打包的关键约束
if command -v ldd &>/dev/null; then
    LDD_VERSION=$(ldd --version 2>&1 | head -1 | grep -oP '[\d]+\.[\d]+' | head -1 || echo "?")
    echo "glibc 版本: $LDD_VERSION"
    # glibc 2.28 ≈ Debian 10 / 麒麟 V10
    # glibc 2.31 ≈ Ubuntu 20.04 / UOS 20
    if [ -n "$LDD_VERSION" ] && [ "$LDD_VERSION" != "?" ]; then
        major=$(echo "$LDD_VERSION" | cut -d. -f1)
        minor=$(echo "$LDD_VERSION" | cut -d. -f2)
        if [ "$major" -gt 2 ] || { [ "$major" -eq 2 ] && [ "$minor" -ge 28 ]; }; then
            pass_msg "glibc $LDD_VERSION ≥ 2.28 — 满足 PyInstaller 打包最低要求"
        else
            fail_msg "glibc $LDD_VERSION < 2.28 — PyInstaller 打包产物可能无法运行"
            ((failures++))
        fi
    fi
else
    warn_msg "未找到 ldd 命令"
fi

# 图形环境
if [ -n "$DISPLAY" ]; then
    pass_msg "图形环境: DISPLAY=$DISPLAY"
elif [ -n "$WAYLAND_DISPLAY" ]; then
    pass_msg "图形环境: Wayland=$WAYLAND_DISPLAY"
else
    warn_msg "未检测到图形环境（$DISPLAY / $WAYLAND_DISPLAY 均为空），GUI 程序可能无法启动"
    info_msg "如通过 SSH 远程测试，请使用 ssh -X 或设置 DISPLAY"
fi

# 桌面环境
DESKTOP="${XDG_CURRENT_DESKTOP:-${DESKTOP_SESSION:-未知}}"
echo "桌面环境: $DESKTOP"
case "$DESKTOP" in
    *UKUI*)  info_msg "麒麟 UKUI 桌面" ;;
    *Deepin*) info_msg "统信/Deepin 桌面" ;;
    *GNOME*|*KDE*|*XFCE*|*MATE*) : ;;
    *)       info_msg "桌面环境非主流，可能影响 UI 风格" ;;
esac

# 内存
mem_total=$(grep MemTotal /proc/meminfo 2>/dev/null | awk '{printf "%.1f GB", $2/1024/1024}')
echo "内存: ${mem_total:-未知}"
if [ -n "$mem_total" ]; then
    mem_gb=$(echo "$mem_total" | grep -oP '[\d.]+')
    if [ "$(echo "$mem_gb < 1" | bc 2>/dev/null || echo 0)" = "1" ]; then
        warn_msg "内存不足 1 GB，可能影响大文件处理"
    fi
fi

# ============================================================
section "2. 中文语言环境"
# ============================================================

if locale -a 2>/dev/null | grep -qi 'zh_CN'; then
    pass_msg "中文 locale 已安装"
else
    warn_msg "未找到 zh_CN locale，界面中文可能显示为乱码"
    info_msg "安装: sudo localedef -i zh_CN -f UTF-8 zh_CN.UTF-8"
fi
echo "当前 locale: $(locale 2>/dev/null | grep LANG= || echo 未设置)"

# ============================================================
section "3. 文件权限与磁盘空间"
# ============================================================

# 检查当前用户
echo "当前用户: $(whoami) (uid=$(id -u))"
if [ "$(id -u)" -eq 0 ]; then
    warn_msg "以 root 运行 GUI 程序可能导致权限问题"
fi

# 检查 HOME
if [ -d "$HOME" ] && [ -w "$HOME" ]; then
    pass_msg "HOME 目录可写: $HOME"
else
    fail_msg "HOME 目录不可写，程序无法保存配置"
    ((failures++))
fi

# 磁盘空间
df_line=$(df -h "$HOME" 2>/dev/null | tail -1)
echo "HOME 磁盘空间: $df_line"
avail_kb=$(df "$HOME" 2>/dev/null | tail -1 | awk '{print $4}')
if [ -n "$avail_kb" ] && [ "$avail_kb" -lt 512000 ]; then
    warn_msg "可用空间不足 500 MB"
fi

# ============================================================
section "4. PyInstaller 二进制诊断"
# ============================================================

if [ -n "$BINARY" ]; then
    # 用户指定了二进制路径
    BIN="$BINARY"
elif [ -f ./DocFormatPro ]; then
    BIN="./DocFormatPro"
elif [ -f dist/DocFormatPro ]; then
    BIN="dist/DocFormatPro"
else
    BIN=""
fi

if [ -n "$BIN" ]; then
    echo "检测二进制: $BIN"

    if [ -f "$BIN" ]; then
        pass_msg "文件存在"
        file_type=$(file "$BIN" 2>/dev/null)
        echo "  $file_type"

        if [ -x "$BIN" ]; then
            pass_msg "文件可执行"
        else
            fail_msg "文件无执行权限"
            info_msg "修复: chmod +x $BIN"
            ((failures++))
        fi

        echo ""
        info_msg "检查动态库依赖（ldd）..."

        # ldd 检测，区分缺失 so 和未找到
        missing_libs=0
        while IFS= read -r line; do
            if echo "$line" | grep -q "not found"; then
                fail_msg "缺失: $line"
                ((missing_libs++))
            fi
        done < <(ldd "$BIN" 2>&1 || true)

        if [ "$missing_libs" -eq 0 ]; then
            pass_msg "动态库依赖检测通过"
        else
            fail_msg "共缺失 $missing_libs 个动态库"
            info_msg "常见缺失库及安装方式："
            info_msg "  libQt5* → sudo apt install libqt5widgets5 libqt5gui5 libqt5core5a"
            info_msg "  libxcb*  → sudo apt install libxcb-util1 libxcb-icccm4"
            ((failures++))
        fi

        # 试跑（带超时，捕获崩溃）
        echo ""
        info_msg "试跑二进制（5 秒超时）..."
        timeout 5 "$BIN" 2>&1 || true
        EXIT_CODE=$?
        case $EXIT_CODE in
            0)   pass_msg "二进制启动成功（正常退出）" ;;
            124) pass_msg "二进制 5 秒内未退出 — 说明 GUI 已成功拉起" ;;
            139) fail_msg "段错误 (SIGSEGV) — 可能是 glibc 或其他 so 版本不兼容"
                 info_msg "请检查构建机 glibc 是否 ≤ 本机 glibc"
                 ((failures++)) ;;
            127) fail_msg "无法执行 — 可能缺少动态链接器 (ld-linux)"
                 ((failures++)) ;;
            *)   warn_msg "退出码 $EXIT_CODE — 请查看上方错误信息" ;;
        esac
    else
        fail_msg "二进制文件不存在: $BIN"
        info_msg "用法: bash diagnose.sh /path/to/DocFormatPro"
        ((failures++))
    fi
else
    warn_msg "未找到 PyInstaller 二进制文件（跳过）"
    info_msg "如需检测二进制，请: bash diagnose.sh /path/to/DocFormatPro"
fi

# ============================================================
section "5. 源码运行环境（Python 依赖）"
# ============================================================

# Python
if command -v python3 &>/dev/null; then
    PY=$(command -v python3)
    PY_VER=$("$PY" --version 2>&1)
    echo "Python: $PY_VER ($PY)"

    PY_MAJOR=$("$PY" -c 'import sys; print(sys.version_info.major)' 2>/dev/null || echo 0)
    PY_MINOR=$("$PY" -c 'import sys; print(sys.version_info.minor)' 2>/dev/null || echo 0)

    if [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -ge 7 ]; then
        pass_msg "Python $PY_MAJOR.$PY_MINOR ≥ 3.7 — 满足要求"
    elif [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -eq 6 ]; then
        warn_msg "Python 3.6 — 可能可行但未经测试，建议升级"
    else
        fail_msg "Python 版本过低，需要 ≥ 3.7"
        info_msg "麒麟 V10 默认 python3 可能是 3.7，若没有请: sudo apt install python3"
        ((failures++))
    fi

    # pip
    if command -v pip3 &>/dev/null; then
        pass_msg "pip3 可用"
    else
        warn_msg "pip3 未安装 — 无法在线安装依赖"
        info_msg "安装: sudo apt install python3-pip"
    fi

    # PyQt5
    echo ""
    info_msg "检测 Python 依赖包..."
    for mod in PyQt5 docx lxml; do
        if "$PY" -c "import $mod" 2>/dev/null; then
            pass_msg "$mod — 已安装"
        else
            fail_msg "$mod — 未安装！"
            case $mod in
                PyQt5) info_msg "  安装: pip3 install PyQt5  (或 apt install python3-pyqt5)" ;;
                docx)  info_msg "  安装: pip3 install python-docx" ;;
                lxml)  info_msg "  安装: pip3 install lxml  (或 apt install python3-lxml)" ;;
            esac
            ((failures++))
        fi
    done
else
    fail_msg "未找到 python3！"
    info_msg "安装: sudo apt install python3 python3-pip"
    ((failures++))
fi

# ============================================================
section "6. LibreOffice（.doc/.wps 格式转换依赖）"
# ============================================================

SOFFICE=""
for cand in soffice libreoffice /usr/bin/soffice /usr/bin/libreoffice; do
    if command -v "$cand" &>/dev/null; then
        SOFFICE="$cand"
        break
    fi
done

if [ -n "$SOFFICE" ]; then
    LO_VER=$("$SOFFICE" --version 2>/dev/null | head -1 || echo "版本未知")
    pass_msg "LibreOffice 可用: $LO_VER"
else
    warn_msg "LibreOffice 未安装 — 将无法处理 .doc / .wps 文件"
    info_msg "安装: sudo apt install libreoffice-writer"
    info_msg "替代方案: 用 WPS 打开文件另存为 .docx"
fi

# ============================================================
section "7. 网络（在线安装、获取帮助）"
# ============================================================

if command -v curl &>/dev/null || command -v wget &>/dev/null; then
    pass_msg "有下载工具 (curl/wget)"
else
    warn_msg "无 curl 或 wget，无法在线下载依赖"
fi

# ============================================================
section "诊断总结"
# ============================================================

echo ""
if [ "$failures" -eq 0 ]; then
    echo -e "  ${GREEN}${BOLD}所有检测通过 ✓${NC}"
    echo "  系统满足 DocFormat Pro 运行条件，可以正常使用。"
elif [ "$failures" -le 2 ]; then
    echo -e "  ${YELLOW}${BOLD}存在 $failures 个问题 ⚠${NC}"
    echo "  请根据上方提示修复后重试。"
else
    echo -e "  ${RED}${BOLD}存在 $failures 个问题 ✗${NC}"
    echo "  请将本诊断输出完整复制反馈给开发者（见下方）。"
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  请将以上全部输出复制反馈给开发者以便针对性修复"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
