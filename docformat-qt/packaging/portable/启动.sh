#!/bin/sh
# DocFormat Pro 绿色版启动脚本
#
# 直接双击 DocFormatPro 通常也能跑，这个脚本是为了兜住普通用户环境里
# 常见的两件事：
#   1) 从压缩包解出来后可执行位丢了（用图形归档工具解压、或经 Windows
#      中转过一道，都可能把 755 抹成 644）——这里补回来；
#   2) /tmp 挂了 noexec（不少信创机器的安全基线这么配）。本程序是单文件
#      打包的，运行时要先把自己解到临时目录再执行，/tmp 不让执行就起不来。
#      改用家目录下的缓存目录，普通用户自己就有写权限，不需要管理员。
set -e
DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
APP="$DIR/DocFormatPro"

if [ ! -f "$APP" ]; then
    echo "找不到主程序：$APP" >&2
    echo "请确认压缩包已完整解开，启动.sh 与 DocFormatPro 在同一个目录里。" >&2
    exit 1
fi
[ -x "$APP" ] || chmod u+rx "$APP" 2>/dev/null || true
if [ ! -x "$APP" ]; then
    echo "主程序没有执行权限，且当前用户改不动它。" >&2
    echo "请把整个目录复制到你自己有写权限的地方（如家目录）再运行。" >&2
    exit 1
fi

# 自带一个可执行的临时目录，绕开 /tmp noexec；顺带把解包内容集中在一处，
# 卸载时删掉这个目录即可，不留散落文件。
CACHE="${XDG_CACHE_HOME:-$HOME/.cache}/docformat-pro/run"
if mkdir -p "$CACHE" 2>/dev/null; then
    printf '#!/bin/sh\nexit 0\n' > "$CACHE/.exectest" 2>/dev/null || true
    chmod u+x "$CACHE/.exectest" 2>/dev/null || true
    if [ -x "$CACHE/.exectest" ] && "$CACHE/.exectest" 2>/dev/null; then
        TMPDIR="$CACHE"
        export TMPDIR
    fi
    rm -f "$CACHE/.exectest" 2>/dev/null || true
fi

exec "$APP" "$@"
