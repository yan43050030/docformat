#!/usr/bin/env bash
# Linux 二进制打包（须在目标架构机器上执行：x86_64 在 x86_64 上打，飞腾/鲲鹏 arm64 在 arm64 上打）
# 注意：PyInstaller 产物受构建机 glibc 版本约束，建议在最老的目标系统上构建
set -e
cd "$(dirname "$0")/.."

pip3 install --user pyinstaller

python3 -m PyInstaller --noconfirm --clean --onefile --windowed \
    --name DocFormatPro \
    --collect-data docx \
    --add-data "assets:assets" \
    --add-data "templates:templates" \
    --hidden-import scripts.formatter \
    --hidden-import scripts.punctuation \
    --hidden-import scripts.analyzer \
    main.py

echo ""
echo "打包完成: dist/DocFormatPro  ($(uname -m))"
echo "分发时附带 packaging/docformat.desktop 并将 Exec 改为二进制路径"
