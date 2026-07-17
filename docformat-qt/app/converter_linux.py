# -*- coding: utf-8 -*-
"""Linux（麒麟/UOS）.doc/.wps → .docx 转换：LibreOffice 命令行回退链"""
import os
import shutil
import subprocess
import tempfile

SOFFICE_CANDIDATES = [
    'soffice',
    'libreoffice',
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/opt/libreoffice/program/soffice',
]


def find_soffice():
    for cand in SOFFICE_CANDIDATES:
        path = shutil.which(cand) if os.sep not in cand else (cand if os.path.exists(cand) else None)
        if path:
            return path
    return None


def convert_to_docx(input_path, output_path=None):
    """用 LibreOffice 无头模式把 .doc/.wps 转成 .docx，返回输出路径。

    找不到 LibreOffice 时抛 RuntimeError，附带可操作的解决建议。
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "未找到 LibreOffice，无法转换 .doc/.wps 文件。\n"
            "解决方案（任选其一）：\n"
            "  1. 安装 LibreOffice：sudo apt install libreoffice-writer\n"
            "  2. 用 WPS 打开该文件，另存为 .docx 后再处理"
        )

    out_dir = tempfile.mkdtemp(prefix='docformat_')
    cmd = [
        soffice, '--headless', '--norestore',
        '--convert-to', 'docx', '--outdir', out_dir, input_path,
    ]
    try:
        proc = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice 转换超时（120 秒）：{}".format(os.path.basename(input_path)))

    stem = os.path.splitext(os.path.basename(input_path))[0]
    produced = os.path.join(out_dir, stem + '.docx')
    if proc.returncode != 0 or not os.path.exists(produced):
        err = proc.stderr.decode('utf-8', errors='replace').strip()
        raise RuntimeError("LibreOffice 转换失败：{}".format(err or '未知错误'))

    if output_path:
        shutil.move(produced, output_path)
        return output_path
    return produced
