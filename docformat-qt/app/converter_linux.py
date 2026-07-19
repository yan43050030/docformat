# -*- coding: utf-8 -*-
"""Linux（麒麟/UOS）.doc/.wps → .docx 转换：LibreOffice → WPS 回退链"""
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
    '/opt/libreoffice24.8/program/soffice',
]

WPS_CANDIDATES = [
    '/opt/kingsoft/wps-office/office6/wps',
    '/opt/kingsoft/wps-office/office6/wps_cli',
    'wps',
    '/usr/bin/wps',
]


def _find_binary(candidates):
    """从候选列表中查找第一个可用的可执行文件"""
    for cand in candidates:
        if os.sep in cand:
            # 绝对路径：检查文件是否存在且可执行
            if os.path.isfile(cand) and os.access(cand, os.X_OK):
                return cand
        else:
            # 短名称：用 which 查找
            found = shutil.which(cand)
            if found:
                return found
    return None


def find_soffice():
    return _find_binary(SOFFICE_CANDIDATES)


def find_wps():
    return _find_binary(WPS_CANDIDATES)


def _convert_via_libreoffice(soffice, input_path, out_dir):
    """用 LibreOffice 无头模式转换，成功返回输出路径，失败返回 None"""
    cmd = [
        soffice, '--headless', '--norestore',
        '--convert-to', 'docx', '--outdir', out_dir, input_path,
    ]
    try:
        proc = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120,
        )
    except subprocess.TimeoutExpired:
        return None  # 超时 → 回退

    stem = os.path.splitext(os.path.basename(input_path))[0]
    produced = os.path.join(out_dir, stem + '.docx')
    if proc.returncode == 0 and os.path.exists(produced):
        return produced
    return None


def _convert_via_wps(wps_bin, input_path, out_dir):
    """用 WPS 尝试转换 .doc → .docx。

    WPS Linux 无公开 headless API，尝试几种已知可行的方法：
      1) wps_cli --convert-to（某些 Pro 版本支持）
      2) wps -w <file>（某些版本支持 Writer 模式）
      3) wps <file>（最后尝试，依赖 DISPLAY）

    成功返回输出路径，失败返回 None。
    """
    stem = os.path.splitext(os.path.basename(input_path))[0]
    produced = os.path.join(out_dir, stem + '.docx')
    abs_input = os.path.abspath(input_path)

    # 方法 1：wps_cli（企业版 Pro 的命令行工具）
    if wps_bin.endswith('wps_cli') or os.path.basename(wps_bin) == 'wps_cli':
        try:
            subprocess.run(
                [wps_bin, '--convert-to', 'docx', '--outdir', out_dir, abs_input],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120,
            )
            if os.path.exists(produced):
                return produced
        except Exception:
            pass

    # 方法 2：wps 二进制，尝试 --writer 模式
    wps_writer = wps_bin.replace('_cli', '') if '_cli' in wps_bin else wps_bin
    try:
        subprocess.run(
            [wps_writer, '-w', abs_input],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60,
        )
        if os.path.exists(produced):
            return produced
    except Exception:
        pass

    # 方法 3：wps 直接打开（需要 DISPLAY，几乎不可靠但试试）
    # 检查同一目录下的 wps 二进制
    wps_main = os.path.join(os.path.dirname(wps_writer), 'wps') if os.path.dirname(wps_writer) else wps_writer
    if os.path.isfile(wps_main) and os.access(wps_main, os.X_OK):
        try:
            env = os.environ.copy()
            # 确保有 DISPLAY，没有则设一个假的（WPS 可能依然拒绝）
            if not env.get('DISPLAY'):
                env['DISPLAY'] = ':0'
            subprocess.run(
                [wps_main, abs_input],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=30, env=env,
            )
            if os.path.exists(produced):
                return produced
        except Exception:
            pass

    return None


def convert_to_docx(input_path, output_path=None):
    """把 .doc/.wps 转成 .docx，返回输出路径。

    回退链：LibreOffice → WPS → 报错（含安装指引）。
    """
    out_dir = tempfile.mkdtemp(prefix='docformat_')

    # 第一优先：LibreOffice（headless 模式最可靠）
    soffice = find_soffice()
    if soffice:
        result = _convert_via_libreoffice(soffice, input_path, out_dir)
        if result:
            if output_path:
                shutil.move(result, output_path)
                return output_path
            return result

    # 第二优先：WPS Office
    wps = find_wps()
    if wps:
        result = _convert_via_wps(wps, input_path, out_dir)
        if result:
            if output_path:
                shutil.move(result, output_path)
                return output_path
            return result
        # WPS 找到了但转换失败 → 不是真的报错，给明确指引
        raise RuntimeError(
            "WPS Office 已安装但转换失败（WPS Linux 不保证支持命令行转换）。\n"
            "解决方案（任选其一）：\n"
            "  1. 安装 LibreOffice 获得可靠转换：sudo apt install libreoffice-writer\n"
            "  2. 用 WPS 手动打开 {}，另存为 .docx 后再处理".format(
                os.path.basename(input_path))
        )

    # 两个都没有
    raise RuntimeError(
        "未找到 LibreOffice 或 WPS，无法转换 .doc/.wps 文件。\n"
        "解决方案（任选其一）：\n"
        "  1. 安装 LibreOffice：sudo apt install libreoffice-writer\n"
        "  2. 用 WPS 打开 {}，另存为 .docx 后再处理".format(
            os.path.basename(input_path))
    )
