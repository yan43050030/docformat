# -*- coding: utf-8 -*-
"""
数据模型与预设定义 — 从 formatter.py 拆分

包含：PRESETS 字典、字体回退映射、自定义预设加载、平台适配
"""

import json
import logging
import os
import re
import sys
from copy import deepcopy
from pathlib import Path

logger = logging.getLogger('docformat.data_model')

# ===== macOS 字体回退映射 =====
MACOS_FONT_FALLBACK = {
    '仿宋_GB2312': 'STFangsong',
    '仿宋': 'STFangsong',
    '黑体': 'STHeiti',
    '楷体_GB2312': 'STKaiti',
    '楷体': 'STKaiti',
    '宋体': 'STSong',
    '方正小标宋简体': 'STSong',
    '方正小标宋_GBK': 'STSong',
    '方正仿宋_GBK': 'STFangsong',
    '方正黑体_GBK': 'STHeiti',
    '方正楷体_GBK': 'STKaiti',
    '华文仿宋': 'STFangsong',
    '华文中宋': 'STZhongsong',
}

MACOS_FONT_ALIASES = {
    '仿宋_GB2312': ['仿宋_GB2312', '仿宋_GB32312', '仿宋', 'FangSong_GB2312', 'FangSong'],
    '楷体_GB2312': ['楷体_GB2312', '楷体_GB32312', '楷体', 'KaiTi_GB2312', 'KaiTi'],
    '方正仿宋_GBK': ['方正仿宋_GBK', '仿宋_GB2312', '仿宋_GB32312', '仿宋', 'FangSong_GB2312', 'FangSong'],
}

_macos_installed_fonts = None
_macos_font_detection_done = False


def _get_macos_installed_fonts():
    """获取 macOS 上已安装的字体族名集合（结果会缓存）"""
    global _macos_installed_fonts, _macos_font_detection_done
    if _macos_font_detection_done:
        return _macos_installed_fonts

    _macos_font_detection_done = True

    if sys.platform != 'darwin':
        _macos_installed_fonts = set()
        return _macos_installed_fonts

    import subprocess

    try:
        result = subprocess.run(
            ['/usr/bin/python3', '-c',
             'from AppKit import NSFontManager;'
             'fm=NSFontManager.sharedFontManager();'
             'print("\\n".join(fm.availableFontFamilies()))'],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode == 0 and result.stdout.strip():
            fonts = {name.strip() for name in result.stdout.strip().split('\n') if name.strip()}
            if len(fonts) > 10:
                _macos_installed_fonts = fonts
                logger.info(f"macOS 字体检测成功（AppKit），共 {len(fonts)} 个字体族")
                return _macos_installed_fonts
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as e:
        logger.debug(f"AppKit 字体检测失败: {e}")

    try:
        result = subprocess.run(
            ['fc-list', '--format=%{family}\n'],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and result.stdout.strip():
            fonts = set()
            for line in result.stdout.strip().split('\n'):
                for name in line.split(','):
                    name = name.strip()
                    if name:
                        fonts.add(name)
            if len(fonts) > 10:
                _macos_installed_fonts = fonts
                logger.info(f"macOS 字体检测成功（fc-list），共 {len(fonts)} 个字体族")
                return _macos_installed_fonts
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as e:
        logger.debug(f"fc-list 字体检测失败: {e}")

    logger.warning("macOS 字体检测失败，将保持原字体名称不替换")
    _macos_installed_fonts = None
    return _macos_installed_fonts


def _resolve_font_for_macos(font_name):
    """解析单个字体名：优先保留 Windows 原字体名，仅在确认未安装时回退"""
    if font_name not in MACOS_FONT_FALLBACK:
        return font_name

    installed = _get_macos_installed_fonts()
    if installed is None:
        logger.debug(f"字体检测不可用，保留原字体名 '{font_name}'")
        return font_name

    if font_name in installed:
        logger.debug(f"字体 '{font_name}' 已安装，直接使用")
        return font_name

    for alias in MACOS_FONT_ALIASES.get(font_name, []):
        if alias in installed:
            logger.info(f"字体 '{font_name}' 未安装，使用兼容字体 '{alias}'")
            return alias

    fallback = MACOS_FONT_FALLBACK[font_name]
    logger.info(f"字体 '{font_name}' 未安装，回退到 '{fallback}'")
    return fallback


def get_missing_cn_fonts(preset):
    """Linux（麒麟/UOS）：用 fc-list 检查预设中文字体是否安装，返回缺失字体列表。"""
    if sys.platform in ('win32', 'darwin'):
        return []
    needed = set()
    for key, fmt in preset.items():
        if isinstance(fmt, dict) and fmt.get('font_cn'):
            needed.add(fmt['font_cn'])
    if preset.get('page_number_font'):
        needed.add(preset['page_number_font'])
    if not needed:
        return []
    import subprocess
    try:
        result = subprocess.run(['fc-list', '--format=%{family}\n'],
                                capture_output=True, text=True, timeout=10)
        if result.returncode != 0 or not result.stdout.strip():
            return []
        installed = set()
        for line in result.stdout.strip().split('\n'):
            for name in line.split(','):
                name = name.strip()
                if name:
                    installed.add(name)
        if len(installed) < 5:
            return []
        return sorted(needed - installed)
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        return []


def _adapt_fonts_for_platform(preset):
    """在 macOS 上适配字体：优先使用 Windows 原字体，确认未安装时才回退到系统字体"""
    if sys.platform != 'darwin':
        return preset
    import copy
    preset = copy.deepcopy(preset)
    for key, fmt in preset.items():
        if isinstance(fmt, dict) and 'font_cn' in fmt:
            fmt['font_cn'] = _resolve_font_for_macos(fmt['font_cn'])
    if 'page_number_font' in preset:
        preset['page_number_font'] = _resolve_font_for_macos(preset['page_number_font'])
    return preset


# ===== 自定义预设加载 =====

def load_custom_preset():
    """加载自定义预设"""
    if sys.platform == 'darwin':
        custom_config_file = Path.home() / 'Library' / 'Application Support' / 'DocFormatter' / "custom_settings.json"
    elif sys.platform == 'win32':
        base = os.environ.get('APPDATA') or str(Path.home() / 'AppData' / 'Roaming')
        custom_config_file = Path(base) / 'DocFormatter' / "custom_settings.json"
    else:
        base = os.environ.get('XDG_CONFIG_HOME') or str(Path.home() / '.config')
        custom_config_file = Path(base) / 'DocFormatter' / "custom_settings.json"

    legacy_config_file = Path(__file__).parent.parent / "custom_settings.json"
    if not custom_config_file.exists() and legacy_config_file.exists():
        custom_config_file = legacy_config_file

    if custom_config_file.exists():
        try:
            with open(custom_config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get('schema_version') == 2:
                presets = data.get('presets', [])
                active_id = data.get('active_preset_id')
                if active_id:
                    for preset in presets:
                        if preset.get('id') == active_id:
                            return preset
                return presets[0] if presets else None
            return data
        except Exception:
            return None
    return None


def _merge_preset_settings(base, overrides):
    """递归合并预设配置"""
    merged = deepcopy(base)
    for key, value in (overrides or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_preset_settings(merged[key], value)
        else:
            merged[key] = deepcopy(value)
    return merged


# ===== 内置预设 =====

PRESETS = {
    'official_gbk': {
        'name': '公文格式（图解版·22行28字）',
        'deep_clean': False,
        'page_number': True,
        'page_number_font': '宋体',
        'page_number_size': 14,
        'page_number_style': 'dash',
        'page_number_position': 'center',
        'page_number_offset_mm': 7,
        'replace_existing_page_number': True,
        'page': {'top': 3.8, 'bottom': 3.3, 'left': 2.8, 'right': 2.8},
        'page_size': 'A4',
        'grid': {'lines_per_page': 22, 'chars_per_line': 28},
        'gb_signature_layout': True,
        'security': {
            'font_cn': '方正黑体_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 0,
            'space_before': 0, 'space_after': 28,
        },
        'docnum': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'center', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'title': {
            'font_cn': '方正小标宋_GBK', 'font_en': 'Times New Roman',
            'size': 22, 'bold': True, 'align': 'center', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'recipient': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 0,
        },
        'heading1': {
            'font_cn': '方正黑体_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading2': {
            'font_cn': '方正楷体_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading3': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading4': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'body': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'justify', 'indent': 32,
            'line_spacing': 28,
        },
        'signature': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'right', 'indent': 0,
        },
        'date': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'right', 'indent': 0,
        },
        'attachment': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'justify', 'indent': 0,
        },
        'roster': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'closing': {
            'font_cn': '方正仿宋_GBK', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_after': 28,
        },
    },
    'official': {
        'name': '公文格式',
        'deep_clean': False,
        'page_size': 'A4',
        'page_number': True,
        'page_number_font': '宋体',
        'page_number_size': 14,
        'page_number_style': 'dash',
        'page_number_position': 'outside',
        'page_number_offset_mm': 7,
        'replace_existing_page_number': True,
        'page': {'top': 3.7, 'bottom': 3.5, 'left': 2.8, 'right': 2.6},
        'security': {
            'font_cn': '黑体', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'docnum': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'center', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'title': {
            'font_cn': '方正小标宋简体', 'font_en': 'Times New Roman',
            'size': 22, 'bold': False, 'align': 'center', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'recipient': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 0,
        },
        'heading1': {
            'font_cn': '黑体', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading2': {
            'font_cn': '楷体_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading3': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': True, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'heading4': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 32,
            'space_before': 0, 'space_after': 0,
        },
        'body': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'justify', 'indent': 32,
            'line_spacing': 28,
        },
        'signature': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'right', 'indent': 0,
        },
        'date': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'right', 'indent': 0,
        },
        'attachment': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'justify', 'indent': 0,
        },
        'roster': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 0,
            'space_before': 0, 'space_after': 0,
        },
        'closing': {
            'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman',
            'size': 16, 'bold': False, 'align': 'left', 'indent': 32,
        },
    },
    'academic': {
        'name': '学术论文格式',
        'deep_clean': False,
        'page': {'top': 2.5, 'bottom': 2.5, 'left': 2.5, 'right': 2.5},
        'security': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 14, 'bold': True, 'align': 'left', 'indent': 0},
        'docnum': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'center', 'indent': 0},
        'title': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 18, 'bold': True, 'align': 'center', 'indent': 0},
        'recipient': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'left', 'indent': 0},
        'heading1': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 15, 'bold': True, 'align': 'left', 'indent': 0},
        'heading2': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 14, 'bold': True, 'align': 'left', 'indent': 0},
        'heading3': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'left', 'indent': 0},
        'heading4': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'left', 'indent': 0},
        'body': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'justify', 'indent': 24, 'line_spacing': None},
        'signature': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'right', 'indent': 0},
        'date': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'right', 'indent': 0},
        'attachment': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'justify', 'indent': 0},
        'roster': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'left', 'indent': 0},
        'closing': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 12, 'bold': False, 'align': 'left', 'indent': 24},
    },
    'legal': {
        'name': '法律文书格式',
        'deep_clean': False,
        'page': {'top': 3.0, 'bottom': 2.5, 'left': 3.0, 'right': 2.5},
        'security': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 16, 'bold': False, 'align': 'left', 'indent': 0},
        'docnum': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'center', 'indent': 0},
        'title': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 22, 'bold': True, 'align': 'center', 'indent': 0},
        'recipient': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'heading1': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'heading2': {'font_cn': '黑体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'heading3': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'heading4': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'body': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'justify', 'indent': 28, 'line_spacing': None},
        'signature': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'right', 'indent': 0},
        'date': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'right', 'indent': 0},
        'attachment': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'justify', 'indent': 0},
        'roster': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 0},
        'closing': {'font_cn': '宋体', 'font_en': 'Times New Roman', 'size': 14, 'bold': False, 'align': 'left', 'indent': 28},
    },
}
