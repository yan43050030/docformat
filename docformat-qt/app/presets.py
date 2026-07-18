# -*- coding: utf-8 -*-
"""模板（预设）管理：内置预设来自引擎 PRESETS，用户模板持久化到配置目录"""
import copy
import json
import os
import sys
import time
from pathlib import Path

from scripts.formatter import PRESETS as BUILTIN_PRESETS

BUILTIN_ORDER = ['official_gbk', 'official', 'academic', 'legal']


def config_dir():
    if sys.platform == 'win32':
        base = os.environ.get('APPDATA') or str(Path.home() / 'AppData' / 'Roaming')
    elif sys.platform == 'darwin':
        base = str(Path.home() / 'Library' / 'Application Support')
    else:
        base = os.environ.get('XDG_CONFIG_HOME') or str(Path.home() / '.config')
    d = Path(base) / 'DocFormatPro'
    d.mkdir(parents=True, exist_ok=True)
    return d


def templates_path():
    return config_dir() / 'templates.json'


# 用户模板在引擎预设基础上额外携带的开关（引擎均会读取）
EXTRA_DEFAULTS = {
    'first_line_bold': False,
    'bold_serial': True,
    'deep_clean': False,
}


class PresetManager(object):
    """key 规则：内置用引擎键（official/academic/legal），用户模板用 user_<ts>"""

    def __init__(self):
        self.user = {}       # key -> preset dict
        self.active_key = 'official_gbk'   # 默认使用图解标准版
        self.load()

    # ---------- 持久化 ----------
    def load(self):
        p = templates_path()
        if p.exists():
            try:
                with open(str(p), 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.user = data.get('user', {}) or {}
                last = data.get('last_used') or {}
                key = last.get('key')
                if key and (key in self.user or key in BUILTIN_PRESETS):
                    self.active_key = key
            except Exception:
                self.user = {}

    def save(self):
        data = {
            'user': self.user,
            'last_used': {'key': self.active_key},
        }
        p = templates_path()
        tmp = str(p) + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, str(p))

    # ---------- 查询 ----------
    def list_all(self):
        """返回 [(key, name, is_builtin)]"""
        items = []
        for k in BUILTIN_ORDER:
            if k in BUILTIN_PRESETS:
                items.append((k, BUILTIN_PRESETS[k].get('name', k), True))
        for k in sorted(self.user.keys()):
            items.append((k, self.user[k].get('name', k), False))
        return items

    def is_builtin(self, key):
        return key in BUILTIN_PRESETS

    def get(self, key):
        if key in BUILTIN_PRESETS:
            return copy.deepcopy(BUILTIN_PRESETS[key])
        preset = copy.deepcopy(self.user.get(key, BUILTIN_PRESETS['official']))
        # 旧版本用户模板缺少后加入的元素节点时，用公文默认值补齐
        for el_key in ('security', 'docnum'):
            if el_key not in preset and el_key in BUILTIN_PRESETS['official']:
                preset[el_key] = copy.deepcopy(BUILTIN_PRESETS['official'][el_key])
        return preset

    def set_active(self, key):
        self.active_key = key
        self.save()

    # ---------- 修改 ----------
    def create(self, name, base_key='official'):
        preset = self.get(base_key)
        preset['name'] = name
        for k, v in EXTRA_DEFAULTS.items():
            preset.setdefault(k, v)
        key = 'user_{}'.format(int(time.time() * 1000))
        self.user[key] = preset
        self.active_key = key
        self.save()
        return key

    def duplicate(self, key):
        src = self.get(key)
        src['name'] = src.get('name', key) + ' (副本)'
        for k, v in EXTRA_DEFAULTS.items():
            src.setdefault(k, v)
        new_key = 'user_{}'.format(int(time.time() * 1000))
        self.user[new_key] = src
        self.active_key = new_key
        self.save()
        return new_key

    def delete(self, key):
        if key in self.user:
            del self.user[key]
            if self.active_key == key:
                self.active_key = 'official_gbk'
            self.save()

    def update(self, key, preset):
        if key in self.user:
            self.user[key] = preset
            self.save()

    def rename(self, key, name):
        if key in self.user:
            self.user[key]['name'] = name
            self.save()

    # ---------- 导入导出 ----------
    def export_to(self, key, path):
        preset = self.get(key)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(preset, f, ensure_ascii=False, indent=2)

    def import_from(self, path):
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        imported = []
        entries = data if isinstance(data, list) else [data]
        for entry in entries:
            if not isinstance(entry, dict) or 'page' not in entry:
                continue
            for k, v in EXTRA_DEFAULTS.items():
                entry.setdefault(k, v)
            entry.setdefault('name', '导入模板')
            key = 'user_{}'.format(int(time.time() * 1000) + len(imported))
            self.user[key] = entry
            imported.append(key)
        if imported:
            self.active_key = imported[0]
            self.save()
        return imported

    # ---------- 处理参数 ----------
    def engine_args(self, key):
        """返回 (preset_name, custom_settings) 供 format_document 使用"""
        if self.is_builtin(key):
            return key, None
        return 'custom', self.get(key)
