# -*- coding: utf-8 -*-
"""持久化日志文件：默认脱敏，带大小轮转，可随时清空。

默认行为（涉密友好）：
  - 文件名哈希化、路径去目录、用户名脱敏后才写入。
  - 用户可在设置里关闭脱敏（仅本机调试，带风险提示），或彻底关闭文件日志。
  - 单文件上限 512KB，超出后转存 .1（只保留一份历史）。
"""
import os
import time

from app.theme import settings
from app.redact import redact_text

_MAX_BYTES = 512 * 1024


def log_dir():
    from app.presets import config_dir
    d = config_dir() / 'logs'
    try:
        d.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return d


def log_path():
    return log_dir() / 'docformat.log'


def persist_enabled():
    return settings().value('log/persist', True, type=bool)


def redact_enabled():
    return settings().value('log/redact', True, type=bool)


def set_persist(on):
    settings().setValue('log/persist', bool(on))


def set_redact(on):
    settings().setValue('log/redact', bool(on))


def _rotate_if_needed(path):
    try:
        if path.exists() and path.stat().st_size > _MAX_BYTES:
            bak = str(path) + '.1'
            if os.path.exists(bak):
                os.remove(bak)
            os.replace(str(path), bak)
    except Exception:
        pass


def write(level, message):
    """写一条日志到文件（若启用）。脱敏默认开启。异常静默，绝不影响主流程。"""
    if not persist_enabled():
        return
    try:
        msg = message.replace('\n', ' ')
        if redact_enabled():
            msg = redact_text(msg)
        p = log_path()
        _rotate_if_needed(p)
        line = '[{}] [{}] {}\n'.format(time.strftime('%Y-%m-%d %H:%M:%S'), level, msg)
        with open(str(p), 'a', encoding='utf-8') as f:
            f.write(line)
    except Exception:
        pass


def clear():
    """清空日志文件（含历史 .1）。"""
    try:
        p = log_path()
        for target in (str(p), str(p) + '.1'):
            if os.path.exists(target):
                os.remove(target)
        return True
    except Exception:
        return False
