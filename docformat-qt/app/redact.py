# -*- coding: utf-8 -*-
"""日志/诊断脱敏：涉密场景下文件名、目录、用户名不明文落盘。

策略：
  - 文档文件名 → 保留扩展名，主体替换为不可逆短哈希（文档-a1b2.docx），
    同名文件哈希一致，方便对照但不泄露真实名。
  - 完整路径 → 丢弃所有目录层级（可能含人名/涉密目录名），只留脱敏后的文件名。
  - 用户主目录（C:\\Users\\张三、/home/张三）→ 用户名段替换为 ***。
"""
import hashlib
import re

_DOC_EXT = r'docx|doc|wps|txt|md|markdown'

# 可选带盘符/根、任意目录层级 + 文档扩展名的文件名（含完整路径）
_FILE_RE = re.compile(
    r'(?:[A-Za-z]:\\|/)?(?:[^\s\\/]+[\\/])*([^\s\\/]+)\.(' + _DOC_EXT + r')',
    re.IGNORECASE)

# 用户主目录前缀后的用户名段
_HOME_RE = re.compile(
    r'([A-Za-z]:\\Users\\|/home/|/Users/|/root)([^\s\\/]*)',
    re.IGNORECASE)


def _hash_stem(stem):
    return hashlib.md5(stem.encode('utf-8', 'replace')).hexdigest()[:4]


def mask_filename(name):
    """测试.docx → 文档-a1b2.docx（保留扩展名）"""
    m = re.match(r'^(.*)\.(' + _DOC_EXT + r')$', name, re.IGNORECASE)
    if not m:
        return name
    return '文档-{}.{}'.format(_hash_stem(m.group(1)), m.group(2))


def mask_home(path):
    """C:\\Users\\张三\\... → C:\\Users\\***\\...（仅用户名段脱敏，保留结构）"""
    return _HOME_RE.sub(lambda m: m.group(1) + '***', path or '')


def redact_text(text):
    """对一行日志/诊断文本做脱敏：文件名哈希化，路径去目录只留脱敏文件名。"""
    if not text:
        return text

    def _repl(m):
        return '文档-{}.{}'.format(_hash_stem(m.group(1)), m.group(2))

    # 文档路径/文件名（含目录）→ 脱敏文件名（目录整体丢弃）
    out = _FILE_RE.sub(_repl, text)
    # 残留的用户主目录（非文档路径场景，如临时目录、错误信息）
    out = mask_home(out)
    return out
