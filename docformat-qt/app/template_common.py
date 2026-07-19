# -*- coding: utf-8 -*-
"""
模板起草 & 模板制作 公共模块

提供：
  - 共享模板目录 TEMPLATE_DIR + 多目录管理
  - 模板解析函数（两个模块共用同一套 .md 模板格式约定）
  - 本文档定义了模块一（template_draft_page）和模块二（template_maker_page）
    之间的对接契约，不要随意修改模板格式。
"""
import json
import os
import re
import sys

from app.presets import config_dir

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")

# 注释语法：
#   // ...         → 整行注释（以 // 开头），生成时完全忽略
#   【...】        → 行内注释，生成时忽略方括号及其内容
COMMENT_LINE_RE = re.compile(r"^\s*//")
INLINE_COMMENT_RE = re.compile(r"【[^】]*】")

TEMPLATE_DIR = os.path.join(config_dir(), "templates")


def bundled_templates_dir():
    """软件自带的模板目录（PyInstaller 打包后从临时目录，开发时从项目目录）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller --onefile：sys._MEIPASS 为临时解压目录
        # PyInstaller --onedir：sys._MEIPASS 为 exe 所在目录
        return os.path.join(sys._MEIPASS, "templates")
    # 开发环境：app/template_common.py → 上一级是 app，再上一级是项目根
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates")


def is_bundled_dir(path):
    """判断是否为打包自带的模板目录（不可移除）"""
    return os.path.normpath(path) == os.path.normpath(bundled_templates_dir())


def _template_dirs_config_path():
    """多模板目录列表的配置文件路径"""
    return os.path.join(config_dir(), "template_dirs.json")


def load_template_dirs():
    """返回所有模板目录列表：自带目录 + 用户保存的目录 + 兜底默认目录"""
    dirs = []
    # 软件自带模板目录（优先、只读）
    bundled = bundled_templates_dir()
    if os.path.isdir(bundled):
        dirs.append(bundled)
    # 用户保存的目录
    cfg = _template_dirs_config_path()
    if os.path.exists(cfg):
        try:
            with open(cfg, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for d in saved:
                d = os.path.expanduser(d)
                if d not in dirs:
                    dirs.append(d)
        except Exception:
            pass
    # 兜底：用户配置目录下的模板文件夹
    if TEMPLATE_DIR not in dirs:
        dirs.append(TEMPLATE_DIR)
    return dirs


def save_template_dirs(dirs):
    """保存用户添加的模板目录（自带目录自动排除，不持久化）"""
    cfg = _template_dirs_config_path()
    bundled = os.path.normpath(bundled_templates_dir())
    seen = []
    uniq = []
    for d in dirs:
        d = os.path.expanduser(d)
        d = os.path.normpath(d)
        if d == bundled:
            continue  # 自带目录无需保存
        if d not in seen:
            seen.append(d)
            uniq.append(d)
    with open(cfg, "w", encoding="utf-8") as f:
        json.dump(uniq, f, ensure_ascii=False, indent=2)


def scan_templates(dirs=None):
    """扫描所有模板目录，返回 [(显示名, 文件路径, 来源目录), ...]"""
    if dirs is None:
        dirs = load_template_dirs()
    results = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if fn.endswith(".md"):
                full = os.path.join(d, fn)
                results.append((fn[:-3], full, d))
    return results


def strip_comments(text):
    """移除注释：// 行注释整行删除，【...】行内注释删除"""
    lines = text.splitlines()
    cleaned = []
    for line in lines:
        if COMMENT_LINE_RE.match(line):
            continue
        line = INLINE_COMMENT_RE.sub("", line)
        cleaned.append(line)
    return "\n".join(cleaned)


def extract_fields(template_text):
    """按出现顺序提取去重后的占位符字段名（忽略注释行）"""
    clean = strip_comments(template_text)
    seen = []
    for m in PLACEHOLDER_RE.finditer(clean):
        name = m.group(1).strip()
        if name not in seen:
            seen.append(name)
    return seen


def parse_template(template_text):
    """解析模板 → {'title', 'body':[{'type','text'}], 'meta':{}}

    模板格式约定（勿改，两个模块靠此对接）：
      - // 开头 → 注释行，生成时忽略
      - 【...】 → 行内注释，生成时忽略
      - 标题: 以 "标题:" 开头的行
      - "一、" 开头 → 一级标题 (h1)
      - "（一）" 开头 → 二级标题 (h2)
      - 其余 → 正文 (body)
      - ---META--- 之后 → 附加字段（落款单位/日期等）
    """
    cleaned = strip_comments(template_text)
    body_part, meta = cleaned, {}
    if "---META---" in cleaned:
        body_part, meta_part = cleaned.split("---META---", 1)
        for line in meta_part.strip().splitlines():
            if ":" in line or "：" in line:
                k, v = re.split(r"[:：]", line, 1)
                meta[k.strip()] = v.strip()

    title, body = "", []
    for line in body_part.strip().splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith("标题:") or s.startswith("标题："):
            title = re.split(r"[:：]", s, 1)[1].strip()
        elif re.match(r"^[一二三四五六七八九十]+、", s):
            body.append({"type": "h1", "text": s})
        elif re.match(r"^（[一二三四五六七八九十]+）", s):
            body.append({"type": "h2", "text": s})
        else:
            body.append({"type": "body", "text": s})
    return {"title": title, "body": body, "meta": meta}


def render(parsed, values):
    """替换占位符；未填的保留 {{字段}} 原样，便于漏项检查"""
    def sub(text):
        return PLACEHOLDER_RE.sub(
            lambda m: values.get(m.group(1).strip()) or "{{" + m.group(1).strip() + "}}",
            text)
    return {
        "title": sub(parsed["title"]),
        "body": [{"type": b["type"], "text": sub(b["text"])} for b in parsed["body"]],
        "meta": {k: sub(v) for k, v in parsed["meta"].items()},
    }


def find_unfilled(rendered):
    remaining = set(PLACEHOLDER_RE.findall(rendered["title"]))
    for b in rendered["body"]:
        remaining |= set(PLACEHOLDER_RE.findall(b["text"]))
    return sorted(x.strip() for x in remaining)


# ================================================================
# 运行依赖检测
# ================================================================
def check_dependencies():
    """检测核心依赖包是否安装，返回 [(状态, 包名, 详情), ...]。状态: 'ok'/'warn'"""
    results = []
    deps = [
        ("PyQt5", "GUI 框架，软件运行必需"),
        ("docx", "python-docx，读写 Word 文档"),
        ("lxml", "XML 解析，python-docx 依赖"),
    ]
    if sys.platform == "win32":
        deps.append(("win32com", "pywin32，Windows .doc/.wps 转换"))
    for mod, desc in deps:
        try:
            __import__(mod)
            results.append(("ok", mod, desc + " ✓"))
        except ImportError:
            results.append(("warn", mod, desc + " ✗ 未安装"))
    return results


# ================================================================
# 快捷插入管理
# ================================================================
QUICK_INSERT_PATH = os.path.join(config_dir(), "quick_insert.json")

DEFAULT_QUICK_INSERTS = [
    {"label": "姓名", "text": "{{嫌疑人姓名}}"},
    {"label": "身份证号", "text": "{{身份证号}}"},
    {"label": "性别", "text": "{{性别}}"},
    {"label": "出生日期", "text": "{{出生日期}}"},
    {"label": "民族", "text": "{{民族}}"},
    {"label": "户籍地", "text": "{{户籍地}}"},
    {"label": "现住址", "text": "{{现住址}}"},
    {"label": "罪名", "text": "{{罪名}}"},
    {"label": "办案单位", "text": "{{办案单位}}"},
    {"label": "法律条文引用", "text": "《中华人民共和国刑法》第{{条}}条"},
    {"label": "刑诉法引用", "text": "《中华人民共和国刑事诉讼法》第{{条}}条"},
    {"label": "落款日期", "text": "{{落款日期}}"},
]


def load_quick_inserts():
    """加载快捷插入列表，无配置则返回默认列表"""
    if os.path.exists(QUICK_INSERT_PATH):
        try:
            with open(QUICK_INSERT_PATH, "r", encoding="utf-8") as f:
                items = json.load(f)
            if items:
                return items
        except Exception:
            pass
    return list(DEFAULT_QUICK_INSERTS)


def save_quick_inserts(items):
    """保存快捷插入列表"""
    with open(QUICK_INSERT_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


# ================================================================
# 占位符定位（右键菜单辅助）
# ================================================================
_PH_FULL_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")


def find_placeholder_at(text, cursor_pos):
    """返回光标所在位置的占位符信息 (full_match, field_name, start, end)，不在占位符内返回 None"""
    for m in _PH_FULL_RE.finditer(text):
        if m.start() <= cursor_pos < m.end():
            return (m.group(0), m.group(1).strip(), m.start(), m.end())
    return None


# ================================================================
# 自动识别常见字段（模板制作时辅助挖空）
# ================================================================
_AUTO_PATTERNS = [
    ("身份证号（18位）", r"\b\d{17}[\dXx]\b"),
    ("身份证号（15位）", r"\b\d{15}\b"),
    ("法律条款引用", r"第[一二三四五六七八九十百千\d]+条(?:之[一二三四五六七八九十\d]+)?"),
    ("日期（中文）", r"\d{4}年\d{1,2}月\d{1,2}日"),
    ("手机号码", r"\b1[3-9]\d{9}\b"),
    ("公文发文字号", r"[A-Za-z\u4e00-\u9fff]+〔\d+〕\d+号"),
]


def detect_auto_fields(text):
    """扫描文本，返回 [(匹配文本, 建议字段名, 类型标签), ...]，按匹配文本去重"""
    seen_texts = set()
    results = []
    for label, pattern in _AUTO_PATTERNS:
        for m in re.finditer(pattern, text):
            matched = m.group(0)
            if matched not in seen_texts:
                seen_texts.add(matched)
                # 生成建议字段名
                if "身份证" in label:
                    field = "身份证号"
                elif "法律条款" in label:
                    field = "条款"
                elif "日期" in label:
                    field = "日期"
                elif "手机" in label:
                    field = "联系电话"
                elif "发文字号" in label:
                    field = "发文字号"
                else:
                    field = label
                results.append((matched, field, label))
    return results
