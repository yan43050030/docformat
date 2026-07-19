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

from app.presets import config_dir

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")

TEMPLATE_DIR = os.path.join(config_dir(), "templates")


def _template_dirs_config_path():
    """多模板目录列表的配置文件路径"""
    return os.path.join(config_dir(), "template_dirs.json")


def load_template_dirs():
    """返回所有模板目录列表。有保存配置则用配置，否则回退到默认目录"""
    cfg = _template_dirs_config_path()
    if os.path.exists(cfg):
        try:
            with open(cfg, "r", encoding="utf-8") as f:
                saved = json.load(f)
            dirs = []
            for d in saved:
                d = os.path.expanduser(d)
                if d not in dirs:
                    dirs.append(d)
            if dirs:
                return dirs
        except Exception:
            pass
    return [TEMPLATE_DIR]


def save_template_dirs(dirs):
    """保存模板目录列表（完整保存，去重）"""
    cfg = _template_dirs_config_path()
    seen = []
    uniq = []
    for d in dirs:
        d = os.path.expanduser(d)
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


def extract_fields(template_text):
    """按出现顺序提取去重后的占位符字段名"""
    seen = []
    for m in PLACEHOLDER_RE.finditer(template_text):
        name = m.group(1).strip()
        if name not in seen:
            seen.append(name)
    return seen


def parse_template(template_text):
    """解析模板 → {'title', 'body':[{'type','text'}], 'meta':{}}

    模板格式约定（勿改，两个模块靠此对接）：
      - 标题: 以 "标题:" 开头的行
      - "一、" 开头 → 一级标题 (h1)
      - "（一）" 开头 → 二级标题 (h2)
      - 其余 → 正文 (body)
      - ---META--- 之后 → 落款单位/日期
    """
    body_part, meta = template_text, {}
    if "---META---" in template_text:
        body_part, meta_part = template_text.split("---META---", 1)
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
