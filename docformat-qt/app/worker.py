# -*- coding: utf-8 -*-
"""批处理线程：四种模式分发 + .doc/.wps 预转换 + 进度/日志信号"""
import os
import re
import sys
import tempfile

from PyQt5.QtCore import QThread, pyqtSignal

MODE_FULL = 'full'
MODE_DIAGNOSE = 'diagnose'
MODE_PUNCTUATION = 'punctuation'
MODE_AI_PASTE = 'ai_paste'


def output_path_for(input_path, suffix):
    d = os.path.dirname(input_path)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(d, '{}{}.docx'.format(stem, suffix))


def ensure_docx(path, log):
    """非 docx 输入先转换为临时 docx，返回 (工作路径, 是否为临时文件)"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        return path, False
    log('info', '检测到 {} 格式，正在转换为 docx...'.format(ext))
    if sys.platform == 'win32':
        from scripts import converter
        tmp = os.path.join(tempfile.mkdtemp(prefix='docformat_'),
                           os.path.splitext(os.path.basename(path))[0] + '.docx')
        converter.convert_to_docx(path, tmp)
        return tmp, True
    else:
        from app import converter_linux
        return converter_linux.convert_to_docx(path), True


MD_FENCE = re.compile(r'^\s*```')
MD_HEADING = re.compile(r'^\s*#{1,6}\s+')
MD_BULLET = re.compile(r'^\s*[-*+]\s+')
MD_QUOTE = re.compile(r'^\s*>\s?')
MD_HR = re.compile(r'^\s*([-*_]\s*){3,}$')
MD_BOLD = re.compile(r'(\*\*|__)(.+?)\1')
MD_ITALIC = re.compile(r'(?<![*_])([*_])([^*_\n]+?)\1(?![*_])')
MD_LINK = re.compile(r'\[([^\]]+)\]\([^)]*\)')
MD_CODE = re.compile(r'`([^`]*)`')


def clean_markdown(text):
    """把 AI 生成的 Markdown 清洗为纯文本段落列表"""
    lines = []
    in_fence = False
    for raw in text.splitlines():
        if MD_FENCE.match(raw):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        line = raw.rstrip()
        if MD_HR.match(line):
            continue
        line = MD_HEADING.sub('', line)
        line = MD_QUOTE.sub('', line)
        line = MD_BULLET.sub('', line)
        line = MD_BOLD.sub(r'\2', line)
        line = MD_ITALIC.sub(r'\2', line)
        line = MD_LINK.sub(r'\1', line)
        line = MD_CODE.sub(r'\1', line)
        lines.append(line.strip())
    # 折叠连续空行
    paras = []
    prev_blank = True
    for line in lines:
        if line:
            paras.append(line)
            prev_blank = False
        else:
            if not prev_blank:
                paras.append('')
            prev_blank = True
    while paras and not paras[-1]:
        paras.pop()
    return paras


def generate_docx_from_text(text, out_path):
    """纯文本/Markdown → 临时 docx（后续交给排版引擎）"""
    from docx import Document
    doc = Document()
    for para in clean_markdown(text):
        doc.add_paragraph(para)
    doc.save(out_path)
    return out_path


def build_diagnose_report(filename, results):
    lines = ['◆ {}'.format(filename)]
    total = 0

    punct = results.get('punctuation') or []
    if punct:
        lines.append('  【标点符号】{} 处问题：'.format(len(punct)))
        shown = {}
        for it in punct:
            key = '{}「{}」'.format(it.get('type', '?'), it.get('char', ''))
            shown.setdefault(key, []).append(it.get('para', 0))
        for key, paras in shown.items():
            preview = ', '.join(str(p) for p in paras[:8])
            more = ' 等' if len(paras) > 8 else ''
            lines.append('    - {}：第 {} 段{}'.format(key, preview, more))
        total += len(punct)

    for it in results.get('numbering') or []:
        lines.append('  【序号格式】{}'.format(it.get('detail', it.get('type', ''))))
        total += 1

    for it in results.get('paragraph') or []:
        if 'paras' in it:
            paras = it['paras']
            preview = ', '.join(str(p) for p in paras[:10])
            more = ' 等 {} 段'.format(len(paras)) if len(paras) > 10 else ''
            lines.append('  【段落格式】{}：第 {} 段{}'.format(it.get('type', ''), preview, more))
        else:
            lines.append('  【段落格式】{}：{}'.format(it.get('type', ''), it.get('detail', '')))
        total += 1

    for it in results.get('font') or []:
        lines.append('  【字体】{}：{}'.format(it.get('type', ''), it.get('detail', '')))
        total += 1

    if total == 0:
        lines.append('  ✓ 未发现格式问题')
    else:
        lines[0] = '◆ {} — 共 {} 类问题'.format(filename, total)
    return '\n'.join(lines)


class ProcessWorker(QThread):
    progressChanged = pyqtSignal(int)                # 0-100
    logMessage = pyqtSignal(str, str)                # level, message
    fileFinished = pyqtSignal(str, str)              # input, output ('' = 无输出)
    diagnoseReady = pyqtSignal(str)                  # 汇总报告
    allFinished = pyqtSignal(int, int)               # ok, fail

    def __init__(self, files, mode, preset_name, custom_settings, suffix, parent=None):
        super(ProcessWorker, self).__init__(parent)
        self.files = list(files)
        self.mode = mode
        self.preset_name = preset_name
        self.custom_settings = custom_settings
        self.suffix = suffix or '_processed'
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def _log(self, level, msg):
        self.logMessage.emit(level, msg)

    def run(self):
        com_inited = False
        if sys.platform == 'win32':
            needs_com = any(not f.lower().endswith('.docx') for f in self.files)
            if needs_com:
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    com_inited = True
                except Exception:
                    pass

        ok, fail = 0, 0
        reports = []
        n = len(self.files)
        try:
            for i, path in enumerate(self.files):
                if self._cancelled:
                    self._log('warning', '已取消，剩余文件未处理')
                    break
                base = os.path.basename(path)
                self._log('info', '正在处理: {}'.format(base))
                try:
                    work, _is_tmp = ensure_docx(path, self._log)
                    if self.mode == MODE_DIAGNOSE:
                        reports.append(self._diagnose(work, base))
                        self.fileFinished.emit(path, '')
                    elif self.mode == MODE_PUNCTUATION:
                        out = output_path_for(path, self.suffix)
                        from scripts import punctuation
                        punctuation.process_document(work, out)
                        self._log('success', '已完成: {} → {}'.format(base, os.path.basename(out)))
                        self.fileFinished.emit(path, out)
                    else:  # MODE_FULL
                        out = output_path_for(path, self.suffix)
                        self._run_format(work, out, i, n)
                        self._log('success', '已完成: {} → {}'.format(base, os.path.basename(out)))
                        self.fileFinished.emit(path, out)
                    ok += 1
                except Exception as e:
                    fail += 1
                    self._log('error', '处理失败 {}: {}'.format(base, e))
                self.progressChanged.emit(int((i + 1) * 100 / n) if n else 100)

            if self.mode == MODE_DIAGNOSE and reports:
                self.diagnoseReady.emit('\n\n'.join(reports))
        finally:
            if com_inited:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        self.allFinished.emit(ok, fail)

    def _run_format(self, work, out, file_idx, total_files):
        """智能一键 = 标点修复(静默) → 排版规范，与原版 docformat-gui 流程一致"""
        from scripts import punctuation
        from scripts.formatter import format_document

        tmp = os.path.join(tempfile.mkdtemp(prefix='docformat_'), 'punct.docx')
        punctuation.process_document(work, tmp)

        def cb(cur, total, _stage):
            if total:
                inner = float(cur) / total
                overall = (file_idx + inner) * 100.0 / total_files
                self.progressChanged.emit(int(overall))

        format_document(
            tmp, out,
            preset_name=self.preset_name,
            progress_callback=cb,
            custom_settings=self.custom_settings,
        )

    def _diagnose(self, work, display_name):
        from docx import Document
        from scripts import analyzer
        doc = Document(work)
        results = {
            'punctuation': analyzer.analyze_punctuation(doc),
            'numbering': analyzer.analyze_numbering(doc),
            'paragraph': analyzer.analyze_paragraph_format(doc),
            'font': analyzer.analyze_font(doc),
        }
        report = build_diagnose_report(display_name, results)
        self._log('info', '诊断完成: {}'.format(display_name))
        return report


class AiPasteWorker(QThread):
    """AI 粘贴生成：文本 → 临时 docx → 排版引擎 → 输出"""
    logMessage = pyqtSignal(str, str)
    finishedWith = pyqtSignal(bool, str)   # success, out_path_or_error

    def __init__(self, text, out_path, preset_name, custom_settings, parent=None):
        super(AiPasteWorker, self).__init__(parent)
        self.text = text
        self.out_path = out_path
        self.preset_name = preset_name
        self.custom_settings = custom_settings

    def run(self):
        try:
            tmp = os.path.join(tempfile.mkdtemp(prefix='docformat_ai_'), 'draft.docx')
            self.logMessage.emit('info', '正在从文本生成文档...')
            generate_docx_from_text(self.text, tmp)
            self.logMessage.emit('info', '正在修复标点并应用排版规范...')
            from scripts import punctuation
            from scripts.formatter import format_document
            tmp2 = os.path.join(os.path.dirname(tmp), 'punct.docx')
            punctuation.process_document(tmp, tmp2)
            format_document(tmp2, self.out_path,
                            preset_name=self.preset_name,
                            custom_settings=self.custom_settings)
            self.logMessage.emit('success', '已生成: {}'.format(self.out_path))
            self.finishedWith.emit(True, self.out_path)
        except Exception as e:
            self.logMessage.emit('error', '生成失败: {}'.format(e))
            self.finishedWith.emit(False, str(e))
