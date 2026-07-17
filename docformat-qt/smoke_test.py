# -*- coding: utf-8 -*-
"""端到端冒烟测试：生成样例公文 → 三种模式处理 → 断言结果"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Cm

OUT_DIR = os.path.join(os.path.dirname(__file__), '_smoke')
os.makedirs(OUT_DIR, exist_ok=True)
SRC = os.path.join(OUT_DIR, 'sample.docx')


def make_sample():
    doc = Document()
    doc.add_paragraph('秘密★1年')
    doc.add_paragraph('关于开展2026年度安全生产检查的通知')
    doc.add_paragraph('各部门、各单位:')
    doc.add_paragraph('为深入贯彻落实上级部署要求(含附件),现将有关事项通知如下.')
    doc.add_paragraph('一、总体要求')
    doc.add_paragraph('坚持"安全第一、预防为主"的方针,全面排查隐患。')
    doc.add_paragraph('(一)检查范围')
    doc.add_paragraph('1.生产车间及仓储区域。')
    doc.add_paragraph('(1)重点检查电气线路。')
    doc.add_paragraph('二、时间安排')
    doc.add_paragraph('检查工作自2026年7月20日起至8月10日结束...')
    doc.add_paragraph('特此通知。')
    doc.add_paragraph('附件:安全检查表')
    doc.add_paragraph('某某公司办公室')
    doc.add_paragraph('2026年7月17日')
    doc.save(SRC)
    print('[1] 样例文档已生成:', SRC)


def test_full():
    """按 worker 的智能一键链路：标点修复 → 排版"""
    from scripts.punctuation import process_document
    from scripts.formatter import format_document
    mid = os.path.join(OUT_DIR, 'sample_punct_stage.docx')
    process_document(SRC, mid)
    out = os.path.join(OUT_DIR, 'sample_full.docx')
    stages = []
    format_document(mid, out, preset_name='official',
                    progress_callback=lambda c, t, s: stages.append((c, t, s)))
    doc = Document(out)
    sec = doc.sections[0]
    top, bottom = sec.top_margin.cm, sec.bottom_margin.cm
    left, right = sec.left_margin.cm, sec.right_margin.cm
    assert abs(top - 3.7) < 0.05 and abs(bottom - 3.5) < 0.05, '上下边距错误: {} {}'.format(top, bottom)
    assert abs(left - 2.8) < 0.05 and abs(right - 2.6) < 0.05, '左右边距错误: {} {}'.format(left, right)

    title_run = doc.paragraphs[1].runs[0]
    fonts = set()
    from docx.oxml.ns import qn
    rpr = title_run._element.rPr
    ea = rpr.rFonts.get(qn('w:eastAsia')) if rpr is not None and rpr.rFonts is not None else None
    assert ea == '方正小标宋简体', '标题字体错误: {}'.format(ea)
    assert title_run.font.size.pt == 22, '标题字号错误: {}'.format(title_run.font.size.pt)

    # 密级标识：黑体、顶格左对齐、不缩进
    sec_para = doc.paragraphs[0]
    assert '秘密' in sec_para.text, '密级行丢失: {}'.format(sec_para.text)
    sec_run = sec_para.runs[0]
    sec_rpr = sec_run._element.rPr
    sec_ea = sec_rpr.rFonts.get(qn('w:eastAsia')) if sec_rpr is not None and sec_rpr.rFonts is not None else None
    assert sec_ea == '黑体', '密级字体错误: {}'.format(sec_ea)
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    assert sec_para.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.CENTER, '密级不应居中'

    body_text = '\n'.join(p.text for p in doc.paragraphs)
    assert '（含附件）' in body_text, '英文括号未转换'
    assert '……' in body_text, '省略号未规范化'
    print('[2] 智能一键: 边距/标题字体字号/标点转换 全部通过 (进度回调 {} 次)'.format(len(stages)))


def test_punctuation():
    from scripts.punctuation import process_document
    out = os.path.join(OUT_DIR, 'sample_punct.docx')
    process_document(SRC, out)
    text = '\n'.join(p.text for p in Document(out).paragraphs)
    assert '（含附件）' in text, '标点模式: 括号未转换'
    print('[3] 标点修复: 通过')


def test_diagnose():
    from scripts import analyzer
    doc = Document(SRC)
    results = {
        'punctuation': analyzer.analyze_punctuation(doc),
        'numbering': analyzer.analyze_numbering(doc),
        'paragraph': analyzer.analyze_paragraph_format(doc),
        'font': analyzer.analyze_font(doc),
    }
    n = sum(len(v) for v in results.values())
    assert len(results['punctuation']) > 0, '诊断应发现英文标点问题'
    from app.worker import build_diagnose_report
    report = build_diagnose_report('sample.docx', results)
    print('[4] 格式诊断: 发现 {} 项问题, 报告生成 OK'.format(n))
    print('    ' + report.splitlines()[0])


def test_custom_preset():
    from scripts.formatter import format_document
    from app.presets import PresetManager
    mgr = PresetManager()
    key = mgr.create('冒烟测试模板')
    preset = mgr.get(key)
    preset['page']['top'] = 5.0
    preset['body']['size'] = 14
    mgr.update(key, preset)

    mgr2 = PresetManager()  # 重新加载验证持久化
    assert key in mgr2.user, '用户模板未持久化'
    assert mgr2.get(key)['page']['top'] == 5.0, '模板参数未保存'

    out = os.path.join(OUT_DIR, 'sample_custom.docx')
    name, custom = mgr2.engine_args(key)
    format_document(SRC, out, preset_name=name, custom_settings=custom)
    doc = Document(out)
    assert abs(doc.sections[0].top_margin.cm - 5.0) < 0.05, '自定义边距未生效'
    mgr2.delete(key)
    print('[5] 自定义模板: 持久化 + 引擎生效 通过')


def test_ai_paste():
    from app.worker import generate_docx_from_text, clean_markdown
    md = "# 关于测试的通知\n\n**各部门**:\n\n- 第一项工作\n- 第二项工作\n\n```\ncode block skip\n```\n\n特此通知。"
    paras = clean_markdown(md)
    assert '关于测试的通知' in paras[0], 'markdown 标题清洗失败'
    assert not any('```' in p or 'code block' in p for p in paras), '代码块未剔除'
    assert '各部门:' in '\n'.join(paras), '加粗标记未清除'
    out = os.path.join(OUT_DIR, 'ai_draft.docx')
    generate_docx_from_text(md, out)
    assert os.path.exists(out)
    from scripts.formatter import format_document
    final = os.path.join(OUT_DIR, 'ai_final.docx')
    format_document(out, final, preset_name='official')
    assert os.path.exists(final)
    print('[6] AI 粘贴生成: markdown 清洗 + 生成 + 排版 通过')


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    make_sample()
    test_full()
    test_punctuation()
    test_diagnose()
    test_custom_preset()
    test_ai_paste()
    print('\n全部冒烟测试通过 ✓')
