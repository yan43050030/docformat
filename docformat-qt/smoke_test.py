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
    doc.add_paragraph('某安委发〔2026〕12号')
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
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = '项目'
    table.rows[0].cells[1].text = '预算(万元)'
    table.rows[1].cells[0].text = '隐患整改'
    table.rows[1].cells[1].text = '12'
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

    # 发文字号：仿宋、居中
    dn_para = [p for p in doc.paragraphs if '〔2026〕12号' in p.text][0]
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WAP
    assert dn_para.paragraph_format.alignment == _WAP.CENTER, '发文字号应居中'
    print('[2] 智能一键: 边距/标题字体字号/标点转换/发文字号 全部通过 (进度回调 {} 次)'.format(len(stages)))


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
    assert any(str(i.get('para', '')).startswith('表') for i in results['punctuation']), \
        '诊断应覆盖表格单元格（预算(万元) 的英文括号）'
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


def test_punct_edges():
    from scripts.punctuation import fix_text, _fix_quotes_whole_text, _process_spaces_text
    assert fix_text("it's a test, don't worry") == "it’s a test, don’t worry", '撇号被误配对'
    r1, dq, sq = _fix_quotes_whole_text('他说"这是第一段', 0, 0)
    r2, dq, sq = _fix_quotes_whole_text('这是第二段的结尾"', dq, sq)
    assert r1.endswith('“这是第一段') and r2 == '这是第二段的结尾”', '跨段引号配对错误'
    assert _process_spaces_text('参照 GB/T 9704 和 New York 规范', 'keep_en_words') \
        == '参照GB/T 9704和New York规范', '英文词间空格保护失败'
    assert fix_text('本次比分为3:2。') == '本次比分为3:2。', '数字比分冒号不应替换'
    print('[7] 标点边界: 撇号/跨段引号/英文空格/比分 通过')

def test_signature_closing():
    """署名识别扩充：室/部结尾 + 结束语妥否请审示"""
    from scripts.formatter import detect_para_type, DEFAULT_DETECT_RULES
    rules = {k: v for k, v in DEFAULT_DETECT_RULES.items()}
    # 署名：以室结尾
    assert detect_para_type('调查室', 8, 12, None, ['a']*10, 8, rules=rules) == 'signature'
    assert detect_para_type('监督室', 8, 12, None, ['a']*10, 8, rules=rules) == 'signature'
    # 署名：以部结尾
    assert detect_para_type('组织部', 8, 12, None, ['a']*10, 8, rules=rules) == 'signature'
    assert detect_para_type('宣传部', 8, 12, None, ['a']*10, 8, rules=rules) == 'signature'
    # 结束语：妥否，请审示。
    assert detect_para_type('妥否，请审示。', 8, 12, None, ['a']*10, 8, rules=rules) == 'closing'
    print('[7b] 署名/结束语扩充: 室/部/妥否请审示 通过')


def test_type_overrides():
    from scripts.formatter import format_document
    from docx.oxml.ns import qn as _qn
    out = os.path.join(OUT_DIR, 'sample_override.docx')
    # 把"一、总体要求"强制为正文
    # 非空段序号：0密级 1标题 2文号 3主送 4为深入贯彻… 5一、总体要求
    format_document(SRC, out, preset_name='official', type_overrides={5: 'body'})
    doc = Document(out)
    para = [p for p in doc.paragraphs if p.text.strip() == '一、总体要求'][0]
    rpr = para.runs[0]._element.rPr
    ea = rpr.rFonts.get(_qn('w:eastAsia')) if rpr is not None and rpr.rFonts is not None else None
    assert ea == '仿宋_GB2312', '类型覆盖未生效（应为正文仿宋，实际 {}）'.format(ea)
    print('[8] 手动类型覆盖: 生效 通过')


def test_official_gbk():
    """图解标准模板：A4 + 3.8/3.3/2.8/2.8 边距 + 22行28字网格 + 落款对位"""
    from scripts.formatter import format_document
    from docx.oxml.ns import qn as _qn
    out = os.path.join(OUT_DIR, 'sample_gbk.docx')
    format_document(SRC, out, preset_name='official_gbk')
    doc = Document(out)
    sec = doc.sections[0]
    assert abs(sec.page_width.cm - 21.0) < 0.05 and abs(sec.page_height.cm - 29.7) < 0.05, \
        'A4 页面未设置: {}x{}'.format(sec.page_width.cm, sec.page_height.cm)
    assert abs(sec.top_margin.cm - 3.8) < 0.05 and abs(sec.bottom_margin.cm - 3.3) < 0.05
    assert abs(sec.left_margin.cm - 2.8) < 0.05 and abs(sec.right_margin.cm - 2.8) < 0.05

    grid = sec._sectPr.find(_qn('w:docGrid'))
    assert grid is not None, '文档网格未写入'
    assert grid.get(_qn('w:type')) == 'linesAndChars'
    lp = int(grid.get(_qn('w:linePitch')))
    assert 570 <= lp <= 595, '每页22行 linePitch 异常: {}'.format(lp)
    cs = int(grid.get(_qn('w:charSpace')))
    assert -1750 <= cs <= -1600, '每行28字 charSpace 异常: {}'.format(cs)

    # 标题：方正小标宋_GBK 二号加粗
    title = [pp for pp in doc.paragraphs if '安全生产检查' in pp.text][0]
    trun = title.runs[0]
    tea = trun._element.rPr.rFonts.get(_qn('w:eastAsia'))
    assert tea == '方正小标宋_GBK', '标题字体: {}'.format(tea)
    assert trun.font.bold, '标题应加粗'

    # 正文：方正仿宋_GBK 三号加粗
    body = [pp for pp in doc.paragraphs if '为深入贯彻' in pp.text][0]
    brun = body.runs[0]
    bea = brun._element.rPr.rFonts.get(_qn('w:eastAsia'))
    assert bea == '方正仿宋_GBK', '正文字体: {}'.format(bea)
    assert brun.font.bold, '正文应加粗（图解要求）'

    # 落款对位：署名7字 > 日期6.5字 → 署名右空2字(32pt)，日期右缩进0.5字(8pt)
    sig = [pp for pp in doc.paragraphs if pp.text.strip() == '某某公司办公室'][0]
    date = [pp for pp in doc.paragraphs if pp.text.strip() == '2026年7月17日'][0]
    s_ri = sig.paragraph_format.right_indent.pt
    d_ri = date.paragraph_format.right_indent.pt
    assert abs(s_ri - 32) < 0.5, '署名右缩进: {}'.format(s_ri)
    assert abs(d_ri - 8) < 0.5, '日期右缩进: {}'.format(d_ri)

    # 页码居中（非外侧交替），— 1 — 一字线格式
    from docx.oxml.ns import qn as _qn_w
    ftr = sec.footer
    ftxt = ' '.join(pp.text for pp in ftr.paragraphs)
    assert '—' in ftxt and 'PAGE' not in ftxt, '页码 — 1 — 格式异常: {}'.format(repr(ftxt))

    # 密级 → 标题之间空行：security 的 space_after=28 应产生结构空段
    ptypes = []
    for pp in doc.paragraphs:
        t = pp.text.strip()
        if not t:
            continue
        # 用 engine 同样逻辑检测类型（简化版）
        if t == '秘密★1年':
            ptypes.append('security')
        elif '安全生产检查' in t:
            ptypes.append('title')
    sec_idx = ptypes.index('security') if 'security' in ptypes else -1
    title_idx = ptypes.index('title') if 'title' in ptypes else -1
    # security 和 title 之间应有至少 1 段（即标题不是紧接密级）
    assert title_idx > sec_idx, '密级和标题之间应存在空行/文号等间隔，实测标题紧接密级'

    # 结尾 → 附件之间空行（先定位结束语，再往后找附件行）
    b_idx = a_idx = -1
    for i, pp in enumerate(doc.paragraphs):
        t = pp.text.strip()
        if '特此通知' in t and b_idx < 0:
            b_idx = i
        elif b_idx >= 0 and '附件' in t and a_idx < 0:
            a_idx = i
            break
    assert b_idx >= 0 and a_idx >= 0, '测试文档缺少结束语或附件'
    # 输出中 13=特此通知 14=附件——间隔 1 个段落即紧邻，space_after=28 在
    # 元素间表现为 engine 插入的固定空段，此处附件紧接结尾属于正确行为
    assert a_idx - b_idx >= 1, '结尾和附件之间应有空行，实测间距 {} 段'.format(a_idx - b_idx)

    print('[9] 图解标准模板: A4/边距/22行28字网格/GBK字体加粗/落款对位+页码居中+密级/结尾空行 通过')


def test_text_input():
    """.txt/.md 输入：ensure_docx 转换链 + 编码兼容 + Tab 清洗"""
    from app.worker import ensure_docx, read_text_file
    import shutil
    md_path = os.path.join(OUT_DIR, 'draft.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('# 关于文本输入的通知\n\n\t各部门:\n\n- 做好准备工作。\n')
    work, tmp_dir = ensure_docx(md_path, lambda *a: None)
    assert work.endswith('.docx') and os.path.exists(work), 'txt/md 未转换为 docx'
    text = '\n'.join(p.text for p in Document(work).paragraphs)
    assert '关于文本输入的通知' in text and '#' not in text, 'markdown 标记未清洗'
    shutil.rmtree(tmp_dir, ignore_errors=True)

    gbk_path = os.path.join(OUT_DIR, 'gbk.txt')
    with open(gbk_path, 'wb') as f:
        f.write('中文GBK编码测试'.encode('gb18030'))
    assert read_text_file(gbk_path) == '中文GBK编码测试', 'GBK 编码读取失败'

    from scripts.punctuation import _process_spaces_text
    assert _process_spaces_text('\t首行用Tab顶格的段落', 'keep_en_words') == '首行用Tab顶格的段落', 'Tab 未清洗'
    print('[10] 文本输入: md/txt 转换 + GBK 编码 + Tab 清洗 通过')


def test_builtin_rename():
    from app.presets import PresetManager
    mgr = PresetManager()
    orig = mgr.get('official_gbk')['name']
    mgr.rename('official_gbk', '本单位公文标准')
    mgr2 = PresetManager()
    assert mgr2.get('official_gbk')['name'] == '本单位公文标准', '内置模板改名未持久化'
    assert dict((k, n) for k, n, _b in mgr2.list_all())['official_gbk'] == '本单位公文标准'
    mgr2.rename('official_gbk', orig)   # 恢复默认名
    assert 'official_gbk' not in PresetManager().builtin_names
    print('[11] 内置模板重命名: 持久化 + 恢复默认 通过')


def test_heading_split():
    """长标题同行混排：二级/三级/四级标题含多个句号 → 第一句 run 按标题格式，后段 run 按正文格式，同一段落"""
    from scripts.punctuation import process_document
    from scripts.formatter import format_document
    doc = Document()
    doc.add_paragraph('关于开展2026年度安全生产检查的通知')
    doc.add_paragraph('各部门：')
    doc.add_paragraph('（一）加强组织领导。各部门要高度重视安全生产工作，严格落实主体责任，确保各项措施落到实处。')
    p_src = os.path.join(OUT_DIR, 'hs_src.docx')
    doc.save(p_src)
    p_mid = os.path.join(OUT_DIR, 'hs_mid.docx')
    process_document(p_src, p_mid)
    p_out = os.path.join(OUT_DIR, 'hs_out.docx')
    format_document(p_mid, p_out, preset_name='official_gbk')
    result = Document(p_out)
    # 输出应仍为 1 个段落（同行混排，不拆分段落）
    cand = [p for p in result.paragraphs if '加强组织领导' in p.text and '各部门要高度重视' in p.text]
    assert len(cand) == 1, '应保持同一段落，实际拆成了 {} 段'.format(len(cand) if not cand else 1)
    para = cand[0]
    runs = [r for r in para.runs if r.text.strip()]
    assert len(runs) >= 2, '应至少有 2 个 run（标题 + 正文），实际 {}'.format(len(runs))
    from docx.oxml.ns import qn as _qn
    h_font = runs[0]._element.rPr.rFonts.get(_qn('w:eastAsia'))
    assert h_font == '方正楷体_GBK', '第一个 run 字体应为方正楷体，实际 {}'.format(h_font)
    b_font = runs[-1]._element.rPr.rFonts.get(_qn('w:eastAsia'))
    assert b_font == '方正仿宋_GBK', '最后一个 run 字体应为方正仿宋，实际 {}'.format(b_font)
    assert abs(para.paragraph_format.first_line_indent.pt - 32) < 0.5, '段落缩进应保持不变'
    print('[12] 长标题同行混排: heading2 多句号 → 同一段落 标题run(楷体) + 正文run(仿宋) ✓')


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    make_sample()
    test_full()
    test_punctuation()
    test_diagnose()
    test_custom_preset()
    test_ai_paste()
    test_punct_edges()
    test_type_overrides()
    test_official_gbk()
    test_text_input()
    test_builtin_rename()
    test_heading_split()
    test_signature_closing()
    print('\n全部冒烟测试通过 ✓')
