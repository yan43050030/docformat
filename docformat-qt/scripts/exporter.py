# -*- coding: utf-8 -*-
"""导出 PDF：优先 Word/WPS（版式最忠实），否则回退 LibreOffice。

公文定稿后通常要发 PDF，避免对方机器缺字体导致错版。导出走的是本机
排版引擎，所以用 Word/WPS 导出的版式与在 Word 里看到的一致；用
LibreOffice 导出在缺方正字体的机器上可能有细微差异，会如实提示。
"""
import logging
import os
import shutil
import subprocess
import sys
import tempfile

logger = logging.getLogger('docformat.exporter')


def _export_via_word(input_path, output_path, log=None):
    """Windows：用 Word/WPS COM 导出 PDF，顺带更新目录域等字段。"""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return False, '本机缺少 pywin32'
    pythoncom.CoInitialize()
    app = None
    app_name = ''
    for prog, name in (('Word.Application', 'Microsoft Word'),
                       ('Kwps.Application', 'WPS'), ('wps.Application', 'WPS')):
        try:
            app = win32.Dispatch(prog)
            app_name = name
            break
        except Exception:
            continue
    if app is None:
        pythoncom.CoUninitialize()
        return False, '本机未安装 Word/WPS'
    try:
        try:
            app.Visible = False
        except Exception:
            pass
        doc = app.Documents.Open(os.path.abspath(input_path), ReadOnly=False)
        try:
            # 更新目录等域，让 PDF 里的页码是最终值
            try:
                for i in range(doc.TablesOfContents.Count):
                    doc.TablesOfContents(i + 1).Update()
            except Exception:
                pass
            try:
                doc.Fields.Update()
            except Exception:
                pass
            WD_FORMAT_PDF = 17
            doc.SaveAs2(os.path.abspath(output_path), FileFormat=WD_FORMAT_PDF)
        finally:
            doc.Close(False)
        return True, app_name
    except Exception as exc:
        return False, '{} 导出失败：{}'.format(app_name or 'Office', exc)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _export_via_libreoffice(input_path, output_path, log=None):
    try:
        from app.converter_linux import find_soffice
    except ImportError:
        return False, '找不到转换模块'
    soffice = find_soffice()
    if not soffice:
        return False, '本机未安装 LibreOffice（或 Word/WPS）'
    tmp = tempfile.mkdtemp(prefix='docformat_pdf_')
    try:
        proc = subprocess.run(
            [soffice, '--headless', '--norestore', '--convert-to', 'pdf',
             '--outdir', tmp, os.path.abspath(input_path)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=300)
        produced = os.path.join(
            tmp, os.path.splitext(os.path.basename(input_path))[0] + '.pdf')
        if proc.returncode == 0 and os.path.exists(produced):
            shutil.move(produced, output_path)
            return True, 'LibreOffice'
        err = (proc.stderr or b'').decode('utf-8', errors='replace').strip()
        return False, 'LibreOffice 导出失败：{}'.format(err or '未产出 PDF')
    except subprocess.TimeoutExpired:
        return False, 'LibreOffice 导出超时'
    except Exception as exc:
        return False, 'LibreOffice 导出失败：{}'.format(exc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def export_pdf(input_path, output_path, log=None):
    """导出 PDF，返回 (成功, 说明)。"""
    if sys.platform == 'win32':
        ok, info = _export_via_word(input_path, output_path, log)
        if ok:
            return True, info
        if log:
            log('info', '{}，改用 LibreOffice 导出'.format(info))
    return _export_via_libreoffice(input_path, output_path, log)


def pdf_output_path(input_path, suffix=''):
    d = os.path.dirname(input_path)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    base = os.path.join(d, '{}{}.pdf'.format(stem, suffix))
    if not os.path.exists(base):
        return base
    for n in range(2, 1000):
        cand = os.path.join(d, '{}{}({}).pdf'.format(stem, suffix, n))
        if not os.path.exists(cand):
            return cand
    return base
