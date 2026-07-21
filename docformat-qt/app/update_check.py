# -*- coding: utf-8 -*-
"""启动时后台检查 GitHub Releases 是否有新版本。

内网/离线环境请求会静默失败，不打扰用户、不影响启动。
"""
import json
import re

from PyQt5.QtCore import QThread, pyqtSignal

RELEASES_API = 'https://api.github.com/repos/yan43050030/docformat/releases/latest'
RELEASES_PAGE = 'https://github.com/yan43050030/docformat/releases/latest'


def _version_tuple(tag):
    m = re.findall(r'\d+', tag or '')
    return tuple(int(x) for x in m[:3]) if m else None


class UpdateChecker(QThread):
    newVersion = pyqtSignal(str, str)   # tag, url

    def __init__(self, current_version, parent=None):
        super(UpdateChecker, self).__init__(parent)
        self.current_version = current_version

    def run(self):
        try:
            from urllib.request import Request, urlopen
            req = Request(RELEASES_API, headers={'User-Agent': 'DocFormatPro'})
            with urlopen(req, timeout=6) as resp:
                data = json.loads(resp.read().decode('utf-8', errors='replace'))
            tag = data.get('tag_name') or ''
            latest = _version_tuple(tag)
            current = _version_tuple(self.current_version)
            if latest and current and latest > current:
                self.newVersion.emit(tag, data.get('html_url') or RELEASES_PAGE)
        except Exception:
            pass   # 离线/内网/接口异常一律静默
