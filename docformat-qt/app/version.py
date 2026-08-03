# -*- coding: utf-8 -*-
"""版本号常量。

单独一个模块，不 import 任何东西——尤其是不碰 PyQt。main.py 的
`--version` 要在导入 Qt 之前就能取到它：没有 X 的环境（服务器、SSH、
打包机的 CI）一旦碰上 Qt，就会先撞 "Could not connect to any X display"，
根本轮不到把版本号打出来。
"""
VERSION = '5.4.0'
