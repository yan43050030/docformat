# -*- coding: utf-8 -*-
"""主题系统：四套配色 + QSS 生成 + QSettings 持久化"""
from PyQt5.QtCore import QSettings

THEMES = {
    'paper': {
        'name': '纸质', 'desc': '米白纸面 · 朱红点缀',
        'bg': '#F5F1E8', 'bg_dark': '#ECE6D9', 'card': '#FFFFFF',
        'sidebar': '#EFEAE0', 'ink': '#2E2A24', 'ink_light': '#5C564C',
        'ink_muted': '#8A8375', 'border': '#DDD5C6', 'border_medium': '#C9BFAC',
        'accent': '#C0392B', 'accent_hover': '#A93226', 'accent_fg': '#FFFFFF',
        'teal': '#1F6F6B', 'success': '#2E7D32', 'warning': '#B26A00', 'error': '#C62828',
    },
    'dark': {
        'name': '暗夜', 'desc': '深色低光 · 护眼',
        'bg': '#1E1F24', 'bg_dark': '#17181C', 'card': '#282A31',
        'sidebar': '#17181C', 'ink': '#E8E6E1', 'ink_light': '#B8B5AD',
        'ink_muted': '#7E7B73', 'border': '#3A3D45', 'border_medium': '#4A4E58',
        'accent': '#E07B6A', 'accent_hover': '#D06553', 'accent_fg': '#1E1F24',
        'teal': '#4FB3AD', 'success': '#66BB6A', 'warning': '#FFB74D', 'error': '#EF5350',
    },
    'ink': {
        'name': '水墨', 'desc': '黑白灰 · 极简',
        'bg': '#F4F4F4', 'bg_dark': '#EAEAEA', 'card': '#FFFFFF',
        'sidebar': '#EDEDED', 'ink': '#1A1A1A', 'ink_light': '#4D4D4D',
        'ink_muted': '#8C8C8C', 'border': '#DCDCDC', 'border_medium': '#C4C4C4',
        'accent': '#1A1A1A', 'accent_hover': '#333333', 'accent_fg': '#FFFFFF',
        'teal': '#555555', 'success': '#2E7D32', 'warning': '#B26A00', 'error': '#C62828',
    },
    'teal': {
        'name': '青碧', 'desc': '青绿清爽 · 现代',
        'bg': '#F0F5F4', 'bg_dark': '#E3EDEB', 'card': '#FFFFFF',
        'sidebar': '#E8F0EF', 'ink': '#1F3331', 'ink_light': '#48605D',
        'ink_muted': '#7E9490', 'border': '#D2E0DE', 'border_medium': '#B5CCc9',
        'accent': '#0F766E', 'accent_hover': '#0B5F58', 'accent_fg': '#FFFFFF',
        'teal': '#0F766E', 'success': '#2E7D32', 'warning': '#B26A00', 'error': '#C62828',
    },
}


def settings():
    return QSettings("DocFormatPro", "DocFormatPro")


def current_theme_id():
    tid = settings().value("theme", "paper")
    return tid if tid in THEMES else 'paper'


def save_theme_id(tid):
    settings().setValue("theme", tid)


def build_qss(tid):
    c = THEMES.get(tid, THEMES['paper'])
    return """
QMainWindow, QWidget#Root {{ background: {bg}; }}
QWidget {{ color: {ink}; font-size: 13px; }}

/* ---- 侧边栏 ---- */
QFrame#Sidebar {{ background: {sidebar}; border-right: 1px solid {border}; }}
QLabel#Brand {{ font-size: 17px; font-weight: 700; color: {ink}; }}
QLabel#BrandAccent {{ font-size: 17px; font-weight: 700; color: {accent}; }}
QLabel#Version {{ color: {ink_muted}; font-size: 11px; }}
QPushButton[navBtn="true"] {{
    text-align: left; padding: 9px 14px; border: none; border-radius: 8px;
    color: {ink_light}; background: transparent; font-size: 13px;
}}
QPushButton[navBtn="true"]:hover {{ background: {bg_dark}; color: {ink}; }}
QPushButton[navBtn="true"]:checked {{ background: {accent}; color: {accent_fg}; font-weight: 600; }}

/* ---- 卡片 ---- */
QFrame[card="true"] {{ background: {card}; border: 1px solid {border}; border-radius: 12px; }}
QLabel[h1="true"] {{ font-size: 19px; font-weight: 700; color: {ink}; }}
QLabel[muted="true"] {{ color: {ink_muted}; font-size: 12px; }}
QLabel[sub="true"] {{ color: {ink_light}; font-size: 12px; }}
QLabel[sectionTitle="true"] {{ font-size: 13px; font-weight: 600; color: {ink}; }}

/* ---- 拖拽区 ---- */
QFrame#DropZone {{
    background: {bg}; border: 2px dashed {border_medium}; border-radius: 12px;
}}
QFrame#DropZone[dragOver="true"] {{ border-color: {accent}; background: {bg_dark}; }}
QLabel#DropTitle {{ font-size: 14px; font-weight: 600; color: {ink}; background: transparent; }}
QLabel#DropHint {{ color: {ink_muted}; font-size: 11px; background: transparent; }}

/* ---- 按钮 ---- */
QPushButton {{
    background: {card}; border: 1px solid {border_medium}; border-radius: 8px;
    padding: 7px 16px; color: {ink}; font-size: 13px;
}}
QPushButton:hover {{ background: {bg_dark}; }}
QPushButton:disabled {{ color: {ink_muted}; background: {bg}; border-color: {border}; }}
QPushButton[primary="true"] {{
    background: {accent}; color: {accent_fg}; border: none; font-weight: 600;
    padding: 10px 26px; font-size: 14px;
}}
QPushButton[primary="true"]:hover {{ background: {accent_hover}; }}
QPushButton[primary="true"]:disabled {{ background: {border_medium}; color: {card}; }}
QPushButton[danger="true"] {{ color: {error}; border-color: {error}; }}
QPushButton[flat="true"] {{ border: none; background: transparent; color: {teal}; }}
QPushButton[flat="true"]:hover {{ color: {accent}; background: transparent; }}

/* ---- 表单 ---- */
QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox {{
    background: {bg}; border: 1px solid {border_medium}; border-radius: 7px;
    padding: 5px 9px; color: {ink}; min-height: 20px;
}}
QComboBox:focus, QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {{ border-color: {accent}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background: {card}; border: 1px solid {border}; color: {ink};
    selection-background-color: {accent}; selection-color: {accent_fg};
}}
QCheckBox, QRadioButton {{ color: {ink}; spacing: 7px; }}
QRadioButton::indicator, QCheckBox::indicator {{ width: 16px; height: 16px; }}
QPlainTextEdit, QTextEdit {{
    background: {card}; border: 1px solid {border_medium}; border-radius: 8px; color: {ink};
    padding: 6px;
}}

/* ---- 进度条 ---- */
QProgressBar {{
    background: {bg_dark}; border: none; border-radius: 6px; height: 12px;
    text-align: center; color: {ink}; font-size: 10px;
}}
QProgressBar::chunk {{ background: {accent}; border-radius: 6px; }}

/* ---- 列表 ---- */
QListWidget {{
    background: {card}; border: 1px solid {border}; border-radius: 8px; padding: 4px;
}}
QListWidget::item {{ padding: 6px 8px; border-radius: 6px; color: {ink}; }}
QListWidget::item:selected {{ background: {bg_dark}; color: {ink}; }}

/* ---- 滚动区 ---- */
QScrollArea {{ border: none; background: transparent; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QScrollArea > QWidget > QScrollBar {{ background: transparent; }}
QScrollBar:vertical {{ background: transparent; width: 9px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {border_medium}; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {ink_muted}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 9px; }}
QScrollBar::handle:horizontal {{ background: {border_medium}; border-radius: 4px; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

/* ---- 折叠分组 ---- */
QToolButton[collapsibleHeader="true"] {{
    background: transparent; border: none; font-size: 13px; font-weight: 600;
    color: {ink}; padding: 4px; text-align: left;
}}

/* ---- 徽章 ---- */
QLabel[badge="true"] {{
    background: {bg_dark}; color: {ink_light}; border-radius: 9px;
    padding: 3px 10px; font-size: 11px;
}}
QLabel[badgeAccent="true"] {{
    background: {accent}; color: {accent_fg}; border-radius: 9px;
    padding: 3px 10px; font-size: 11px;
}}

/* ---- 文件列表状态与移除按钮 ---- */
QLabel[statusLevel="processing"] {{ color: {warning}; font-weight: 600; background: transparent; }}
QLabel[statusLevel="ok"] {{ color: {success}; font-weight: 600; background: transparent; }}
QLabel[statusLevel="fail"] {{ color: {error}; font-weight: 600; background: transparent; }}
QPushButton[removeBtn="true"] {{
    border: none; background: transparent; color: {ink_muted};
    font-size: 14px; padding: 2px 6px; border-radius: 6px;
}}
QPushButton[removeBtn="true"]:hover {{ color: {error}; background: {bg_dark}; }}

/* ---- 状态栏 ---- */
QFrame#StatusBar {{ background: {sidebar}; border-top: 1px solid {border}; }}

/* ---- 日志 ---- */
QTextEdit#LogView {{ font-family: Consolas, "Courier New", monospace; font-size: 12px; }}

QMessageBox {{ background: {card}; }}
QToolTip {{ background: {ink}; color: {card}; border: none; padding: 5px 8px; }}
""".format(**c)
