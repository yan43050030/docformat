# -*- coding: utf-8 -*-
"""套打版面画布：按真实 A4 比例画整张纸，可直接拖着字挪位置。

原来的预览是拿富文本拼的表格，Qt 那边按自己的规矩排版，横竖比例都不是
A4，字也小得看不清。这里改成自己画：纸就是 21×29.7cm 的比例，每个字排
在它真实的厘米坐标上，一眼能看出"印出来是什么样"。

能拖是顺带的好处——位置既然是按厘米算出来的，反过来把鼠标位置换算成
厘米就是新的目标位置。拖黑字改的是**那一个字段**的横向位置（引擎只支持
逐字段调横向）；在空白处拖动的是**整张纸**，横竖都能挪，用来补打印机
走纸和纸张裁切的整体偏差。
"""
from PyQt5.QtCore import QRectF, QSize, Qt, pyqtSignal
from PyQt5.QtGui import (QBrush, QColor, QFont, QFontMetricsF, QPainter, QPen,
                         QPixmap)
from PyQt5.QtWidgets import QWidget

from scripts.overprint import PT_PER_CM, _seg_width_cm

# 预印内容（纸上已有、不打印）画成浅灰；要打印的画成黑色
INK_PRE = QColor('#B9B9B9')
INK_PRINT = QColor('#111111')
RULE = QColor('#D8D8D8')
WARN = QColor(253, 236, 234)     # 放不下的格子铺这个淡红
PICK = QColor('#1F7A4D')


class OverprintCanvas(QWidget):
    """按厘米画版面；黑字可拖。"""

    fieldMoved = pyqtSignal(str, float)      # 字段名, 新的距纸左边 cm
    sheetMoved = pyqtSignal(float, float)    # 整张纸 dx, dy（cm）
    fieldPicked = pyqtSignal(str)            # 选中了哪个字段（'' = 没选）

    def __init__(self, parent=None):
        super(OverprintCanvas, self).__init__(parent)
        self.setMouseTracking(True)
        self.setMinimumSize(240, 340)
        self._plan = None
        self._bg = None              # 套头底图 QPixmap
        self._zoom = 0.0             # 0 = 自适应窗口
        self._items = []             # 可拖的黑字：[(字段, QRectF(cm), 起点x)]
        self._hot = None             # 鼠标底下的那个
        self._drag = None            # 正在拖的
        self._sheet = (0.0, 0.0)     # 整体平移（cm）
        self._sheet_drag = None
        self.setCursor(Qt.ArrowCursor)

    # ---------- 数据 ----------
    def set_plan(self, plan, shift=(0.0, 0.0)):
        self._plan = plan
        self._sheet = (float(shift[0] or 0), float(shift[1] or 0))
        self.update()

    def set_background(self, pixmap):
        self._bg = pixmap
        self.update()

    def set_zoom(self, z):
        """z=0 表示自适应窗口"""
        self._zoom = float(z or 0)
        self.updateGeometry()
        self.update()

    def page_cm(self):
        if not self._plan:
            return 21.0, 29.7
        p = self._plan['page']
        return p['width_cm'], p['height_cm']

    # ---------- 坐标换算 ----------
    def _scale(self):
        """一厘米画多少像素"""
        w_cm, h_cm = self.page_cm()
        if self._zoom:
            return self._zoom
        # 自适应：整张纸完整放进窗口，留一点边
        avail_w = max(40, self.width() - 24)
        avail_h = max(40, self.height() - 24)
        return max(4.0, min(avail_w / w_cm, avail_h / h_cm))

    def _page_rect(self):
        s = self._scale()
        w_cm, h_cm = self.page_cm()
        w, h = w_cm * s, h_cm * s
        x = max(12, (self.width() - w) / 2.0)
        y = 12
        return QRectF(x, y, w, h)

    def minimumSizeHint(self):
        # 指定了缩放就按纸张实际尺寸要地方（外面套滚动区）；
        # 自适应模式下只要个下限，剩下的交给窗口
        w_cm, h_cm = self.page_cm()
        if self._zoom:
            return QSize(int(w_cm * self._zoom) + 24, int(h_cm * self._zoom) + 24)
        return QSize(240, 340)

    sizeHint = minimumSizeHint

    # ---------- 排版 ----------
    @staticmethod
    def _seg_w(seg):
        # 宽度算法与排版侧共用一套，含字距——模板里栏目名是收着排的
        return _seg_width_cm(seg, 14 / PT_PER_CM)

    @classmethod
    def _lines_of(cls, segs):
        """按换行片段切成 [[seg,...], ...]"""
        lines, cur = [], []
        for s in segs:
            if s.get('text') == '\n':
                lines.append(cur)
                cur = []
            else:
                cur.append(s)
        lines.append(cur)
        return lines

    def _layout(self):
        """把 plan 摊成一串可画的元素，坐标一律是厘米（距纸左边/上边）。

        返回 (画的东西, 可拖的黑字)。两者分开：可拖的要单独记矩形做命中判断。
        """
        draw, items = [], []
        if not self._plan:
            return draw, items
        pg = self._plan['page']
        left, right = pg['left_cm'], pg['right_cm']
        content_w = pg['width_cm'] - left - right
        dx, dy = self._sheet

        def emit_line(line, x0, top, line_cm, movable=True):
            x = x0
            for s in line:
                w = self._seg_w(s)
                txt = s.get('text', '')
                if s.get('pad_cm') is None and txt.strip():
                    draw.append({'x': x + dx, 'y': top + dy, 'text': txt,
                                 'pt': s.get('pt') or 14,
                                 'white': bool(s.get('white')),
                                 'line_cm': line_cm})
                    # 居中排的内容（标题）不给拖：它的位置由格子决定，
                    # 拖出来的横坐标存回去也不会生效，白让人试
                    if s.get('field') and movable:
                        items.append({'field': s['field'],
                                      'rect': QRectF(x + dx, top + dy, w, line_cm),
                                      'x_cm': x + dx})
                x += w

        for b in self._plan['blocks']:
            if b['kind'] == 'para':
                line_cm = b.get('line_cm') or 0.5
                for i, line in enumerate(self._lines_of(b.get('segs') or [])):
                    w = sum(self._seg_w(s) for s in line)
                    if b.get('align') == 'center':
                        x0 = left + (content_w - w) / 2.0
                    elif b.get('align') == 'right':
                        x0 = left + content_w - w
                    else:
                        x0 = left
                    emit_line(line, x0, b['top_cm'] + i * line_cm, line_cm)
            else:
                bd = b.get('borders') or {}
                for r in b['rows']:
                    cx = left
                    for ci, c in enumerate(r['cells']):
                        cw = c.get('width_cm') or 0
                        if c.get('overflow'):
                            # 缩到最小仍放不下：整格铺一层淡红，一眼看得见
                            draw.append({'warn': True, 'x': cx + dx,
                                         'y': r['top_cm'] + dy,
                                         'w': cw, 'h': r['height_cm']})
                        # 格线：模板里是白线（不显影），预览画成淡灰当参照
                        if ci and bd.get('insideV', 'none') != 'none':
                            draw.append({'rule': 'v', 'x': cx + dx,
                                         'y': r['top_cm'] + dy,
                                         'h': r['height_cm']})
                        segs = c.get('segs') or []
                        pt = c.get('font_pt') or 14
                        line_cm = pt / PT_PER_CM * 1.15
                        centered = bool(c.get('is_title'))
                        for i, line in enumerate(self._lines_of(segs)):
                            lw = sum(self._seg_w(s) for s in line)
                            x0 = cx + ((cw - lw) / 2.0 if centered and lw < cw
                                       else 0.0)
                            emit_line(line, x0, r['top_cm'] + i * line_cm,
                                      line_cm, movable=not centered)
                        cx += cw
                    if bd.get('insideH', 'none') != 'none' and not any(
                            c.get('vmerge_cont') for c in r['cells']):
                        draw.append({'rule': 'h', 'x': left + dx,
                                     'y': r['top_cm'] + dy,
                                     'w': pg['width_cm'] - left - right})
        return draw, items

    # ---------- 绘制 ----------
    def paintEvent(self, _e):
        qp = QPainter(self)
        qp.setRenderHint(QPainter.Antialiasing, True)
        qp.setRenderHint(QPainter.TextAntialiasing, True)
        rect = self._page_rect()
        s = self._scale()

        # 纸
        qp.fillRect(self.rect(), QColor('#8E8E8E'))
        qp.fillRect(rect, QColor('#FFFFFF'))
        if self._bg is not None and not self._bg.isNull():
            qp.setOpacity(0.9)
            qp.drawPixmap(rect.toRect(), self._bg)
            qp.setOpacity(1.0)
        qp.setPen(QPen(QColor('#555555'), 1))
        qp.drawRect(rect)

        if not self._plan:
            qp.setPen(QColor('#666666'))
            qp.drawText(rect, Qt.AlignCenter, '填入内容后这里显示版面')
            qp.end()
            return

        draw, self._items = self._layout()

        def px(cm_x, cm_y):
            return rect.left() + cm_x * s, rect.top() + cm_y * s

        for d in draw:
            if d.get('warn'):
                x, y = px(d['x'], d['y'])
                qp.fillRect(QRectF(x, y, d['w'] * s, d['h'] * s), WARN)
                continue
            if d.get('rule'):
                qp.setPen(QPen(RULE, 1))
                x, y = px(d['x'], d['y'])
                if d['rule'] == 'h':
                    qp.drawLine(int(x), int(y), int(x + d['w'] * s), int(y))
                else:
                    qp.drawLine(int(x), int(y), int(x), int(y + d['h'] * s))
                continue
            f = QFont()
            f.setPixelSize(max(1, int(round(d['pt'] / PT_PER_CM * s))))
            qp.setFont(f)
            qp.setPen(INK_PRE if d['white'] else INK_PRINT)
            x, y = px(d['x'], d['y'])
            fm = QFontMetricsF(f)
            # 厘米坐标给的是行盒顶端，文字要按基线画
            qp.drawText(int(x), int(y + (d['line_cm'] * s + fm.ascent()
                                         - fm.descent()) / 2.0), d['text'])

        # 可拖的黑字：鼠标扫过时描一个框，提示"这个能拖"
        for it in self._items:
            if it is not self._hot and it is not self._drag:
                continue
            r = it['rect']
            x, y = px(r.left(), r.top())
            qp.setPen(QPen(PICK, 1, Qt.DashLine))
            qp.setBrush(QBrush(Qt.NoBrush))
            qp.drawRect(QRectF(x - 2, y - 1, r.width() * s + 4, r.height() * s + 2))
        qp.end()

    # ---------- 交互 ----------
    def _cm_at(self, pos):
        rect, s = self._page_rect(), self._scale()
        return (pos.x() - rect.left()) / s, (pos.y() - rect.top()) / s

    def _hit(self, pos):
        cx, cy = self._cm_at(pos)
        for it in self._items:
            r = it['rect']
            if r.adjusted(-0.1, -0.05, 0.1, 0.05).contains(cx, cy):
                return it
        return None

    def mouseMoveEvent(self, e):
        if self._drag is not None:
            cx, _cy = self._cm_at(e.pos())
            new_x = max(0.0, round(cx - self._drag['grab'], 2))
            self._drag['rect'].moveLeft(new_x)
            self.update()
            return
        if self._sheet_drag is not None:
            cx, cy = self._cm_at(e.pos())
            ox, oy, sx, sy = self._sheet_drag
            self._sheet = (round(sx + cx - ox, 2), round(sy + cy - oy, 2))
            self.update()
            return
        hot = self._hit(e.pos())
        if hot is not self._hot:
            self._hot = hot
            self.setCursor(Qt.SizeHorCursor if hot else Qt.OpenHandCursor)
            self.setToolTip('拖动可改「{}」的横向位置'.format(hot['field'])
                            if hot else '在空白处拖动可整体平移这张纸')
            self.update()

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        hit = self._hit(e.pos())
        cx, cy = self._cm_at(e.pos())
        if hit is not None:
            hit['grab'] = cx - hit['rect'].left()
            self._drag = hit
            self.fieldPicked.emit(hit['field'])
            self.setCursor(Qt.ClosedHandCursor)
        else:
            self._sheet_drag = (cx, cy, self._sheet[0], self._sheet[1])
            self.fieldPicked.emit('')
            self.setCursor(Qt.ClosedHandCursor)

    def mouseReleaseEvent(self, e):
        if self._drag is not None:
            field = self._drag['field']
            x_cm = round(self._drag['rect'].left(), 2)
            self._drag = None
            self.setCursor(Qt.ArrowCursor)
            # 拖出来的是"含整体平移之后"的位置，存回去要减掉平移量，
            # 否则整体一挪、逐字段的位置也跟着叠一遍
            self.fieldMoved.emit(field, round(x_cm - self._sheet[0], 2))
        elif self._sheet_drag is not None:
            self._sheet_drag = None
            self.setCursor(Qt.ArrowCursor)
            self.sheetMoved.emit(self._sheet[0], self._sheet[1])
        self.update()

    def text_dump(self):
        """按绘制顺序把纸上的字连起来——自动化测试拿它核版面顺序"""
        draw, _items = self._layout()
        return ''.join(d.get('text', '') for d in draw if not d.get('rule')
                       and not d.get('warn'))

    def fields(self):
        """当前画面上可拖的字段 → 距纸左边 cm"""
        _draw, items = self._layout()
        return {it['field']: round(it['x_cm'], 2) for it in items}

    def leaveEvent(self, _e):
        self._hot = None
        self.update()
