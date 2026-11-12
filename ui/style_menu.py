"""「배경색·글자·줄간격」 메뉴 — 업무창 · 새 메모장 · AI 대화 창이 함께 쓴다.

build(menu, ...)에 지금 값과 바꿀 때 부를 함수를 넘기면, 배경색 견본 · 글꼴 · 글자 크기 · 줄 간격을 만든다.
"""
import qtawesome as qta
from PyQt5.QtCore import QPointF, QRectF, Qt
from PyQt5.QtGui import QColor, QPainter, QPen, QPolygonF
from PyQt5.QtWidgets import QAbstractButton, QGridLayout, QWidget, QWidgetAction

from . import menu_style
from .common import NOTE_FONTS, note_font, note_font_available, refresh_installed_fonts
from .theme import PALETTE, PALETTE_NAMES, accent_for_note, note_color, tc

TITLE = '배경색·글자·줄간격'
SIZES = (('기본', 1.0), ('크게', 1.1), ('더 크게', 1.2))
SAME_AS_NOTE = '='                 # AI 창: 「업무창과 같게」


class Swatch(QAbstractButton):
    """배경색 견본: 둥근 네모 + 옅은 테두리. 고른 색은 그 색의 진한 색으로 바깥 테두리와 ✓."""
    SIZE, PAD = 26, 3

    def __init__(self, color, chosen=False, parent=None):
        super().__init__(parent)
        self.color, self.chosen = color, chosen
        side = self.SIZE + 2 * self.PAD
        self.setFixedSize(side, side)
        self.setCursor(Qt.PointingHandCursor)
        self.setAttribute(Qt.WA_Hover, True)
        if color == SAME_AS_NOTE:
            self.setToolTip('업무창과 같은 색')
        else:
            self.setToolTip(PALETTE_NAMES.get(color.upper(), ''))

    def _fill(self):
        return QColor(tc('#FFFFFF') if self.color == SAME_AS_NOTE else note_color(self.color))

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        pad, size = self.PAD, self.SIZE
        box = QRectF(pad + 0.5, pad + 0.5, size - 1, size - 1)
        fill = self._fill()
        accent = QColor(accent_for_note(fill.name()))
        p.setPen(QPen(QColor(0, 0, 0, 40 if fill.lightness() > 245 else 26), 1))
        p.setBrush(fill)
        p.drawRoundedRect(box, 7, 7)
        if self.chosen or self.underMouse():
            ring = QColor(accent) if self.chosen else QColor(0, 0, 0, 60)
            p.setPen(QPen(ring, 2 if self.chosen else 1.5))
            p.setBrush(Qt.NoBrush)
            p.drawRoundedRect(QRectF(1, 1, size + 2 * pad - 2, size + 2 * pad - 2), 9, 9)
        if self.color == SAME_AS_NOTE:
            icon = qta.icon('ph.link-bold', color=QColor('#64748B'))
            r = int(size * 0.55)
            icon.paint(p, int(pad + (size - r) / 2), int(pad + (size - r) / 2), r, r)
        elif self.chosen:
            p.setPen(QPen(accent, 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            x, y = pad, pad
            p.drawPolyline(QPolygonF([QPointF(x + 8, y + 13.5), QPointF(x + 11.5, y + 17), QPointF(x + 18, y + 9.5)]))
        p.end()


def color_name(color):
    if color == SAME_AS_NOTE:
        return '업무창과 같게'
    return PALETTE_NAMES.get((color or '').upper(), '')


def color_swatches(menu, current, on_pick, same_as_note=False):
    """배경색 견본(둥근 네모). same_as_note면 흰색 다음에 「업무창과 같게」(연결 아이콘)."""
    w = QWidget()
    w.setStyleSheet('background: transparent;')
    grid = QGridLayout(w)
    grid.setContentsMargins(10, 2, 10, 8)
    grid.setHorizontalSpacing(4)
    grid.setVerticalSpacing(4)
    cur = (current or '').upper()
    if same_as_note:                     # AI 창: 흰색(처음 값) · 업무창과 같게 · 메모 색들
        cur = cur or '#FFFFFF'
        items = ['#FFFFFF', SAME_AS_NOTE] + [c for c in PALETTE if c.upper() != '#FFFFFF']
    else:
        items = list(PALETTE)
    for i, color in enumerate(items):
        cb = Swatch(color, chosen=color.upper() == cur)
        cb.clicked.connect(lambda _, c=color: (on_pick(c), menu.close()))
        grid.addWidget(cb, i // 6, i % 6)
    return w


def build(menu, *, color, on_color, font_key, on_font, on_font_missing, scale, on_scale,
          gaps, gap, on_gap, same_as_note=False, icon='ph.palette'):
    """menu 아래에 「배경색·글자·줄간격」 하위 메뉴를 만든다.
    gaps: [(글, 값)] — 줄 간격 단계 (업무창은 업무 사이 간격, 메모장·AI 창은 글 줄 높이)."""
    sub = menu_style.add_submenu(menu, TITLE, icon=icon)
    name = color_name(color if (color or not same_as_note) else '#FFFFFF')
    menu_style.add_header(sub, f'배경색 · {name}' if name else '배경색')
    wa = QWidgetAction(sub)
    wa.setDefaultWidget(color_swatches(menu, color, on_color, same_as_note))
    sub.addAction(wa)
    sub.addSeparator()
    refresh_installed_fonts()
    fm = menu_style.add_submenu(sub, '글꼴', icon='ph.text-t')
    fm.setMinimumWidth(200)
    for key, label in NOTE_FONTS:
        ok = note_font_available(key)
        a = menu_style.add_action(fm, label if ok else f'{label} (설치 필요)',
                                  (lambda k=key: on_font(k)) if ok else (lambda k=key: on_font_missing(k)),
                                  checked=font_key == key)
        a.setFont(note_font(11, key))                 # 글꼴 이름을 그 글꼴로 미리 보기
        if not ok:
            a.setToolTip('이 PC에 설치되어 있지 않아요 — 누르면 받는 곳과 설치 방법을 알려 드려요')
    sm = menu_style.add_submenu(sub, '글자 크기', icon='ph.text-aa')
    sm.setMinimumWidth(110)
    for label, val in SIZES:
        menu_style.add_action(sm, label, lambda v=val: on_scale(v), checked=abs(scale - val) < 0.01)
    gm = menu_style.add_submenu(sub, '줄 간격', icon='ph.rows')
    gm.setMinimumWidth(110)
    for label, val in gaps:
        menu_style.add_action(gm, label, lambda v=val: on_gap(v), checked=gap == val)
    return sub
