"""「배경색·글자·줄간격」 메뉴 — 업무창 · 새 메모장 · AI 대화 창이 함께 쓴다.

build(menu, ...)에 지금 값과 바꿀 때 부를 함수를 넘기면, 배경색 견본 · 글꼴 · 글자 크기 · 줄 간격을 만든다.
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QGridLayout, QPushButton, QWidget, QWidgetAction

from . import menu_style
from .common import NOTE_FONTS, note_font, note_font_available, refresh_installed_fonts
from .theme import PALETTE, note_color, tc

TITLE = '배경색·글자·줄간격'
SIZES = (('기본', 1.0), ('크게', 1.1), ('더 크게', 1.2))
SAME_AS_NOTE = ''                  # AI 창: 「업무창과 같게」


def color_swatches(menu, current, on_pick, same_as_note=False):
    """배경색 견본(동그라미). 지금 색은 바깥 고리. same_as_note면 맨 앞에 「업무창과 같게」."""
    w = QWidget()
    w.setStyleSheet('background: transparent;')
    grid = QGridLayout(w)
    grid.setContentsMargins(12, 4, 12, 8)
    grid.setHorizontalSpacing(8)
    grid.setVerticalSpacing(8)
    cur = (current or '').upper()
    items = ([SAME_AS_NOTE] if same_as_note else []) + list(PALETTE)
    for i, color in enumerate(items):
        cb = QPushButton()
        cb.setFixedSize(24, 24)
        cb.setCursor(Qt.PointingHandCursor)
        chosen = color.upper() == cur
        ring = f'2px solid {menu_style.TEXT.name()}' if chosen else f'1px solid {tc("rgba(0,0,0,0.12)")}'
        if color == SAME_AS_NOTE:
            cb.setToolTip('업무창과 같은 색')
            cb.setText('=')
            bg = tc('#FFFFFF')
        else:
            cb.setToolTip('이 색으로')
            bg = note_color(color)
        cb.setStyleSheet(f"""
            QPushButton {{ background: {bg}; border-radius: 12px; border: {ring}; font-weight: 800; padding: 0; }}
            QPushButton:hover {{ border: 2px solid {tc("rgba(0,0,0,0.35)")}; }}
        """)
        cb.clicked.connect(lambda _, c=color: (on_pick(c), menu.close()))
        grid.addWidget(cb, i // 6, i % 6)
    return w


def build(menu, *, color, on_color, font_key, on_font, on_font_missing, scale, on_scale,
          gaps, gap, on_gap, same_as_note=False, icon='ph.palette'):
    """menu 아래에 「배경색·글자·줄간격」 하위 메뉴를 만든다.
    gaps: [(글, 값)] — 줄 간격 단계 (업무창은 업무 사이 간격, 메모장·AI 창은 글 줄 높이)."""
    sub = menu_style.add_submenu(menu, TITLE, icon=icon)
    menu_style.add_header(sub, '배경색')
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
