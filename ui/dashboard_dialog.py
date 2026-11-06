"""ui/dashboard_dialog.py — 「오늘 한눈에 보기」 창.

지난 마감 / 오늘 마감 / 이번 주 / 공문 업무 개수, 급한 업무 목록, 예산 집행률을 한 장에 보여준다.
업무를 누르면 해당 메모 창에서 그 업무를 강조한다.
"""

from datetime import date

from PyQt5.QtCore import Qt, QRectF, pyqtSignal
from PyQt5.QtGui import QColor, QPainter
from PyQt5.QtWidgets import (
    QCheckBox, QDialog, QFrame, QGridLayout, QHBoxLayout, QLabel, QPushButton, QSizePolicy,
    QVBoxLayout, QWidget,
)

import dashboard
from db import get_budget_items, get_open_tasks, get_setting, set_setting
from .common import pr_font, px
from .theme import T

SHOW_ON_START_KEY = 'dashboard_on_start'
LAST_SHOWN_KEY = 'dashboard_last_shown'
_MAX_ROWS = 5

_STYLE = f"""
    QDialog {{ background: {T.BG_WARM}; }}
    QLabel {{ color: {T.TEXT_WARM}; background: transparent; }}
    QLabel#Title {{ color: {T.INK}; font-size: 15pt; font-weight: 700; }}
    QLabel#Date {{ color: {T.TEXT_WARM_MUTED}; font-size: 10pt; }}
    QLabel#Headline {{ color: {T.INK}; font-size: 10.5pt; }}
    QLabel#Section {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; font-weight: 700; }}
    QLabel#Muted {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QFrame#Card {{ background: {T.SURFACE}; border: 1px solid {T.BORDER_WARM}; border-radius: 8px; }}
    QLabel#TileLabel {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QLabel#TileValue {{ color: {T.INK}; font-size: 19pt; font-weight: 600; }}
    QPushButton#TaskRow {{
        background: transparent; border: none; border-radius: 5px; text-align: left; padding: 0;
    }}
    QPushButton#TaskRow:hover {{ background: {T.BG_WARM_ALT}; }}
    QPushButton {{
        background: {T.SURFACE_WARM}; color: {T.TEXT_ON_NOTE}; border: none; border-radius: 5px;
        padding: 6px 14px;
    }}
    QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
"""


def should_show_on_start(today: date = None) -> bool:
    """하루 한 번만 — 설정이 켜져 있고 오늘 아직 안 보여줬으면 True."""
    today = today or date.today()
    return get_setting(SHOW_ON_START_KEY, '1') != '0' and get_setting(LAST_SHOWN_KEY, '') != today.isoformat()


class _Meter(QWidget):
    """집행률 막대. 바탕은 같은 노랑 계열의 옅은 단계, 100% 넘으면 위험 색."""

    def __init__(self, pct: float, parent=None):
        super().__init__(parent)
        self._pct = max(0.0, pct)
        self.setFixedHeight(8)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setAccessibleName(f'예산 집행률 {pct:.1f}%')

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        r = QRectF(self.rect())
        p.setBrush(QColor(T.ACCENT_SOFT))
        p.drawRoundedRect(r, 4, 4)
        w = r.width() * min(self._pct, 100.0) / 100.0
        if w > 0:
            p.setBrush(QColor(T.DANGER if self._pct > 100 else T.ACCENT))
            p.drawRoundedRect(QRectF(r.x(), r.y(), max(w, 8.0), r.height()), 4, 4)
        p.end()


def _tile(label: str, value: int, marker: str = '') -> QFrame:
    """숫자 카드. 값은 글자색(INK) 그대로, 상태 색은 라벨 앞 점(●)에만 — 색만으로 뜻을 전하지 않는다."""
    card = QFrame()
    card.setObjectName('Card')
    lay = QVBoxLayout(card)
    lay.setContentsMargins(12, 9, 12, 9)
    lay.setSpacing(0)
    lbl = QLabel(f'<span style="color:{marker}">●</span>&nbsp;{label}' if marker else label)
    lbl.setObjectName('TileLabel')
    val = QLabel(f'{value:,}')
    val.setObjectName('TileValue')
    lay.addWidget(lbl)
    lay.addWidget(val)
    card.setAccessibleName(f'{label} {value}건')
    return card


class DashboardDialog(QDialog):
    jump_requested = pyqtSignal(dict)   # 업무 클릭 → 해당 메모 창에서 강조
    open_budget = pyqtSignal()

    def __init__(self, parent=None, summary: dict = None):
        super().__init__(parent, Qt.Window | Qt.WindowStaysOnTopHint)
        self.setWindowTitle('오늘 한눈에 보기')
        self.setStyleSheet(_STYLE)
        self.setFont(pr_font(10))
        self.setFixedWidth(px(440))
        s = summary if summary is not None else dashboard.build_summary(get_open_tasks(), get_budget_items())
        self._build(s)
        set_setting(LAST_SHOWN_KEY, date.today().isoformat())
        self.resize(self.width(), self.sizeHint().height())

    def _build(self, s: dict):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 14)
        root.setSpacing(12)

        head = QVBoxLayout()
        head.setSpacing(2)
        top = QHBoxLayout()
        title = QLabel('오늘')
        title.setObjectName('Title')
        date_lbl = QLabel(s['date_label'])
        date_lbl.setObjectName('Date')
        top.addWidget(title)
        top.addWidget(date_lbl, 0, Qt.AlignBottom)
        top.addStretch()
        head.addLayout(top)
        hl = QLabel(s['headline'])
        hl.setObjectName('Headline')
        hl.setWordWrap(True)
        head.addWidget(hl)
        root.addLayout(head)

        tiles = QGridLayout()
        tiles.setSpacing(8)
        tiles.addWidget(_tile('지난 마감', len(s['overdue']), T.DANGER if s['overdue'] else ''), 0, 0)
        tiles.addWidget(_tile('오늘 마감', len(s['today']), T.HIGHLIGHT if s['today'] else ''), 0, 1)
        tiles.addWidget(_tile('이번 주', len(s['week'])), 0, 2)
        tiles.addWidget(_tile('공문 업무', s['doc_open']), 0, 3)
        root.addLayout(tiles)

        sections = [('지난 마감', s['overdue']), ('오늘', s['today']),
                    ('이번 주', s['week']), ('중요 · 기한 없음', s['important'])]
        shown = [(n, items) for n, items in sections if items]
        if shown:
            card = QFrame()
            card.setObjectName('Card')
            lay = QVBoxLayout(card)
            lay.setContentsMargins(10, 8, 10, 8)
            lay.setSpacing(2)
            for i, (name, items) in enumerate(shown):
                if i:
                    lay.addSpacing(6)
                sec = QLabel(name)
                sec.setObjectName('Section')
                lay.addWidget(sec)
                for t in items[:_MAX_ROWS]:
                    lay.addWidget(self._task_row(t, overdue=(name == '지난 마감')))
                if len(items) > _MAX_ROWS:
                    more = QLabel(f'   외 {len(items) - _MAX_ROWS}건')
                    more.setObjectName('Muted')
                    lay.addWidget(more)
            root.addWidget(card)
        else:
            empty = QLabel('표시할 마감 업무가 없습니다. 메모 창에서 업무에 날짜를 넣으면 여기에 모입니다.')
            empty.setObjectName('Muted')
            empty.setWordWrap(True)
            root.addWidget(empty)

        b = s.get('budget')
        if b:
            card = QFrame()
            card.setObjectName('Card')
            lay = QVBoxLayout(card)
            lay.setContentsMargins(12, 10, 12, 10)
            lay.setSpacing(6)
            row = QHBoxLayout()
            lbl = QLabel('예산 집행률')
            lbl.setObjectName('TileLabel')
            val = QLabel(f"{b['pct']:.1f}%")
            val.setStyleSheet(f'color: {T.INK}; font-size: 13pt; font-weight: 600;')
            row.addWidget(lbl)
            row.addStretch()
            row.addWidget(val)
            lay.addLayout(row)
            lay.addWidget(_Meter(b['pct']))
            detail = QLabel(f"집행 {dashboard.format_won(b['spent'])} / 편성 {dashboard.format_won(b['allocated'])}"
                            f" · 잔액 {dashboard.format_won(b['remaining'])}")
            detail.setObjectName('Muted')
            lay.addWidget(detail)
            root.addWidget(card)

        foot = QHBoxLayout()
        self.chk_start = QCheckBox('하루 한 번, 앱을 켤 때 보기')
        self.chk_start.setChecked(get_setting(SHOW_ON_START_KEY, '1') != '0')
        self.chk_start.toggled.connect(lambda on: set_setting(SHOW_ON_START_KEY, '1' if on else '0'))
        foot.addWidget(self.chk_start)
        foot.addStretch()
        if b:
            btn_budget = QPushButton('예산 창')
            btn_budget.clicked.connect(self.open_budget.emit)
            foot.addWidget(btn_budget)
        btn_close = QPushButton('닫기')
        btn_close.setDefault(True)
        btn_close.clicked.connect(self.close)
        foot.addWidget(btn_close)
        root.addLayout(foot)

    def _task_row(self, t: dict, overdue=False) -> QPushButton:
        btn = QPushButton()
        btn.setObjectName('TaskRow')
        btn.setCursor(Qt.PointingHandCursor)
        lay = QHBoxLayout(btn)
        lay.setContentsMargins(6, 4, 6, 4)
        lay.setSpacing(8)
        dd = QLabel(t.get('dday') or '')
        dd.setFixedWidth(px(54))
        dd.setStyleSheet(f"color: {T.DANGER if overdue else T.TEXT_ON_NOTE}; font-weight: 700;")
        name = QLabel(('★ ' if t.get('priority') else '') + (t.get('name') or ''))
        name.setStyleSheet(f'color: {T.INK};')
        name.setTextInteractionFlags(Qt.NoTextInteraction)
        lay.addWidget(dd)
        lay.addWidget(name, 1)
        if t.get('time'):
            tm = QLabel(t['time'])
            tm.setObjectName('Muted')
            lay.addWidget(tm)
        btn.setMinimumHeight(28)
        btn.setToolTip('눌러서 메모 창에서 보기')
        btn.setAccessibleName(f"{t.get('dday', '')} {t.get('name', '')}")
        btn.clicked.connect(lambda: self.jump_requested.emit(t))
        return btn
