"""ui/dashboard_dialog.py — 「오늘 한눈에 보기」 창.

지난 마감 / 오늘 마감 / 이번 주 개수, 급한 업무 목록, 예산 집행률을 한 장에 보여준다.
업무를 누르면 해당 메모 창에서 그 업무를 강조한다.
"""

from datetime import date

from PyQt5.QtCore import Qt, QRectF, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QPainter
from PyQt5.QtWidgets import (
    QCheckBox, QDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy,
    QVBoxLayout, QWidget,
)

import dashboard
from db import get_budget_items, get_open_tasks, get_setting, set_setting
from .common import pr_font, px

SHOW_ON_START_KEY = 'dashboard_on_start'
LAST_SHOWN_KEY = 'dashboard_last_shown'
_MAX_ROWS = 5

# 이 창의 글꼴: 앱 공통 Pretendard (없는 PC는 맑은 고딕)
_FONT_FAMILIES = ['Pretendard GOV', 'Pretendard', 'Malgun Gothic']
_FONT_STACK = ', '.join(f"'{f}'" for f in _FONT_FAMILIES)


def _font(pt, weight=QFont.Normal) -> QFont:
    f = QFont(_FONT_FAMILIES[0], 10)
    f.setFamilies(_FONT_FAMILIES)
    f.setPointSizeF(pt)
    f.setWeight(weight)
    return f


def _level_color(pct: float) -> str:
    """집행률 색 (예산 창과 같은 기준): 100% 이상 빨강 · 80% 이상 주황 · 그 밖은 기본 색."""
    from .calendar_view import P
    return P.DANGER if pct >= 100 else ('#D97706' if pct >= 80 else P.PRIMARY)


def _style():
    from .calendar_view import P
    from .common import icon_png
    return f"""
    QDialog {{ background: {P.BG}; }}
    QWidget {{ font-family: {_FONT_STACK}; }}
    QLabel {{ color: {P.TEXT}; background: transparent; }}
    QLabel#Eyebrow {{ color: {P.MUTED}; font-size: 9.5pt; font-weight: 600; }}
    QLabel#Title {{ color: {P.TEXT}; font-size: 21pt; font-weight: 600; }}
    QLabel#Headline {{ color: {P.TEXT2}; font-size: 10.5pt; }}
    QLabel#Section {{ color: {P.MUTED}; font-size: 9pt; font-weight: 700; }}
    QLabel#Count {{ color: {P.FAINT}; font-size: 9pt; font-weight: 600; }}
    QLabel#Muted {{ color: {P.MUTED}; font-size: 9pt; }}
    QLabel#StatLabel {{ color: {P.MUTED}; font-size: 9pt; font-weight: 500; }}
    QFrame#Card {{ background: {P.BG}; border: 1px solid {P.LINE}; border-radius: 14px; }}
    QFrame#Rule {{ background: {P.LINE}; border: none; }}
    QPushButton#TaskRow {{ background: transparent; border: none; border-radius: 10px; text-align: left; padding: 0; }}
    QPushButton#TaskRow:hover {{ background: {P.HOVER}; }}
    QPushButton#Ghost {{
        background: transparent; color: {P.TEXT}; border: 1px solid {P.LINE}; border-radius: 8px; padding: 6px 14px;
        font-size: 10pt; font-weight: 600;
    }}
    QPushButton#Ghost:hover {{ background: {P.HOVER}; }}
    QPushButton#Primary {{
        background: {P.PRIMARY}; color: {P.PRIMARY_TEXT}; border: none; border-radius: 8px; padding: 6px 18px;
        font-size: 10pt; font-weight: 600;
    }}
    QPushButton#Primary:hover {{ background: {P.PRIMARY_HOVER}; }}
    QPushButton#Link {{ background: transparent; border: none; color: {P.TEXT2}; padding: 2px 4px;
        font-size: 9.5pt; font-weight: 600; }}
    QPushButton#Link:hover {{ color: {P.TEXT}; }}
    QCheckBox {{ color: {P.TEXT2}; spacing: 7px; font-size: 9.5pt; }}
    QCheckBox::indicator {{ width: 15px; height: 15px; border: 1px solid {P.FAINT}; border-radius: 4px; background: {P.BG}; }}
    QCheckBox::indicator:checked {{ background: {P.PRIMARY}; border-color: {P.PRIMARY};
        image: url({icon_png('ph.check-bold', P.PRIMARY_TEXT, 14)}); }}
"""


def should_show_on_start(today: date = None) -> bool:
    """하루 한 번만 — 설정이 켜져 있고 오늘 아직 안 보여줬으면 True."""
    today = today or date.today()
    return get_setting(SHOW_ON_START_KEY, '1') != '0' and get_setting(LAST_SHOWN_KEY, '') != today.isoformat()


class _Meter(QWidget):
    """집행률 막대: 옅은 바탕 위 막대 (80% 이상 주황, 100% 이상 빨강)."""

    def __init__(self, pct: float, parent=None):
        super().__init__(parent)
        self._pct = max(0.0, pct)
        self.setFixedHeight(8)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setAccessibleName(f'예산 집행률 {pct:.1f}%')

    def paintEvent(self, _):
        from .calendar_view import P
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        r = QRectF(self.rect())
        p.setBrush(QColor(P.CHIP))
        p.drawRoundedRect(r, 4, 4)
        w = r.width() * min(self._pct, 100.0) / 100.0
        if w > 0:
            p.setBrush(QColor(_level_color(self._pct)))
            p.drawRoundedRect(QRectF(r.x(), r.y(), max(w, 8.0), r.height()), 4, 4)
        p.end()


def _tile(label: str, value: int, color: str, tinted: bool) -> QFrame:
    """숫자 칸: 큰 숫자와 이름이 칸의 가로 가운데에 온다.
    tinted(지난 마감·오늘 마감이 1건 이상)면 그 색의 연한 바탕을 깔아 눈에 띄게 한다."""
    from .calendar_view import P
    from .theme import rgba
    w = QFrame()
    w.setObjectName('Tile')
    if tinted:
        bg, border = rgba(color, 0.10), rgba(color, 0.24)
    else:
        bg, border = P.BG, P.LINE
    w.setStyleSheet(f'QFrame#Tile {{ background: {bg}; border: 1px solid {border}; border-radius: 14px; }}')
    lay = QVBoxLayout(w)
    lay.setContentsMargins(6, 14, 6, 12)
    lay.setSpacing(1)
    val = QLabel(f'{value:,}')
    val.setFont(_font(24, QFont.DemiBold))
    val.setAlignment(Qt.AlignCenter)
    val.setStyleSheet(f'color: {color};')
    lbl = QLabel(label)
    lbl.setObjectName('StatLabel')
    lbl.setAlignment(Qt.AlignCenter)
    lay.addWidget(val)
    lay.addWidget(lbl)
    w.value_label, w.name_label = val, lbl
    w.setAccessibleName(f'{label} {value}건')
    return w


class DashboardDialog(QDialog):
    jump_requested = pyqtSignal(dict)   # 업무 클릭 → 해당 메모 창에서 강조
    open_budget = pyqtSignal()

    def __init__(self, parent=None, summary: dict = None):
        super().__init__(parent, Qt.Window | Qt.WindowStaysOnTopHint)
        self.setWindowTitle('오늘 한눈에 보기')
        self.setStyleSheet(_style())
        self.setFont(_font(10))
        self.setFixedWidth(px(480))
        self.tiles = []
        self.rows = []
        s = summary if summary is not None else dashboard.build_summary(get_open_tasks(), get_budget_items())
        self._build(s)
        set_setting(LAST_SHOWN_KEY, date.today().isoformat())
        self.resize(self.width(), self.sizeHint().height())

    def _build(self, s: dict):
        from .calendar_view import P
        root = QVBoxLayout(self)
        root.setContentsMargins(26, 24, 26, 18)
        root.setSpacing(18)

        # ── 머리말: 오늘 날짜 + 한 줄 요약 ──
        head = QVBoxLayout()
        head.setSpacing(3)
        eyebrow = QLabel('오늘 한눈에 보기')
        eyebrow.setObjectName('Eyebrow')
        import re
        title = QLabel(re.sub(r'\s*\((.)\)$', r' \1요일', s['date_label']))     # '11월 7일 (토)' → '11월 7일 토요일'
        title.setObjectName('Title')
        hl = QLabel(s['headline'])
        hl.setObjectName('Headline')
        hl.setWordWrap(True)
        head.addWidget(eyebrow)
        head.addWidget(title)
        head.addWidget(hl)
        root.addLayout(head)

        # ── 숫자 세 칸 (둥근 타일, 글자는 각 칸의 가운데) ──
        stats = QHBoxLayout()
        stats.setSpacing(10)
        items = [('지난 마감', len(s['overdue']), P.DANGER, True),
                 ('오늘 마감', len(s['today']), '#E8590C', True),
                 ('이번 주', len(s['week']), P.TEXT, False)]
        for label, value, color, can_tint in items:
            on = can_tint and value > 0
            tile = _tile(label, value, color if (on or not can_tint) else P.FAINT, tinted=on)
            self.tiles.append(tile)
            stats.addWidget(tile, 1)
        root.addLayout(stats)

        # ── 업무 목록 ──
        sections = [('지난 마감', s['overdue']), ('오늘', s['today']),
                    ('이번 주', s['week']), ('중요 · 기한 없음', s['important'])]
        shown = [(n, items) for n, items in sections if items]
        dots = {'지난 마감': P.DANGER, '오늘': '#E8590C', '이번 주': P.SAT, '중요 · 기한 없음': P.DOT_PRIORITY}
        if shown:
            lst = QVBoxLayout()
            lst.setSpacing(2)
            for i, (name, items) in enumerate(shown):
                if i:
                    lst.addSpacing(12)
                hdr = QHBoxLayout()
                hdr.setSpacing(7)
                dot = QLabel()
                dot.setFixedSize(8, 8)
                dot.setStyleSheet(f'background: {dots[name]}; border-radius: 4px;')
                hdr.addWidget(dot)
                sec = QLabel(name)
                sec.setObjectName('Section')
                cnt = QLabel(str(len(items)))
                cnt.setObjectName('Count')
                hdr.addWidget(sec)
                hdr.addWidget(cnt)
                hdr.addStretch()
                lst.addLayout(hdr)
                lst.addSpacing(2)
                for t in items[:_MAX_ROWS]:
                    lst.addWidget(self._task_row(t, overdue=(name == '지난 마감')))
                if len(items) > _MAX_ROWS:
                    more = QLabel(f'외 {len(items) - _MAX_ROWS}건 더 — 메모 창이나 달력에서 볼 수 있어요')
                    more.setObjectName('Muted')
                    more.setContentsMargins(10, 2, 0, 0)
                    lst.addWidget(more)
            root.addLayout(lst)
        else:
            root.addWidget(self._empty_state())

        # ── 예산 (카드 안은 모두 가운데 정렬) ──
        b = s.get('budget')
        if b:
            card = QFrame()
            card.setObjectName('Card')
            lay = QVBoxLayout(card)
            lay.setContentsMargins(20, 16, 20, 14)
            lay.setSpacing(6)
            center = Qt.AlignHCenter
            lbl = QLabel('예산 집행률')
            lbl.setObjectName('Section')
            lay.addWidget(lbl, 0, center)
            val = QLabel(f"{b['pct']:.1f}%")
            val.setFont(_font(26, QFont.DemiBold))
            val.setAlignment(Qt.AlignCenter)
            val.setStyleSheet(f"color: {_level_color(b['pct']) if b['pct'] >= 80 else P.TEXT};")
            lay.addWidget(val)
            lay.addSpacing(2)
            lay.addWidget(_Meter(b['pct']))
            lay.addSpacing(2)
            detail = QLabel(f"집행 {dashboard.format_won(b['spent'])} / 편성 {dashboard.format_won(b['allocated'])}")
            detail.setObjectName('Muted')
            detail.setAlignment(Qt.AlignCenter)
            lay.addWidget(detail)
            rem = QLabel(f"잔액 {dashboard.format_won(b['remaining'])}")
            rem.setAlignment(Qt.AlignCenter)
            rem.setStyleSheet(f"color: {P.DANGER if b['remaining'] < 0 else P.TEXT2}; font-weight: 600;")
            lay.addWidget(rem)
            btn_budget = QPushButton('예산 창 열기 ›')
            btn_budget.setObjectName('Link')
            btn_budget.setCursor(Qt.PointingHandCursor)
            btn_budget.clicked.connect(self.open_budget.emit)
            lay.addWidget(btn_budget, 0, center)
            self.budget_labels = (val, detail, rem)
            root.addWidget(card)

        # ── 아래 ──
        rule = QFrame()
        rule.setObjectName('Rule')
        rule.setFixedHeight(1)
        root.addWidget(rule)
        foot = QHBoxLayout()
        self.chk_start = QCheckBox('하루 한 번, 앱을 켤 때 보기')
        self.chk_start.setChecked(get_setting(SHOW_ON_START_KEY, '1') != '0')
        self.chk_start.toggled.connect(lambda on: set_setting(SHOW_ON_START_KEY, '1' if on else '0'))
        foot.addWidget(self.chk_start)
        foot.addStretch()
        btn_close = QPushButton('닫기')
        btn_close.setObjectName('Primary')
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setDefault(True)
        btn_close.clicked.connect(self.close)
        foot.addWidget(btn_close)
        root.addLayout(foot)

    def _empty_state(self) -> QWidget:
        from .calendar_view import P
        import qtawesome as qta
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 18, 0, 18)
        lay.setSpacing(6)
        ic = QLabel()
        ic.setPixmap(qta.icon('ph.check-circle', color=P.FAINT).pixmap(34, 34))
        ic.setAlignment(Qt.AlignCenter)
        t = QLabel('급한 업무가 없어요')
        t.setFont(_font(11, QFont.DemiBold))
        t.setAlignment(Qt.AlignCenter)
        sub = QLabel('메모 창에서 업무에 마감일을 넣으면 여기에 모여요.')
        sub.setObjectName('Muted')
        sub.setAlignment(Qt.AlignCenter)
        for x in (ic, t, sub):
            lay.addWidget(x)
        return w

    def _task_row(self, t: dict, overdue=False) -> QPushButton:
        """업무 한 줄: D-day 표시(메모 창과 같은 색) · 이름 · 시각. 누르면 메모 창에서 그 업무를 보여준다."""
        from .calendar_view import P
        from .widgets import DdayLabel
        import qtawesome as qta
        btn = QPushButton()
        btn.setObjectName('TaskRow')
        btn.setCursor(Qt.PointingHandCursor)
        lay = QHBoxLayout(btn)
        lay.setContentsMargins(8, 6, 12, 6)
        lay.setSpacing(8)
        dday = t.get('dday') or ''
        if overdue:
            tone = 'overdue'
        elif dday in ('D-day', '오늘') or '시간' in dday:
            tone = 'today'
        elif dday.startswith('D-') and dday[2:].isdigit():
            n = int(dday[2:])
            tone = 'soon' if n <= 3 else ('week' if n <= 7 else 'later')
        else:
            tone = 'later'
        slot = QWidget()                       # D-day 자리: 기한이 없어도 같은 폭을 비워 둬 이름 줄이 맞는다
        slot.setFixedWidth(px(60))
        slot.setAttribute(Qt.WA_TransparentForMouseEvents)
        sl = QHBoxLayout(slot)
        sl.setContentsMargins(0, 0, 0, 0)
        if dday:
            dd = DdayLabel(dday, dday, tone=tone, align_left=True)
            dd.setFont(pr_font(9))
            dd.setFixedWidth(px(60))
            dd.setAttribute(Qt.WA_TransparentForMouseEvents)
            sl.addWidget(dd)
        lay.addWidget(slot)
        star = QLabel()                        # ★ 자리도 고정 (중요하지 않은 줄도 같은 폭)
        star.setFixedWidth(16)
        star.setAttribute(Qt.WA_TransparentForMouseEvents)
        if t.get('priority'):
            star.setPixmap(qta.icon('ph.star-fill', color=P.DOT_PRIORITY).pixmap(13, 13))
        lay.addWidget(star)
        name = QLabel(t.get('name') or '')
        name.setStyleSheet(f'color: {P.TEXT}; font-size: 10.5pt;')
        name.setTextInteractionFlags(Qt.NoTextInteraction)
        name.setAttribute(Qt.WA_TransparentForMouseEvents)
        lay.addWidget(name, 1)
        if t.get('time'):
            tm = QLabel(t['time'])
            tm.setObjectName('Muted')
            tm.setAttribute(Qt.WA_TransparentForMouseEvents)
            lay.addWidget(tm)
        btn.name_label = name
        self.rows.append(btn)
        btn.setMinimumHeight(38)
        btn.setToolTip('눌러서 메모 창에서 보기')
        btn.setAccessibleName(f"{dday} {t.get('name', '')}")
        btn.clicked.connect(lambda: self.jump_requested.emit(t))
        return btn
