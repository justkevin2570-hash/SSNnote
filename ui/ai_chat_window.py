"""업무창의 AI 대화 창 (앱 전체에 하나).

- 붙는 자리 3가지: 업무창 안에 넣기(기본: 업무 목록 아래, 경계선을 끌어 높이 조절) · 오른쪽에 붙이기 · 아래쪽에 붙이기.
  오른쪽·아래쪽은 업무창(합체한 묶음이면 묶음 전체)을 따라 움직이고, 화면 끝이면 반대쪽(왼쪽·위)에 붙는다.
  맞닿는 곳은 테두리 한 줄을 함께 쓰고, 머리줄은 업무창 제목줄과 같은 높이·색이다.
- ⋯ 메뉴: 붙는 자리 · 배경색·글자·줄간격 · 대화 지우기. ✕로 숨기기. 고른 값은 모두 기억한다 (처음엔 열림 · 안에 넣기).
- 대화: 내 질문은 오른쪽 말풍선, AI 답은 왼쪽 글. 대화 내용은 저장하지 않는다 (앱을 끄면 지워짐).
- 모든 업무창의 업무를 보고 답하고, 새 업무는 붙어 있는 업무창에 넣는다 (task_ai).
- AI는 제안만 한다. [넣기]를 눌러야 반영되고, 넣은 뒤 되돌리기 한 번이 된다.
"""
import html
import re

import qtawesome as qta
from PyQt5.QtCore import QEvent, QRect, QSize, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QScrollArea, QSizePolicy, QVBoxLayout, QWidget)

from .ai_app import AppAdapter
from .common import DEFAULT_NOTE_FONT, NOTE_FONTS, TITLE_BAR_HEIGHT, note_font, px
from .theme import DEFAULT_NOTE_COLOR, T, note_color, tc, tct

SIDE_KEY, SIZE_KEY, OPEN_KEY = 'ai_chat_side', 'ai_chat_size', 'ai_chat_open'
HOST_KEY = 'ai_chat_host'           # 마지막으로 붙어 있던 업무창 id
COLOR_KEY, FONT_KEY, SCALE_KEY, GAP_KEY = 'ai_chat_color', 'ai_chat_font', 'ai_chat_scale', 'ai_chat_gap'
SIDES = (('inside', '업무창 안에 넣기'), ('left', '왼쪽에 붙이기'), ('right', '오른쪽에 붙이기'),
         ('bottom', '아래쪽에 붙이기'))
DEFAULT_SIDE = 'inside'
DEFAULT_SIZE = {'left': 320, 'right': 320, 'bottom': 300, 'inside': 320}
MIN_SIZE = {'left': 240, 'right': 240, 'bottom': 220, 'inside': 220}
GAP = -1                          # 업무창과 1px 겹쳐 테두리 한 줄을 함께 씀
BORDER = 'rgba(0,0,0,0.4)'        # 업무창 테두리와 같은 색
LINE_HEIGHT = {0: 135, 5: 155, 10: 175, 15: 195}       # 줄 간격 단계 → AI 답 줄 높이(%)
GAPS = (('기본', 0), ('넓게', 5), ('더 넓게', 10), ('아주 넓게', 15))
SAME_AS_NOTE = '='                # 배경색 「업무창과 같게」 ('' = 처음 값 흰색)
BUBBLE_PAD_V, BUBBLE_PAD_H = 8, 12                     # 내 질문 말풍선 안쪽 여백 (위아래 · 좌우 같게)
CLEAR_CMDS = ('/clear',)                               # 입력칸 명령: 대화 지우기


# ── 붙는 자리 계산 (화면 없이 시험할 수 있게 따로) ─────────────────
def dock_rect(host: QRect, side: str, size: int, screen: QRect = None) -> tuple:
    """업무창(host) 옆 AI 창 자리 → (QRect, 붙은 쪽). 고른 쪽에만 붙는다 — 화면 끝이면 일반 창처럼
    일부가 화면 밖으로 나간다 (반대쪽으로 넘어가지 않음). 1px 겹쳐 맞닿은 테두리가 한 줄로 보이게 한다."""
    if side == 'bottom':
        return QRect(host.left(), host.bottom() + 1 + GAP, host.width(), size), 'bottom'
    if side == 'left':
        return QRect(host.left() - GAP - size, host.top(), size, host.height()), 'left'
    return QRect(host.right() + 1 + GAP, host.top(), size, host.height()), 'right'


def answer_html(text: str, line_height: int = 135) -> str:
    """AI 답 글 → 읽기 쉬운 HTML: 문단 사이 간격, '- ' 목록은 들여쓰기(둘째 줄부터 글머리 뒤에 맞춤)."""
    out = []
    for raw in (text or '').strip().split('\n'):
        line = raw.strip()
        if not line:
            out.append('<div style="font-size:4pt">&nbsp;</div>')
            continue
        m = re.match(r'^(?:[-•·*]|\d+[.)])\s+(.*)$', line)
        if m:
            mark = '•' if line[0] in '-•·*' else line.split()[0]
            out.append(f'<p style="margin:0 0 3px 14px; text-indent:-12px; line-height:{line_height}%">'
                       f'{html.escape(mark)}&nbsp;{html.escape(m.group(1))}</p>')
        else:
            out.append(f'<p style="margin:0 0 3px 0; line-height:{line_height}%">{html.escape(line)}</p>')
    return ''.join(out)


def _setting(key, default=''):
    try:
        from db import get_setting
        return get_setting(key, default)
    except Exception:
        return default


def _save(key, value):
    try:
        from db import set_setting
        set_setting(key, str(value))
    except Exception:
        pass


def saved_side():
    s = _setting(SIDE_KEY, DEFAULT_SIDE)
    return s if s in dict(SIDES) else DEFAULT_SIDE


def saved_open():
    return _setting(OPEN_KEY, '1') != '0'


class _EdgeGrip(QWidget):
    """먼 쪽 가장자리를 끌어 AI 창 폭(오른쪽에 붙을 때)이나 높이(아래쪽)를 바꾼다."""

    def __init__(self, win):
        super().__init__(win)
        self._win = win
        self._start = None
        self._where = 'right'

    def set_side(self, where):
        self._where = where
        self.setCursor(Qt.SizeVerCursor if where in ('bottom', 'top') else Qt.SizeHorCursor)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._start = (e.globalPos(), self._win.dock_size)

    def mouseMoveEvent(self, e):
        if not self._start:
            return
        p0, size0 = self._start
        d = e.globalPos() - p0
        delta = {'right': d.x(), 'left': -d.x(), 'bottom': d.y(), 'top': -d.y()}[self._where]
        self._win.set_dock_size(size0 + delta)

    def mouseReleaseEvent(self, e):
        if self._start:
            self._start = None
            _save(SIZE_KEY + '_' + self._win.side, self._win.dock_size)


class AiChatWindow(QWidget):
    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_StyledBackground, True)     # 업무창 안에 넣어도 바탕색이 칠해지게
        self.setObjectName('AiChat')
        self.host = None
        self.side = saved_side()
        self.dock_size = self._saved_size(self.side)
        self._where = self.side
        self._worker = None
        self._history = []
        self._msgs = []                   # 대화 칸 (위젯)
        self._bubbles = []                # 내 질문 말풍선 (폭을 창에 맞춤)
        self._q = ''
        self._pending = []
        self._pending_del = []
        self._undo = None
        self.color = _setting(COLOR_KEY, '')                     # '' = 업무창과 같게
        self.font_key = _setting(FONT_KEY, DEFAULT_NOTE_FONT) or DEFAULT_NOTE_FONT
        if self.font_key not in dict(NOTE_FONTS):
            self.font_key = DEFAULT_NOTE_FONT
        try:
            self.scale = float(_setting(SCALE_KEY, '1.0') or 1.0)
            self.gap = int(_setting(GAP_KEY, '0') or 0)
        except ValueError:
            self.scale, self.gap = 1.0, 0
        self._build()
        self._apply_style()
        self._update_enabled()
        from . import ai_connect
        ai_connect.on_change(self._on_ai_changed)

    # ── 화면 ──────────────────────────────────────────────────────
    def _build(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(1, 1, 1, 1)
        v.setSpacing(0)

        self.head = QFrame()
        self.head.setObjectName('AiHead')
        hl = QHBoxLayout(self.head)
        hl.setContentsMargins(12, 0, 6, 0)
        hl.setSpacing(1)
        self.title_icon = QLabel()
        self.title = QLabel('AI 대화')
        self.title.setObjectName('AiTitle')
        self.btn_more = self._icon('ph.dots-three-bold', '붙는 자리 · 배경색·글자·줄간격 · 대화 지우기')
        self.btn_more.clicked.connect(self._more_menu)
        self.btn_hide = self._icon('ph.x-bold', '')
        self.btn_hide.clicked.connect(self.hide_by_user)
        self._head_layout = hl
        self._head_mirrored = None
        self._layout_head(False)
        v.addWidget(self.head)

        body = QVBoxLayout()
        body.setContentsMargins(10, 8, 10, 10)
        body.setSpacing(6)
        self.scroll = QScrollArea()
        self.scroll.setObjectName('AiScroll')
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.chat = QWidget()
        self.chat.setObjectName('AiChatList')
        self.chat_layout = QVBoxLayout(self.chat)
        self.chat_layout.setContentsMargins(0, 2, 0, 2)          # 말풍선 오른쪽 끝 = 보내기 버튼 오른쪽 끝
        self.chat_layout.setSpacing(10)
        self.chat_layout.addStretch()
        self.scroll.setWidget(self.chat)
        self.scroll.viewport().installEventFilter(self)
        self.scroll.verticalScrollBar().rangeChanged.connect(
            lambda _a, b: self.scroll.verticalScrollBar().setValue(b) if self._stick else None)
        self._stick = True
        body.addWidget(self.scroll, 1)

        self.proposal = QFrame()
        self.proposal.setObjectName('Proposal')
        pl = QVBoxLayout(self.proposal)
        pl.setContentsMargins(10, 8, 10, 8)
        pl.setSpacing(6)
        self.lbl_proposal = QLabel()
        self.lbl_proposal.setWordWrap(True)
        self.lbl_proposal.setTextFormat(Qt.RichText)
        pl.addWidget(self.lbl_proposal)
        prow = QHBoxLayout()
        prow.addStretch()
        self.btn_apply = QPushButton('넣기')
        self.btn_apply.setObjectName('Primary')
        self.btn_apply.setCursor(Qt.PointingHandCursor)
        self.btn_apply.clicked.connect(self._apply)
        self.btn_skip = QPushButton('취소')
        self.btn_skip.setCursor(Qt.PointingHandCursor)
        self.btn_skip.clicked.connect(self._hide_proposal)
        prow.addWidget(self.btn_apply)
        prow.addWidget(self.btn_skip)
        pl.addLayout(prow)
        self.proposal.hide()
        body.addWidget(self.proposal)

        # 지우기 확인 (빨간 칸, 따로 [지우기])
        self.proposal_del = QFrame()
        self.proposal_del.setObjectName('ProposalDel')
        dl = QVBoxLayout(self.proposal_del)
        dl.setContentsMargins(10, 8, 10, 8)
        dl.setSpacing(6)
        self.lbl_delete = QLabel()
        self.lbl_delete.setWordWrap(True)
        self.lbl_delete.setTextFormat(Qt.RichText)
        dl.addWidget(self.lbl_delete)
        drow = QHBoxLayout()
        drow.addStretch()
        self.btn_delete = QPushButton('지우기')
        self.btn_delete.setObjectName('Danger')
        self.btn_delete.setCursor(Qt.PointingHandCursor)
        self.btn_delete.clicked.connect(self._apply_deletes)
        self.btn_delete_skip = QPushButton('취소')
        self.btn_delete_skip.setCursor(Qt.PointingHandCursor)
        self.btn_delete_skip.clicked.connect(self._hide_deletes)
        drow.addWidget(self.btn_delete)
        drow.addWidget(self.btn_delete_skip)
        dl.addLayout(drow)
        self.proposal_del.hide()
        body.addWidget(self.proposal_del)
        # 확인 칸은 글 높이만큼 (대화 칸이 줄어듦) — 줄이 잘려 제안이 안 보이는 일이 없게
        for f, lb in ((self.proposal, self.lbl_proposal), (self.proposal_del, self.lbl_delete)):
            lb.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
            f.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        self.scroll.setMinimumHeight(px(60))

        self.quick_buttons = []           # 빠른 질문 버튼은 없앰 (직접 묻기)

        row = QHBoxLayout()
        row.setSpacing(6)
        self.input = QLineEdit()
        self.input.returnPressed.connect(lambda: self.ask(self.input.text()))
        self.btn_send = QPushButton()
        self.btn_send.setObjectName('Primary')
        self.btn_send.setIcon(qta.icon('fa5s.paper-plane', color='white'))
        self.btn_send.setFixedWidth(px(38))
        self.btn_send.setCursor(Qt.PointingHandCursor)
        self.btn_send.setToolTip('보내기 (Enter)')
        self.btn_send.clicked.connect(lambda: self.ask(self.input.text()))
        row.addWidget(self.input, 1)
        row.addWidget(self.btn_send)
        body.addLayout(row)
        self.lbl_status = QLabel()
        self.lbl_status.setObjectName('Faint')
        self.lbl_status.hide()
        body.addWidget(self.lbl_status)
        v.addLayout(body, 1)

        self.grip = _EdgeGrip(self)
        self._show_welcome()

    def _icon(self, name, tip):
        b = QPushButton()
        b.setObjectName('Icon')
        b._icon_name = name
        b.setIconSize(QSize(px(18), px(18)))
        b.setFixedSize(px(30), px(28))
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        return b

    # ── 모양 (배경색·글자·줄간격 · 붙는 자리) ──────────────────────
    def _bg(self):
        """AI 창 바탕색 (표시색). 처음(고르지 않음)은 흰색, 「=」를 고르면 붙은 업무창과 같은 색."""
        c = self.color
        if c == SAME_AS_NOTE:
            try:
                c = getattr(self.host, '_bg_color', '') or DEFAULT_NOTE_COLOR
            except RuntimeError:
                c = DEFAULT_NOTE_COLOR
        return note_color(c) if c else tc('#FFFFFF')

    def _layout_head(self, mirrored):
        """머리줄 배치. 업무창 왼쪽에 붙으면 좌우를 뒤집어 ✕ · ⋯ 를 바깥(왼쪽) 끝에 둔다
        — 업무창 제목줄 버튼과 맞붙어 한 줄로 섞여 보이지 않게."""
        if mirrored == self._head_mirrored:
            return
        self._head_mirrored = mirrored
        hl = self._head_layout
        while hl.count():
            hl.takeAt(0)
        if mirrored:
            hl.setContentsMargins(6, 0, 12, 0)
            hl.addWidget(self.btn_hide, 0, Qt.AlignVCenter)
            hl.addWidget(self.btn_more, 0, Qt.AlignVCenter)
            hl.addStretch()
            hl.addWidget(self.title_icon)
            hl.addSpacing(4)
            hl.addWidget(self.title)
        else:
            hl.setContentsMargins(12, 0, 6, 0)
            hl.addWidget(self.title_icon)
            hl.addSpacing(4)
            hl.addWidget(self.title)
            hl.addStretch()
            hl.addWidget(self.btn_more, 0, Qt.AlignVCenter)
            hl.addWidget(self.btn_hide, 0, Qt.AlignVCenter)

    def _apply_style(self):
        bg = self._bg()
        # 머리줄: 업무창 제목줄과 같은 계산(바탕보다 진하게). 흰 바탕이면 아주 옅은 회색
        head_bg = QColor(bg).darker(104 if QColor(bg).lightness() > 250 else 115).name()
        inside = self.side == 'inside'
        ink = tct('#1E293B')
        faint = tct('#64748B')
        bubble = tc('rgba(0,0,0,0.07)')
        f = note_font(10 * self.scale, self.font_key)
        family = f.family()
        size = f.pointSizeF()
        border = 'none' if inside else f'1px solid {tc(BORDER)}'
        self.setStyleSheet(f"""
            QWidget#AiChat {{ background: {bg}; border: {border}; }}
            QWidget#AiChat QWidget {{ font-family: '{family}'; font-size: {size:.1f}pt; color: {ink}; }}
            QLabel {{ background: transparent; }}
            QFrame#AiHead {{ background: {'transparent' if inside else head_bg}; border: none; }}
            QLabel#AiTitle {{ font-weight: 700; color: {T.TEXT_ON_NOTE}; }}
            QScrollArea#AiScroll, QWidget#AiChatList {{ background: transparent; border: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; margin: 0; }}
            QScrollBar::handle:vertical {{ background: {tc('rgba(0,0,0,0.22)')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
            QLabel#UserBubble {{ background: {bubble}; border-radius: 14px; padding: {BUBBLE_PAD_V}px {BUBBLE_PAD_H}px; }}
            QLabel#Faint {{ color: {faint}; font-size: {max(8.0, size - 1):.1f}pt; }}
            QLabel#Error {{ color: {tc('#BE123C')}; }}
            QPushButton {{ background: {tc('rgba(255,255,255,0.55)')}; border: 1px solid {tc('rgba(0,0,0,0.12)')};
                           border-radius: 7px; padding: 3px 10px; }}
            QPushButton:hover {{ background: {tc('rgba(255,255,255,0.85)')}; }}
            QPushButton:disabled {{ color: {tc('#94A3B8')}; }}
            QPushButton#Icon {{ background: transparent; border: none; padding: 0; border-radius: 6px; }}
            QPushButton#Icon:hover {{ background: {tc('rgba(0,0,0,0.10)')}; }}
            QPushButton#Quick {{ border-radius: 10px; padding: 2px 9px; }}
            QWidget#AiChat QPushButton#Primary {{ background: {T.ACCENT}; border: 1px solid {T.ACCENT}; color: white;
                                                 font-weight: 700; }}
            QLineEdit {{ background: {tc('#FFFFFF')}; border: 1px solid {tc('rgba(0,0,0,0.15)')}; border-radius: 8px;
                         padding: 6px 8px; }}
            QLineEdit:focus {{ border: 1px solid {T.ACCENT}; }}
            QFrame#Proposal {{ background: {tc('#FFF8E1')}; border: 1px solid {tc('#F2D48A')}; border-radius: 8px; }}
            QFrame#ProposalDel {{ background: {tc('#FFF1F2')}; border: 1px solid {tc('#FECDD3')}; border-radius: 8px; }}
            QWidget#AiChat QPushButton#Danger {{ background: {tc('#E11D48')}; border: 1px solid {tc('#E11D48')};
                                                color: white; font-weight: 700; }}
        """)
        self.layout().setContentsMargins(0, 0, 0, 0) if inside else self.layout().setContentsMargins(1, 1, 1, 1)
        self.head.setFixedHeight(px(30) if inside else TITLE_BAR_HEIGHT)   # 바깥: 업무창 제목줄과 같은 높이
        icon_color = QColor(0, 0, 0, 165)
        from .chrome import ai_chat_icon
        self.title_icon.setPixmap(ai_chat_icon(icon_color).pixmap(px(18), px(18)))
        for b in (self.btn_more, self.btn_hide):
            b.setIcon(qta.icon(b._icon_name, color=icon_color))
        self.grip.setVisible(not inside)
        self._render_all()

    def _set_and_save(self, key, attr, value):
        setattr(self, attr, value)
        _save(key, value)
        self._apply_style()

    def set_color(self, color):
        self._set_and_save(COLOR_KEY, 'color', color or '')

    def set_font(self, key):
        self._set_and_save(FONT_KEY, 'font_key', key if key in dict(NOTE_FONTS) else DEFAULT_NOTE_FONT)

    def set_scale(self, scale):
        self._set_and_save(SCALE_KEY, 'scale', scale)

    def set_gap(self, gap):
        self._set_and_save(GAP_KEY, 'gap', gap)

    def _more_menu(self):
        from . import menu_style, style_menu
        m = menu_style.make_menu(self)
        menu_style.add_header(m, '붙는 자리')
        for side, text in SIDES:
            menu_style.add_action(m, text, lambda s=side: self.set_side(s), checked=self.side == side)
        m.addSeparator()

        def font_missing(k):
            host = self.host
            if host is not None and hasattr(host, '_font_missing'):
                host._font_missing(k, m, apply=self.set_font)
        style_menu.build(m, color=self.color, on_color=self.set_color, font_key=self.font_key,
                         on_font=self.set_font, on_font_missing=font_missing,
                         scale=self.scale, on_scale=self.set_scale, gaps=GAPS, gap=self.gap, on_gap=self.set_gap,
                         same_as_note=True)
        m.addSeparator()
        menu_style.add_action(m, '대화 지우기', self._clear_chat, icon='ph.eraser')
        self._last_menu = m
        m.exec_(self.btn_more.mapToGlobal(self.btn_more.rect().bottomLeft()))

    # ── 대화 칸 ───────────────────────────────────────────────────
    def _show_welcome(self):
        """빈 대화 칸 (안내 문구 없이 — 입력칸 안내 글로 충분)."""
        self._clear_widgets()

    def _clear_widgets(self):
        for w in self._msgs:
            w.setParent(None)
            w.deleteLater()
        self._msgs, self._bubbles = [], []

    def _add_msg(self, kind, text, link=None):
        """kind: user(오른쪽 말풍선) · ai(왼쪽 글) · note(작은 회색 글) · error · welcome."""
        if kind != 'welcome' and self._msgs and getattr(self._msgs[0], '_kind', '') == 'welcome':
            self._clear_widgets()
        w = QWidget()
        w._kind, w._text, w._link = kind, text, link
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lbl = QLabel()
        lbl.setWordWrap(True)
        lbl.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.LinksAccessibleByMouse)
        lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        w._label = lbl
        if kind == 'user':
            lbl.setObjectName('UserBubble')
            lay.addStretch(1)
            lay.addWidget(lbl, 0, Qt.AlignRight)
            self._bubbles.append(lbl)
        else:
            lbl.setObjectName({'note': 'Faint', 'welcome': 'Faint', 'error': 'Error'}.get(kind, 'AiText'))
            lay.addWidget(lbl, 1)
            if link:
                lbl.linkActivated.connect(lambda _h: self._undo_last())
        self._fill(w)
        self.chat_layout.addWidget(w)
        self._msgs.append(w)
        if len(self._msgs) > 80:                  # 오래된 것부터 정리
            old = self._msgs.pop(0)
            if old._label in self._bubbles:
                self._bubbles.remove(old._label)
            old.setParent(None)
            old.deleteLater()
        self._fit_bubbles()
        self._stick = True
        return w

    def _fill(self, w):
        lbl, kind, text = w._label, w._kind, w._text
        lh = LINE_HEIGHT.get(self.gap, 135)
        if kind == 'ai':
            lbl.setTextFormat(Qt.RichText)
            lbl.setText(answer_html(text, lh))
        elif kind == 'note' and w._link:
            lbl.setTextFormat(Qt.RichText)
            lbl.setText(f'{html.escape(text)} <a href="undo">되돌리기</a>')
        elif kind == 'user':
            # 말풍선: 줄 높이를 넣으면 늘어난 간격이 글자 아래에만 붙어 글자가 위로 뜬다 → 그냥 글
            lbl.setTextFormat(Qt.PlainText)
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            lbl.setIndent(0)                   # 기본 들여쓰기(글자 x 반 폭)가 왼쪽에만 붙어 좌우가 어긋남
            lbl.setText(text)
        else:
            lbl.setTextFormat(Qt.RichText)
            body = html.escape(text).replace('\n', '<br>')
            lbl.setText(f'<div style="line-height:{lh}%">{body}</div>')

    def _render_all(self):
        for w in self._msgs:
            self._fill(w)
        self._fit_bubbles()

    def _fit_bubbles(self):
        """내 질문 말풍선은 창 폭의 85%까지."""
        width = max(120, int(self.scroll.viewport().width() * 0.85))
        for b in self._bubbles:
            try:
                b.setMaximumWidth(width)
            except RuntimeError:
                pass

    def eventFilter(self, obj, e):
        if obj is self.scroll.viewport() and e.type() == QEvent.Resize:
            self._fit_bubbles()
        return super().eventFilter(obj, e)

    def messages(self):
        """지금 대화 [(종류, 글)] — 테스트·확인용."""
        return [(w._kind, w._text) for w in self._msgs]

    # ── 붙이기 · 따라가기 ──────────────────────────────────────────
    @staticmethod
    def _saved_size(side):
        try:
            return max(MIN_SIZE[side], int(_setting(SIZE_KEY + '_' + side, DEFAULT_SIZE[side])))
        except (ValueError, KeyError):
            return DEFAULT_SIZE.get(side, 240)

    def _host_rect(self):
        """붙을 기준: 업무창, 합체한 묶음이면 묶음 전체를 감싸는 사각형."""
        h = self.host
        group = [m for m in (getattr(h, '_link_group', None) or {h}) if m.isVisible()] or [h]
        r = group[0].frameGeometry()
        for m in group[1:]:
            r = r.united(m.frameGeometry())
        return r

    def follow(self):
        """업무창이 움직이거나 크기가 바뀌면 붙은 자리로 옮긴다 (안에 넣었으면 할 일 없음)."""
        if self.host is None or self.side == 'inside' or not self.isVisible():
            return
        try:
            host_rect = self._host_rect()
            scr = QApplication.screenAt(host_rect.center()) or QApplication.primaryScreen()
        except RuntimeError:          # 업무창이 닫힘
            self.detach()
            return
        rect, where = dock_rect(host_rect, self.side, self.dock_size, scr.availableGeometry())
        self._where = where
        self.setGeometry(rect)
        self._place_grip()

    def _place_grip(self):
        self._layout_head(self.side != 'inside' and self._where == 'left')
        if self.side == 'inside' or self._where not in ('right', 'left', 'bottom', 'top'):
            return
        g = px(6)
        r = self.rect()
        geo = {'right': QRect(r.width() - g, 0, g, r.height()), 'left': QRect(0, 0, g, r.height()),
               'bottom': QRect(0, r.height() - g, r.width(), g), 'top': QRect(0, 0, r.width(), g)}[self._where]
        self.grip.set_side(self._where)
        self.grip.setGeometry(geo)
        self.grip.raise_()

    def set_dock_size(self, size):
        geo = QApplication.primaryScreen().availableGeometry()
        limit = geo.width() // 2 if self.side in ('left', 'right') else geo.height() // 2
        self.dock_size = max(MIN_SIZE[self.side], min(int(size), max(MIN_SIZE[self.side], limit)))
        self.follow()

    def set_side(self, side):
        if side not in dict(SIDES) or side == self.side:
            return
        self.side = side
        self._where = side
        _save(SIDE_KEY, side)                         # 다음에 켤 때도 이 자리
        self.dock_size = self._saved_size(side)
        host = self.host
        if host is not None:
            self.attach(host, remember=False)
        else:
            self._apply_style()

    def _slot(self, host):
        return getattr(host, 'ai_slot', None)

    def _embed(self, host):
        """업무창 안(업무 목록 아래)에 넣는다."""
        slot = self._slot(host)
        if slot is None:
            return False
        if self.parent() is not slot:
            self.setParent(slot)                       # 창이 아닌 업무창 속 칸이 됨
            slot.layout().addWidget(self)
        slot.show()
        self.show()
        split = host.list_split
        total = max(split.height(), 1)
        h = min(self._saved_size('inside'), max(MIN_SIZE['inside'], total - px(110)))
        split.setSizes([max(1, total - h), h])
        return True

    def _unembed(self):
        """업무창 밖으로 꺼낸다 (따로 떠 있는 창으로)."""
        slot = self.parent()
        if slot is not None:
            if slot.layout() is not None:
                slot.layout().removeWidget(self)
            self.setParent(None, Qt.FramelessWindowHint | Qt.Tool)
            try:
                slot.hide()
            except RuntimeError:
                pass

    def attach(self, host, remember=True):
        """host 업무창에 붙여(또는 안에 넣어) 보인다."""
        if self.host is not None and self.host is not host:
            self._mark_host(False)                     # 전에 붙어 있던 업무창의 버튼은 끔
        try:
            pin = bool(getattr(host, 'pin_active', False))
        except RuntimeError:
            return
        self.host = host
        if self.side == 'inside':
            if not self._embed(host):
                return
        else:
            self._unembed()
            flags = Qt.FramelessWindowHint | Qt.Tool | (Qt.WindowStaysOnTopHint if pin else Qt.Widget)
            if self.windowFlags() != flags:
                self.setWindowFlags(flags)
            self.show()
            self.follow()
            self.raise_()
        self._apply_style()
        self._mark_host(True)
        self._update_enabled()
        wid = getattr(host, 'window_id', None)
        if wid is not None:
            _save(HOST_KEY, wid)                       # 다음에 켤 때 이 업무창에
        if remember:
            _save(OPEN_KEY, '1')

    def hide_chat(self):
        """보이지 않게 (안에 넣었으면 업무 목록이 다시 넓어지게 칸도 숨김)."""
        self.hide()
        slot = self.parent()
        if slot is not None:
            try:
                slot.hide()
            except RuntimeError:
                pass
        self._mark_host(False)

    def detach(self):
        """업무창이 닫힐 때: 업무창과 함께 지워지지 않게 밖으로 꺼내 숨긴다."""
        self._mark_host(False)
        self._unembed()
        self.host = None
        self.hide()

    def hide_by_user(self):
        _save(OPEN_KEY, '0')
        self.hide_chat()

    def is_shown(self):
        return not self.isHidden() and (self.parent() is None or not self.parent().isHidden())

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._place_grip()

    def showEvent(self, e):
        super().showEvent(e)
        self._mark_host(True)

    def _mark_host(self, on, host=None):
        """붙은 업무창 제목줄의 AI 버튼을 열림(옅은 바탕)/닫힘으로."""
        host = host or self.host
        try:
            host.set_ai_button_active(on)              # 업무창 아래 줄의 AI 버튼
        except (AttributeError, RuntimeError):
            pass

    # ── AI 연결 (설정은 업무창 메뉴 · 공문 작성 · 예산 창에서) ───────
    def _ai_ready(self):
        try:
            from ai_client import load_ai_mode
            return load_ai_mode() != 'none'
        except Exception:
            return False

    def _update_enabled(self):
        busy = self._worker is not None and self._worker.isRunning()
        ready = self._ai_ready()
        self.input.setEnabled(not busy)               # AI가 꺼져 있어도 /clear는 되게
        self.btn_send.setEnabled(ready and not busy)
        self.input.setPlaceholderText('AI에게 물어보기 (Enter) · /clear 대화 지우기' if ready
                                      else 'AI가 꺼져 있어요 — 메뉴 → 설정 → AI 연결에서 연결하세요')

    def _on_ai_changed(self):
        try:
            self._update_enabled()
        except RuntimeError:
            pass

    def _clear_chat(self):
        """대화 지우기 (⋯ → 대화 지우기, 입력칸 /clear): 대화 · 이전 질문 기억 · 넣기·지우기 확인 칸."""
        self._history = []
        self._undo = None
        self._hide_proposal()
        self._hide_deletes()
        self._show_welcome()

    # ── 묻기 ──────────────────────────────────────────────────────
    def ask(self, question):
        q = (question or '').strip()
        if not q or (self._worker is not None and self._worker.isRunning()) or self.host is None:
            return
        if q.startswith('/'):                          # 입력칸 명령 (AI에게 보내지 않음)
            self.input.clear()
            if q.lower() in CLEAR_CMDS:
                self._clear_chat()
            else:
                self._add_msg('error', f'모르는 명령이에요: {q} — /clear(대화 지우기)만 쓸 수 있어요.')
            return
        if not self._ai_ready():
            self._add_msg('error', 'AI가 꺼져 있어요. 메뉴 → 설정 → AI 연결에서 연결하세요.')
            return
        import ai_actions
        app = ChatApp(self)
        ctx = ai_actions.build_state(app)
        self._ctx = ctx
        self._q = q
        self.input.clear()
        self._hide_proposal()
        self._hide_deletes()
        self._add_msg('user', q)
        self._set_status('AI가 생각하는 중…')
        self._worker = ai_actions.AskWorker(q, ctx, self._history)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._update_enabled)
        self._update_enabled()
        self._worker.start()

    def _set_status(self, text):
        self.lbl_status.setText(text)
        self.lbl_status.setVisible(bool(text))

    def _on_done(self, reply, masked):
        import ai_actions
        answer = reply.get('answer') or ''
        self._history = (self._history + [(self._q, answer)])[-2:]
        if answer:
            self._add_msg('ai', answer)
        if masked:
            try:
                from privacy import describe
                self._add_msg('note', f'🛡 개인정보 {describe(masked)}을(를) 가리고 보냈어요.')
            except Exception:
                pass
        self._set_status('')
        now = reply.get('now') or []
        if now:                                         # 보기·열기는 바로
            try:
                ai_actions.execute(now, ChatApp(self))
                self._add_msg('note', ' · '.join(it['desc'] for it in now) + ' 했어요.')
            except Exception as e:
                self._add_msg('error', f'실행하지 못했어요: {e}')
        self._show_proposal(reply.get('changes') or [], reply.get('rejected') or [])
        self._show_deletes(reply.get('deletes') or [])

    def _on_failed(self, msg):
        self._add_msg('error', msg)
        self._set_status('')

    # ── 제안 [넣기] · [지우기] · [되돌리기] ────────────────────────
    def _show_proposal(self, changes, rejected=()):
        self._pending = list(changes)
        if not changes and not rejected:
            self.proposal.hide()
            return
        lines = [f'• {html.escape(c["desc"])}' for c in changes]
        lines += [f'<span style="color:{tc("#94A3B8")}">• {html.escape(r["desc"])} — {html.escape(r["reason"])}'
                  f' (할 수 없어요)</span>' for r in rejected]
        head = '이렇게 할까요?' if changes else '할 수 있는 제안이 없어요'
        self.lbl_proposal.setText(f'<b>{head}</b><br>' + '<br>'.join(lines))
        self.btn_apply.setVisible(bool(changes))
        self.btn_skip.setText('취소' if changes else '닫기')
        self.proposal.show()

    def _hide_proposal(self):
        self._pending = []
        self.proposal.hide()

    def _show_deletes(self, deletes):
        self._pending_del = list(deletes)
        if not deletes:
            self.proposal_del.hide()
            return
        lines = '<br>'.join(f'• {html.escape(d["desc"])}' for d in deletes)
        self.lbl_delete.setText(f'<b style="color:{tc("#BE123C")}">지울까요?</b><br>{lines}<br>'
                                f'<span style="color:{tc("#64748B")}">지운 뒤에도 되돌리기를 한 번 할 수 있어요.</span>')
        self.proposal_del.show()

    def _hide_deletes(self):
        self._pending_del = []
        self.proposal_del.hide()

    def _run(self, items, done_word):
        import ai_actions
        try:
            journal = ai_actions.execute(items, ChatApp(self))
        except Exception as e:
            self._add_msg('error', f'실행하지 못했어요: {e}')
            return
        self._undo = journal
        self._add_msg('note', f'✓ {len(items)}건을 {done_word}.', link='undo')

    def _apply(self):
        if not self._pending or self.host is None:
            return
        items = self._pending
        self._hide_proposal()
        self._run(items, '했어요')

    def _apply_deletes(self):
        if not getattr(self, '_pending_del', None) or self.host is None:
            return
        items = self._pending_del
        self._hide_deletes()
        self._run(items, '지웠어요')

    def _undo_last(self):
        if not self._undo:
            return
        import ai_actions
        journal, self._undo = self._undo, None
        try:
            ai_actions.undo(journal, ChatApp(self))
        except Exception as e:
            self._add_msg('error', f'되돌리지 못했어요: {e}')
            return
        self._add_msg('note', '↩ 하기 전으로 되돌렸어요.')


class ChatApp(AppAdapter):
    """AI 대화 창에서 쓰는 어댑터 (붙어 있는 업무창이 기준)."""

    def __init__(self, chat):
        super().__init__(chat.host, chat=chat, on_clear=chat._clear_chat)


# ── 앱 전체에 하나 ────────────────────────────────────────────────
_win = None


def instance(create=True):
    global _win
    if _win is None and create:
        _win = AiChatWindow()
    return _win


def is_attached_to(host):
    w = instance(create=False)
    return w is not None and w.host is host and w.is_shown()


def toggle(host):
    """AI 버튼: 이 업무창에 열려 있으면 숨기고, 아니면 이 업무창에 연다."""
    w = instance()
    if w.host is host and w.is_shown():
        w.hide_by_user()
    else:
        w.attach(host)


def follow(win):
    """업무창(또는 같은 합체 묶음의 창)이 움직이거나 크기가 바뀌었을 때."""
    w = instance(create=False)
    if w is None or w.host is None:
        return
    try:
        group = getattr(w.host, '_link_group', None) or {w.host}
    except RuntimeError:
        return
    if win is w.host or win in group:
        w.follow()


def inside_resized(host):
    """업무창 안 경계선을 끌어 AI 칸 높이를 바꿨을 때: 그 높이를 기억."""
    w = instance(create=False)
    if w is not None and w.host is host and w.side == 'inside':
        h = host.list_split.sizes()[1]
        if h >= MIN_SIZE['inside']:
            _save(SIZE_KEY + '_inside', h)


def host_restyled(host):
    """업무창 색을 바꿨을 때: 「업무창과 같게」면 AI 창도 그 색으로."""
    w = instance(create=False)
    if w is not None and w.host is host and not w.color:
        try:
            w._apply_style()
        except RuntimeError:
            pass


def host_hidden(host):
    """업무창이 접히거나 메모 모드가 되면 함께 숨는다 (열림 설정은 그대로)."""
    w = instance(create=False)
    if w is not None and w.host is host:
        if w.side == 'inside' and not host.isVisible():
            return                                     # 업무창째 숨은 것: 다시 보이면 그대로
        w.hide_chat()


def host_shown(host):
    """업무창이 다시 보이거나 활성화되면, 열려 있던 AI 창도 다시 붙인다.
    앱을 켤 때 못 붙었으면(창이 늦게 뜸 · 접혀 있다 펼침) 여기서 처음으로 붙인다."""
    if not saved_open():
        return
    w = instance(create=False)
    if w is None or w.host is None:
        auto_open(host)
    elif w.host is host and not w.is_shown():
        w.attach(host, remember=False)


def host_closed(host):
    w = instance(create=False)
    if w is not None and w.host is host:
        w.detach()


def _usable(win):
    try:
        return (win.isVisible() and not getattr(win, '_deleted', False)
                and bool(getattr(win, '_ai_chat_usable', lambda: False)()))
    except RuntimeError:
        return False


def _saved_host_alive(host):
    """기억한 업무창이 지금 떠 있고 AI를 붙일 수 있는 업무창(메모장 · 접힘 · 메모 모드 아님)이면 그 id."""
    try:
        wid = int(_setting(HOST_KEY, '') or 0)
    except ValueError:
        return None
    if not wid:
        return None
    try:
        wins = list(getattr(host, '_open_windows', None) or [])
    except RuntimeError:
        wins = []
    if not wins:
        wins = [t for t in QApplication.topLevelWidgets() if type(t).__name__ == 'MemoWindow']
    for win in wins:
        try:
            if win.window_id == wid and _usable(win):
                return wid
        except RuntimeError:
            pass
    return None


def reset_layout():
    """「창 배치 기본값으로」: 업무창 안에 넣기 · 기본 크기 (열림 여부·붙은 업무창·색·글꼴은 그대로)."""
    for side in DEFAULT_SIZE:
        _save(SIZE_KEY + '_' + side, DEFAULT_SIZE[side])
    _save(SIDE_KEY, DEFAULT_SIDE)
    w = instance(create=False)
    if w is not None:
        w.dock_size = DEFAULT_SIZE[DEFAULT_SIDE]
        if w.side != DEFAULT_SIDE:
            w.set_side(DEFAULT_SIDE)
        elif w.host is not None and w.is_shown():
            w.attach(w.host, remember=False)


def auto_open(host):
    """앱을 켰을 때 (처음엔 열림): 지난번에 붙어 있던 업무창에, 그 창을 쓸 수 없으면 이 업무창에."""
    if not saved_open() or not _usable(host):
        return
    w = instance(create=False)
    if w is not None and w.host is not None:
        return
    want = _saved_host_alive(host)
    if want is not None and want != getattr(host, 'window_id', None):
        return                                         # 기억한 업무창이 스스로 붙임
    instance().attach(host, remember=False)
