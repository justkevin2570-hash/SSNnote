"""업무창 옆에 붙는 AI 대화 창 (앱 전체에 하나).

- 업무창(합체한 묶음이면 묶음 전체)의 오른쪽 또는 아래쪽에 붙어 따라 움직인다. 화면 끝이면 반대쪽(왼쪽·위)에 붙는다.
- ⋯ 메뉴에서 「오른쪽에 붙이기 / 아래쪽에 붙이기」, ✕로 숨기기. 위치·크기·열림 여부는 기억한다 (처음엔 열림).
- 모든 업무창의 업무를 보고 답하고, 새 업무는 붙어 있는 업무창에 넣는다 (task_ai).
- AI는 제안만 한다. [넣기]를 눌러야 반영되고, 넣은 뒤 [되돌리기] 한 번이 된다.
"""
import html

import qtawesome as qta
from PyQt5.QtCore import QPoint, QRect, QSize, Qt
from PyQt5.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel, QLineEdit, QMenu, QPushButton,
                             QTextBrowser, QVBoxLayout, QWidget)

from .common import px
from .theme import T, tc

SIDE_KEY, SIZE_KEY, OPEN_KEY = 'ai_chat_side', 'ai_chat_size', 'ai_chat_open'
HOST_KEY = 'ai_chat_host'           # 마지막으로 붙어 있던 업무창 id
DEFAULT_SIZE = {'right': 320, 'bottom': 240}
MIN_SIZE = {'right': 240, 'bottom': 170}
GAP = 0                           # 업무창과 AI 창 사이 간격


# ── 붙는 자리 계산 (화면 없이 시험할 수 있게 따로) ─────────────────
def dock_rect(host: QRect, side: str, size: int, screen: QRect) -> tuple:
    """업무창(host) 옆 AI 창 자리 → (QRect, 실제로 붙은 쪽). 자리가 없으면 반대쪽(left/top)에 붙인다."""
    if side == 'bottom':
        h = size
        y = host.bottom() + 1 + GAP
        where = 'bottom'
        if y + h - 1 > screen.bottom() and host.top() - GAP - h >= screen.top():
            y, where = host.top() - GAP - h, 'top'
        return QRect(host.left(), y, host.width(), h), where
    w = size
    x = host.right() + 1 + GAP
    where = 'right'
    if x + w - 1 > screen.right() and host.left() - GAP - w >= screen.left():
        x, where = host.left() - GAP - w, 'left'
    return QRect(x, host.top(), w, host.height()), where


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
    return 'bottom' if _setting(SIDE_KEY, 'right') == 'bottom' else 'right'


def saved_open():
    return _setting(OPEN_KEY, '1') != '0'


class _EdgeGrip(QWidget):
    """먼 쪽 가장자리를 끌어 AI 창 폭(오른쪽에 붙을 때)이나 높이(아래쪽)를 바꾼다."""

    def __init__(self, win):
        super().__init__(win)
        self._win = win
        self._start = None

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
        self.host = None
        self.side = saved_side()
        self.dock_size = self._saved_size(self.side)
        self._where = self.side
        self._worker = None
        self._history = []
        self._log = []                    # 대화 내용 (HTML 조각)
        self._q = ''
        self._pending = []
        self._undo = None
        self._build()
        self._apply_ai_status()
        from . import ai_connect
        ai_connect.on_change(self._on_ai_changed)

    # ── 화면 ──────────────────────────────────────────────────────
    def _build(self):
        self.setObjectName('AiChat')
        self.setStyleSheet(f"""
            QWidget#AiChat {{ background: {tc('#FFFFFF')}; border: 1px solid {tc('rgba(0,0,0,0.25)')}; }}
            QWidget {{ font-family: 'Pretendard', 'Malgun Gothic'; font-size: 10pt; color: {tc('#1E293B')}; }}
            QLabel {{ background: transparent; }}
            QFrame#AiHead {{ background: {tc('#F6F3E7')}; border: none;
                             border-bottom: 1px solid {tc('rgba(0,0,0,0.10)')}; }}
            QPushButton {{ background: {tc('#FFFFFF')}; border: 1px solid {tc('#D9D4C3')}; border-radius: 7px;
                           padding: 3px 10px; }}
            QPushButton:hover {{ background: {tc('#F4F1E6')}; }}
            QPushButton:disabled {{ color: {tc('#94A3B8')}; }}
            QPushButton#Icon {{ background: transparent; border: none; padding: 0; border-radius: 6px; }}
            QPushButton#Icon:hover {{ background: {tc('rgba(0,0,0,0.08)')}; }}
            QPushButton#AiStatus {{ background: transparent; border: none; padding: 2px 4px; color: {tc('#4F5F76')}; }}
            QPushButton#AiStatus:hover {{ color: {tc('#1E293B')}; }}
            QPushButton#Quick {{ border-radius: 10px; padding: 2px 9px; font-size: 9.5pt; }}
            QPushButton#Primary {{ background: {T.ACCENT}; border: 1px solid {T.ACCENT}; color: white; font-weight: 700; }}
            QTextBrowser {{ border: none; background: transparent; }}
            QLineEdit {{ background: {tc('#FFFFFF')}; border: 1px solid {tc('#D9D4C3')}; border-radius: 8px;
                         padding: 6px 8px; }}
            QLineEdit:focus {{ border: 1px solid {T.ACCENT}; }}
            QFrame#Proposal {{ background: {tc('#FFF8E1')}; border: 1px solid {tc('#F2D48A')}; border-radius: 8px; }}
        """)
        v = QVBoxLayout(self)
        v.setContentsMargins(1, 1, 1, 1)
        v.setSpacing(0)

        head = QFrame()
        head.setObjectName('AiHead')
        hl = QHBoxLayout(head)
        hl.setContentsMargins(10, 4, 4, 4)
        hl.setSpacing(2)
        title = QLabel('💬 AI')
        title.setStyleSheet('font-weight: 800;')
        self.btn_status = QPushButton()
        self.btn_status.setObjectName('AiStatus')
        self.btn_status.setCursor(Qt.PointingHandCursor)
        self.btn_status.setToolTip('AI 연결 바꾸기 (OpenCode · Claude · ChatGPT · Claude 구독)')
        self.btn_status.clicked.connect(self._ai_menu)
        self.btn_more = self._icon('ph.dots-three-bold', '붙는 위치 바꾸기')
        self.btn_more.clicked.connect(self._side_menu)
        self.btn_hide = self._icon('ph.x-bold', '숨기기 (업무창 제목줄의 💬로 다시 열기)')
        self.btn_hide.clicked.connect(self.hide_by_user)
        hl.addWidget(title)
        hl.addWidget(self.btn_status)
        hl.addStretch()
        hl.addWidget(self.btn_more)
        hl.addWidget(self.btn_hide)
        v.addWidget(head)

        body = QVBoxLayout()
        body.setContentsMargins(10, 8, 10, 10)
        body.setSpacing(6)
        self.view = QTextBrowser()
        self.view.setOpenExternalLinks(False)
        self.view.setOpenLinks(False)
        self.view.anchorClicked.connect(lambda _u: self._undo_last())
        body.addWidget(self.view, 1)

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

        quick = QHBoxLayout()
        quick.setSpacing(4)
        from task_ai import QUICK
        self.quick_buttons = []
        for label, q in QUICK:
            b = QPushButton(label)
            b.setObjectName('Quick')
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _, q=q: self.ask(q))
            quick.addWidget(b)
            self.quick_buttons.append(b)
        quick.addStretch()
        body.addLayout(quick)

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
        self.lbl_status.setStyleSheet(f'color: {tc("#64748B")};')
        self.lbl_status.hide()
        body.addWidget(self.lbl_status)
        v.addLayout(body, 1)

        self.grip = _EdgeGrip(self)
        self._show_welcome()

    def _icon(self, name, tip):
        b = QPushButton()
        b.setObjectName('Icon')
        b.setIcon(qta.icon(name, color=tc('#475569')))
        b.setIconSize(QSize(px(16), px(16)))
        b.setFixedSize(px(28), px(26))
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        return b

    def _show_welcome(self):
        self.view.setHtml(
            f'<p style="color:{tc("#64748B")}">업무에 대해 묻거나 시켜 보세요.<br>'
            '예: 이번 주 마감 알려줘 · 다음 주 금요일까지 공문 회신 추가해줘 · '
            '협의회 자료 마감을 수요일 3시로 바꿔줘</p>')

    # ── 붙이기 · 따라가기 ──────────────────────────────────────────
    @staticmethod
    def _saved_size(side):
        try:
            return max(MIN_SIZE[side], int(_setting(SIZE_KEY + '_' + side, DEFAULT_SIZE[side])))
        except ValueError:
            return DEFAULT_SIZE[side]

    def _host_rect(self):
        """붙을 기준: 업무창, 합체한 묶음이면 묶음 전체를 감싸는 사각형."""
        h = self.host
        group = [m for m in (getattr(h, '_link_group', None) or {h}) if m.isVisible()] or [h]
        r = group[0].frameGeometry()
        for m in group[1:]:
            r = r.united(m.frameGeometry())
        return r

    def follow(self):
        """업무창이 움직이거나 크기가 바뀌면 붙은 자리로 옮긴다."""
        if self.host is None or not self.isVisible():
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
        g = px(6)
        r = self.rect()
        geo = {'right': QRect(r.width() - g, 0, g, r.height()), 'left': QRect(0, 0, g, r.height()),
               'bottom': QRect(0, r.height() - g, r.width(), g), 'top': QRect(0, 0, r.width(), g)}[self._where]
        self.grip.set_side(self._where)
        self.grip.setGeometry(geo)
        self.grip.raise_()

    def set_dock_size(self, size):
        limit = (QApplication.primaryScreen().availableGeometry().width() // 2 if self.side == 'right'
                 else QApplication.primaryScreen().availableGeometry().height() // 2)
        self.dock_size = max(MIN_SIZE[self.side], min(int(size), max(MIN_SIZE[self.side], limit)))
        self.follow()

    def set_side(self, side):
        if side not in ('right', 'bottom') or side == self.side:
            return
        self.side = side
        _save(SIDE_KEY, side)
        self.dock_size = self._saved_size(side)
        self.follow()

    def attach(self, host, remember=True):
        """host 업무창 옆에 붙여 보인다."""
        if self.host is not None and self.host is not host:
            self._mark_host(False)                     # 전에 붙어 있던 업무창의 버튼은 끔
        self.host = host
        try:
            pin = bool(getattr(host, 'pin_active', False))
        except RuntimeError:
            return
        flags = Qt.FramelessWindowHint | Qt.Tool | (Qt.WindowStaysOnTopHint if pin else Qt.Widget)
        if self.windowFlags() != flags:
            self.setWindowFlags(flags)
        self.show()
        self.follow()
        self.raise_()
        self._mark_host(True)
        self._update_enabled()
        wid = getattr(host, 'window_id', None)
        if wid is not None:
            _save(HOST_KEY, wid)                       # 다음에 켤 때 이 업무창 옆에
        if remember:
            _save(OPEN_KEY, '1')

    def detach(self):
        self.host = None
        self.hide()

    def hide_by_user(self):
        _save(OPEN_KEY, '0')
        self.hide()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._place_grip()

    def showEvent(self, e):
        super().showEvent(e)
        self._mark_host(True)

    def hideEvent(self, e):
        super().hideEvent(e)
        self._mark_host(False)

    def _mark_host(self, on, host=None):
        """붙은 업무창 제목줄의 AI 버튼을 열림(옅은 바탕)/닫힘으로."""
        host = host or self.host
        try:
            bar = host.title_bar
            bar.set_active(bar.btn_ai, on)
            bar.btn_ai.setToolTip('AI 대화 숨기기' if on else 'AI 대화 열기')
        except (AttributeError, RuntimeError):
            pass

    # ── AI 연결 ───────────────────────────────────────────────────
    def _apply_ai_status(self):
        from . import ai_connect
        try:
            text, state = ai_connect.status()
        except Exception:
            text, state = 'AI 꺼짐 · 연결하기', 'off'
        color = {'on': '#16A34A', 'legacy': '#D97706'}.get(state, '#94A3B8')
        self.btn_status.setText(text)
        self.btn_status.setIcon(qta.icon('fa5s.circle', color=color))
        self.btn_status.setIconSize(QSize(8, 8))

    def _ai_ready(self):
        try:
            from ai_client import load_ai_mode
            return load_ai_mode() != 'none'
        except Exception:
            return False

    def _update_enabled(self):
        busy = self._worker is not None and self._worker.isRunning()
        ready = self._ai_ready()
        for w in [self.input, self.btn_send] + self.quick_buttons:
            w.setEnabled(ready and not busy)
        self.input.setPlaceholderText('AI에게 물어보기 (Enter)' if ready
                                      else 'AI가 꺼져 있어요 — 위의 「AI 꺼짐 · 연결하기」를 눌러 연결하세요')

    def _on_ai_changed(self):
        try:
            self._apply_ai_status()
            self._update_enabled()
        except RuntimeError:
            pass

    def _ai_menu(self):
        from . import ai_connect
        choice = ai_connect.show_menu(self, self.btn_status)
        if choice == 'none':
            ai_connect.turn_off()
        elif choice:
            ai_connect.connect(self, choice)

    def _side_menu(self):
        m = QMenu(self)
        for side, text in (('right', '오른쪽에 붙이기'), ('bottom', '아래쪽에 붙이기')):
            a = m.addAction(text, lambda s=side: self.set_side(s))
            a.setCheckable(True)
            a.setChecked(self.side == side)
        m.addSeparator()
        m.addAction('대화 지우기', self._clear_chat)
        m.exec_(self.btn_more.mapToGlobal(self.btn_more.rect().bottomLeft()))

    def _clear_chat(self):
        self._history = []
        self._log = []
        self._hide_proposal()
        self._show_welcome()

    # ── 묻기 ──────────────────────────────────────────────────────
    def ask(self, question):
        q = (question or '').strip()
        if not q or (self._worker is not None and self._worker.isRunning()) or self.host is None:
            return
        import task_ai
        tasks, labels = task_ai.task_windows_tasks()
        self._q = q
        self.input.clear()
        self._hide_proposal()
        self._append(f'<p><b>나</b> {html.escape(q)}</p>')
        self._set_status('AI가 생각하는 중…')
        self._worker = task_ai.TaskAskWorker(q, tasks, labels, getattr(self.host, 'window_id', None),
                                             self._history)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._update_enabled)
        self._update_enabled()
        self._worker.start()

    def _append(self, html_text):
        self._log = (self._log + [html_text])[-60:]
        self.view.setHtml(''.join(self._log))
        sb = self.view.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _set_status(self, text):
        self.lbl_status.setText(text)
        self.lbl_status.setVisible(bool(text))

    def _on_done(self, reply, masked):
        answer = reply.get('answer') or ''
        self._history = (self._history + [(self._q, answer)])[-2:]
        text = html.escape(answer).replace('\n', '<br>')
        if masked:
            try:
                from privacy import describe
                text += f'<br><span style="color:{tc("#64748B")}">🛡 개인정보 {html.escape(describe(masked))}을(를) 가리고 보냈어요.</span>'
            except Exception:
                pass
        self._append(f'<p><b>AI</b> {text}</p>')
        self._set_status('')
        self._show_proposal(reply.get('changes') or [], reply.get('rejected') or [])

    def _on_failed(self, msg):
        self._append(f'<p style="color:{tc("#BE123C")}"><b>AI</b> {html.escape(msg)}</p>')
        self._set_status('')

    # ── 제안 [넣기] · [되돌리기] ───────────────────────────────────
    def _show_proposal(self, changes, rejected=()):
        import task_ai
        self._pending = list(changes)
        if not changes and not rejected:
            self.proposal.hide()
            return
        lines = [f'• {html.escape(task_ai.describe(c))}' for c in changes]
        lines += [f'<span style="color:{tc("#94A3B8")}">• {html.escape(r["label"])} — {html.escape(r["reason"])}'
                  f' (넣을 수 없어요)</span>' for r in rejected]
        head = '이렇게 바꿀까요?' if changes else '넣을 수 있는 제안이 없어요'
        self.lbl_proposal.setText(f'<b>{head}</b><br>' + '<br>'.join(lines))
        self.btn_apply.setVisible(bool(changes))
        self.btn_skip.setText('취소' if changes else '닫기')
        self.proposal.show()

    def _hide_proposal(self):
        self._pending = []
        self.proposal.hide()

    def _refresh_task_windows(self):
        host = self.host
        for w in list(getattr(host, '_open_windows', None) or [host]):
            try:
                if not getattr(w, '_memo_mode', False):
                    w._refresh_tasks()
            except RuntimeError:
                pass

    def _apply(self):
        if not self._pending or self.host is None:
            return
        import task_ai
        n = len(self._pending)
        self._undo = task_ai.apply_changes(self._pending, self.host.window_id)
        self._hide_proposal()
        self._refresh_task_windows()
        self._append(f'<p style="color:{tc("#047857")}">✓ {n}건을 넣었어요. '
                     f'<a href="undo:1">되돌리기</a></p>')

    def _undo_last(self):
        if not self._undo:
            return
        import task_ai
        task_ai.undo_changes(self._undo)
        self._undo = None
        self._refresh_task_windows()
        self._append(f'<p style="color:{tc("#64748B")}">↩ 넣기 전으로 되돌렸어요.</p>')


# ── 앱 전체에 하나 ────────────────────────────────────────────────
_win = None


def instance(create=True):
    global _win
    if _win is None and create:
        _win = AiChatWindow()
    return _win


def is_attached_to(host):
    w = instance(create=False)
    return w is not None and w.host is host and w.isVisible()


def toggle(host):
    """💬 버튼: 이 업무창 옆에 열려 있으면 숨기고, 아니면 이 업무창 옆으로 연다."""
    w = instance()
    if w.host is host and w.isVisible():
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


def host_hidden(host):
    """업무창이 숨거나 접히거나 메모 모드가 되면 함께 숨는다 (열림 설정은 그대로)."""
    w = instance(create=False)
    if w is not None and w.host is host:
        w.hide()


def host_shown(host):
    """업무창이 다시 보이면, 열려 있던 AI 창도 다시 붙인다."""
    w = instance(create=False)
    if w is not None and w.host is host and saved_open() and not w.isVisible():
        w.attach(host, remember=False)


def host_closed(host):
    w = instance(create=False)
    if w is not None and w.host is host:
        w.detach()


def _saved_host_alive(host):
    """기억한 업무창이 아직 있고 펼쳐진 업무창(메모장 아님)인지 → 그 id, 아니면 None."""
    try:
        wid = int(_setting(HOST_KEY, '') or 0)
    except ValueError:
        return None
    if not wid:
        return None
    try:
        from db import get_all_windows
        alive = {w['id'] for w in get_all_windows() if not w.get('memo_mode') and not w.get('collapsed')}
    except Exception:
        return None
    return wid if wid in alive else None


def reset_layout():
    """「창 배치 기본값으로」: 오른쪽에 붙이기 · 기본 크기 (열림 여부·붙은 업무창은 그대로)."""
    for side in ('right', 'bottom'):
        _save(SIZE_KEY + '_' + side, DEFAULT_SIZE[side])
    _save(SIDE_KEY, 'right')
    w = instance(create=False)
    if w is not None:
        w.side = 'right'
        w.dock_size = DEFAULT_SIZE['right']
        w.follow()


def auto_open(host):
    """앱을 켰을 때 (처음엔 열림): 지난번에 붙어 있던 업무창 옆에, 그 창이 없으면 먼저 뜬 업무창 옆에."""
    if not saved_open() or instance(create=False) is not None:
        return
    want = _saved_host_alive(host)
    if want is not None and want != getattr(host, 'window_id', None):
        return                                         # 기억한 업무창이 스스로 붙임
    instance().attach(host, remember=False)
