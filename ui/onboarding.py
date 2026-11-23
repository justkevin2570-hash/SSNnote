"""첫 사용 안내 (코치마크) — 업무창 위에 단계별 말풍선.

- 처음 켰을 때 한 번만 (db 설정 'onboarding_done'). 이미 업무 · 지난 기록이 있는 PC는 건너뜀.
- 다시 보기: 「…」 → 도움말 → 사용법 다시 보기.
- 창을 어둡게 덮고 설명하는 곳만 뚫어 둔다 (뚫린 곳은 그대로 누르고 · 쓰고 · 파일을 놓을 수 있음).
- 1단계는 업무를 하나 넣어야 [다음] (업무가 하나도 없을 때만). Esc · [건너뛰기]는 안내를 끝낸다.
"""
from PyQt5.QtCore import QEvent, QPoint, QRect, QRectF, Qt, QTimer
from PyQt5.QtGui import QColor, QKeySequence, QPainter, QPainterPath, QPen, QRegion
from PyQt5.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QShortcut, QVBoxLayout, QWidget

DONE_KEY = 'onboarding_done'
STEPS = ('add', 'row', 'file', 'bottom', 'menu')

TEXT = {
    'add': ('업무 추가  (Alt+N)',
            '할 일을 하나 써 보세요.\n'
            '· Tab → 마감 날짜 → 시간 → 반복\n'
            '· Shift+Tab → 한 칸 뒤로\n'
            '· Enter → 업무 넣기'),
    'row': ('업무 다루기',
            '· ☆ 누르기 → ★ 중요 표시\n'
            '· 연필 아이콘 · Enter → 메모 쓰기\n'
            '· 더블클릭 · F2 → 이름 고치기\n'
            '· ⋮ → 완료 · 반복 · 삭제'),
    'file': ('파일 바로가기',
             '관련 파일을 이 줄에 끌어다 놓으면\n바로가기로 저장돼요.'),
    'file_done': ('파일 바로가기',
                  '저장됐어요!\n파일 이름을 누르면 바로 열려요.'),
    'bottom': ('아래 줄 버튼',
               '공문 · 예산 · 달력 · AI\n왼쪽부터 Alt+1 · 2 · 3 · 4'),
    'menu': ('메뉴',
             '색 · 글꼴 · 설정은 여기에 다 있어요.\n이 안내는 도움말 → 사용법 다시 보기'),
}


def is_done() -> bool:
    try:
        from db import get_setting
        return get_setting(DONE_KEY, '') == '1'
    except Exception:
        return True


def mark_done():
    try:
        from db import set_setting
        set_setting(DONE_KEY, '1')
    except Exception:
        pass


def has_history() -> bool:
    """이미 쓰던 PC인지 (업무나 지난 기록이 있으면)."""
    try:
        from db import _connect
        with _connect() as conn:
            return bool(conn.execute('SELECT EXISTS(SELECT 1 FROM tasks) OR '
                                     'EXISTS(SELECT 1 FROM task_history)').fetchone()[0])
    except Exception:
        return True


def maybe_start(win):
    """처음 켰을 때만 안내를 띄운다. 쓰던 PC면 띄우지 않고 본 것으로 기록."""
    if is_done():
        return None
    if has_history():
        mark_done()
        return None
    return start(win)


def start(win):
    old = getattr(win, '_tour', None)
    if old is not None:
        try:
            old.finish(mark=False)
        except RuntimeError:
            pass
    win._tour = TourOverlay(win)
    win._tour.show_step(0)
    return win._tour


class _Bubble(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        from ui.common import px
        self.setObjectName('TourBubble')
        self.setStyleSheet("""
            QFrame#TourBubble { background: #FFFFFF; border-radius: 10px; }
            QLabel { background: transparent; }
            QLabel#TourStep { color: #4F46E5; font-weight: 700; font-size: 8pt; }
            QLabel#TourTitle { color: #1E293B; font-weight: 700; font-size: 11pt; }
            QLabel#TourBody { color: #475569; }
            QPushButton#TourSkip { border: none; background: transparent; color: #94A3B8; padding: 4px 2px; }
            QPushButton#TourSkip:hover { color: #64748B; }
            QPushButton#TourNext { background: #4F46E5; color: white; font-weight: 700; border: none;
                                   border-radius: 6px; padding: 5px 14px; }
            QPushButton#TourNext:hover { background: #4338CA; }
            QPushButton#TourNext:disabled { background: #C7D2FE; }
        """)
        self.setMinimumWidth(px(272))
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 10, 14, 10)
        v.setSpacing(4)
        self.lbl_step = QLabel()
        self.lbl_step.setObjectName('TourStep')
        self.lbl_title = QLabel()
        self.lbl_title.setObjectName('TourTitle')
        self.lbl_body = QLabel()
        self.lbl_body.setObjectName('TourBody')
        self.lbl_body.setWordWrap(False)        # 줄바꿈은 글에 직접 (높이가 어긋나지 않게)
        v.addWidget(self.lbl_step)
        v.addWidget(self.lbl_title)
        v.addWidget(self.lbl_body)
        v.addSpacing(4)
        h = QHBoxLayout()
        self.btn_skip = QPushButton('건너뛰기')
        self.btn_skip.setObjectName('TourSkip')
        self.btn_next = QPushButton('다음')
        self.btn_next.setObjectName('TourNext')
        for b in (self.btn_skip, self.btn_next):
            b.setCursor(Qt.PointingHandCursor)
            b.setFocusPolicy(Qt.NoFocus)
        h.addWidget(self.btn_skip)
        h.addStretch(1)
        h.addWidget(self.btn_next)
        v.addLayout(h)


class TourOverlay(QWidget):
    """업무창 전체를 덮는 반투명 막 + 말풍선. 설명하는 곳은 뚫려 있어 그대로 쓸 수 있다."""

    PAD = 4

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.step = 0
        self.task_id = None            # 1단계에서 넣은 업무 (없으면 맨 위 업무)
        self._attach_base = None       # 3단계 시작 때 붙임파일 개수
        self._hole = QRect()
        self.bubble = _Bubble(self)
        self.bubble.btn_skip.clicked.connect(lambda: self.finish())
        self.bubble.btn_next.clicked.connect(self.next)
        self._timer = QTimer(self)     # 줄이 다시 그려지거나 창 크기가 바뀌어도 따라가게
        self._timer.timeout.connect(self._tick)
        self._timer.start(250)
        win.installEventFilter(self)
        self._esc = QShortcut(QKeySequence(Qt.Key_Escape), win)
        self._esc.activated.connect(lambda: self.finish())
        self.setGeometry(win.rect())
        self.raise_()
        self.show()

    # ── 대상 찾기 ──
    def _rect(self, w):
        if w is None or not w.isVisible():
            return QRect()
        return QRect(w.mapTo(self.win, QPoint(0, 0)), w.size())

    def _row(self):
        from ui.task_row import TaskRow
        rows = [r for r in self.win.findChildren(TaskRow) if r.isVisible()]
        for r in rows:
            if r.task.get('id') == self.task_id:
                return r
        if rows:
            rows.sort(key=lambda r: r.mapTo(self.win, QPoint(0, 0)).y())
            return rows[0]
        return None

    def _target(self):
        key = STEPS[self.step]
        w = self.win
        if key == 'add':
            r = self._rect(w.input_name)
            dr = getattr(w, 'date_row_widget', None)
            if dr is not None and dr.isVisible():
                r = r.united(self._rect(dr))
            return r
        if key in ('row', 'file'):
            return self._rect(self._row())
        if key == 'bottom':
            return self._rect(w.btn_ai_bottom.parentWidget())
        if key == 'menu':
            return self._rect(w.title_bar.btn_menu)
        return QRect()

    def _has_tasks(self):
        return self._row() is not None

    # ── 단계 ──
    def show_step(self, i):
        self.step = i
        key = STEPS[i]
        if key == 'file':
            row = self._row()
            self._attach_base = len(getattr(row, '_attachments', []) or []) if row else 0
        b = self.bubble
        b.lbl_step.setText(f'{i + 1} / {len(STEPS)}')
        b.lbl_title.setText(TEXT[key][0])
        b.lbl_body.setText(TEXT[key][1])
        b.btn_next.setText('시작하기' if i == len(STEPS) - 1 else '다음')
        self._hole = QRect(-1, -1, 0, 0)                  # 말풍선 크기 · 자리 다시 맞춤
        if key == 'add':
            b.btn_next.setEnabled(self._has_tasks())
            self.win.input_name.setFocus()
        else:
            b.btn_next.setEnabled(True)
        self._tick()

    def next(self):
        if self.step >= len(STEPS) - 1:
            self.finish()
        else:
            self.show_step(self.step + 1)

    def on_task_added(self, task_id):
        """업무창이 업무를 넣은 뒤 부른다 — 1단계면 다음으로."""
        self.task_id = task_id
        if STEPS[self.step] == 'add':
            QTimer.singleShot(0, lambda: self.show_step(1))

    def finish(self, mark=True):
        if mark:
            mark_done()
        self._timer.stop()
        try:
            self._esc.setEnabled(False)
            self._esc.deleteLater()
        except RuntimeError:
            pass
        try:
            self.win.removeEventFilter(self)
        except RuntimeError:
            pass
        if getattr(self.win, '_tour', None) is self:
            self.win._tour = None
        self.hide()
        self.deleteLater()

    # ── 위치 · 그리기 ──
    def _tick(self):
        if not self.isVisible():
            return
        if self.geometry() != self.win.rect():
            self.setGeometry(self.win.rect())
        key = STEPS[self.step]
        if key == 'add' and not self.bubble.btn_next.isEnabled() and self._has_tasks():
            self.bubble.btn_next.setEnabled(True)
        if key == 'file':
            row = self._row()
            n = len(getattr(row, '_attachments', []) or []) if row else 0
            if (self._attach_base is not None and n > self._attach_base
                    and self.bubble.lbl_body.text() != TEXT['file_done'][1]):
                self.bubble.lbl_title.setText(TEXT['file_done'][0])
                self.bubble.lbl_body.setText(TEXT['file_done'][1])
                self._hole = QRect(-1, -1, 0, 0)          # 말풍선 크기 · 자리 다시 맞춤
        r = self._target()
        hole = r.adjusted(-self.PAD, -self.PAD, self.PAD, self.PAD) if not r.isEmpty() else QRect()
        hole = hole.intersected(self.rect().adjusted(1, 1, -1, -1)) if not hole.isEmpty() else hole
        if hole != self._hole or not self.bubble.isVisible():
            self._hole = hole
            region = QRegion(self.rect())
            if not hole.isEmpty():
                region = region.subtracted(QRegion(hole))
            self.setMask(region)
            self._place_bubble()
            self.update()
        self.raise_()

    def _place_bubble(self):
        b = self.bubble
        b.adjustSize()
        bw, bh = b.width(), b.height()
        W, H = self.width(), self.height()
        hole = self._hole if not self._hole.isEmpty() else QRect(W // 2, H // 2, 0, 0)
        gap = 10
        if hole.bottom() + gap + bh <= H - 4:
            y = hole.bottom() + gap
        elif hole.top() - gap - bh >= 4:
            y = hole.top() - gap - bh
        else:
            y = max(4, (H - bh) // 2)
        x = hole.center().x() - bw // 2
        x = max(8, min(x, W - bw - 8))
        b.move(x, y)
        b.show()
        b.raise_()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        full = QPainterPath()
        full.addRect(QRectF(self.rect()))
        if not self._hole.isEmpty():
            hole = QPainterPath()
            hole.addRoundedRect(QRectF(self._hole), 8, 8)
            full = full.subtracted(hole)
        p.fillPath(full, QColor(15, 23, 42, 140))
        if not self._hole.isEmpty():
            p.setPen(QPen(QColor('#818CF8'), 2))
            p.setBrush(Qt.NoBrush)
            p.drawRoundedRect(QRectF(self._hole).adjusted(-1, -1, 1, 1), 9, 9)   # 뚫린 곳 바로 바깥

    def mousePressEvent(self, e):          # 어두운 곳을 눌러도 아무 일 없음
        e.accept()

    def eventFilter(self, obj, event):
        if obj is self.win:
            t = event.type()
            if t == QEvent.Resize:
                QTimer.singleShot(0, self._tick)
        return False
