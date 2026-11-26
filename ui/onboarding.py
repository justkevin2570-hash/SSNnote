"""첫 사용 안내 — 게임 튜토리얼처럼 하나씩 직접 해 보기 (유주가 안내).

- 처음 켰을 때 한 번만 (db 설정 'onboarding_done'). 이미 업무 · 지난 기록이 있는 PC는 건너뜀.
  다시 보기: 「…」 → 도움말 → 사용법 다시 보기.
- 6단계: 업무 추가 · 업무 다루기 · 파일 바로가기 · 아래 줄 버튼 · AI 연결 · 메뉴 → 클리어 화면.
  단계마다 할 일(퀘스트)을 직접 하면 ✓, 다 하면 「✓ 완료!」 뒤 다음 단계로. [다음] 단추는 없다.
- 창을 어둡게 덮고 할 곳만 뚫어 빛나는 테두리로 (뚫린 곳은 그대로 누르고 · 쓰고 · 파일을 놓을 수 있음).
- 「이 단계 건너뛰기」 · 「그만두기(Esc)」. 연습으로 만든 업무는 그대로 둔다.
"""
import math
import random

from PyQt5.QtCore import QEvent, QPoint, QPointF, QRect, QRectF, Qt, QTimer
from PyQt5.QtGui import (QBrush, QColor, QFont, QKeySequence, QLinearGradient, QPainter, QPainterPath,
                         QPen, QPixmap, QRadialGradient, QRegion)
from PyQt5.QtWidgets import QApplication, QMenu, QPushButton, QShortcut, QWidget

DONE_KEY = 'onboarding_done'
STEPS = ('add', 'row', 'file', 'bottom', 'ai', 'menu')
STEP_TITLES = {'add': '업무 추가하기', 'row': '업무 다루기', 'file': '파일 바로가기',
               'bottom': '아래 줄 버튼', 'ai': 'AI 연결', 'menu': '메뉴'}
BADGES = ('업무 추가', '업무 다루기', '파일 바로가기', '아래 줄 버튼', 'AI 연결', '메뉴')
YUJU = {'add': 6, 'row': 5, 'file': 1, 'bottom': 3, 'ai': 4, 'menu': 2, 'clear': 1}
AI_BENEFITS = ('📝  제목만 쓰면 공문 초안을 써 줘요 (내 공문 모양대로)',
               '📸  캡처한 공문의 제목 · 번호를 정확하게 읽어요',
               '📊  예산 화면을 캡처하면 항목별로 정리해 줘요',
               '💬  「이번 주 마감 뭐야?」 물어보면 알려 줘요')
BURST_MS = 900

# 색 (남보라 유리 카드)
_C_BG1, _C_BG2 = QColor(30, 27, 75, 246), QColor(49, 46, 129, 246)
_C_EDGE = QColor(165, 180, 252, 130)
_C_GOLD, _C_GREEN, _C_LAV = QColor('#FDE68A'), QColor('#34D399'), QColor('#A5B4FC')


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


def _yuju_pixmap(n, height, dpr):
    from .common import _base_path
    import os
    pm = QPixmap(os.path.join(_base_path(), 'assets', 'popup_images', f'yuju_{n}.png'))
    if pm.isNull():
        return pm
    pm = pm.scaledToHeight(int(height * dpr), Qt.SmoothTransformation)
    pm.setDevicePixelRatio(dpr)
    return pm


def _font(base: QFont, pt, bold=False):
    f = QFont(base)
    f.setPointSizeF(pt)
    f.setBold(bold)
    return f


class _LinkButton(QPushButton):
    def __init__(self, text, parent, primary=False):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.NoFocus)
        if primary:
            self.setStyleSheet("""
                QPushButton { color: white; font-weight: 700; border: none; border-radius: 14px; padding: 6px 16px; font-size: 10pt;
                              background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366F1, stop:1 #A855F7); }
                QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4F46E5, stop:1 #9333EA); }""")
        else:
            self.setStyleSheet("""
                QPushButton { color: rgba(199,210,254,170); border: none; background: transparent; padding: 4px 2px; font-size: 9pt; }
                QPushButton:hover { color: #E0E7FF; text-decoration: underline; }""")


class TourOverlay(QWidget):
    """업무창을 덮는 어두운 막 + 빛나는 구멍 + 퀘스트 카드 + 유주. 직접 해야 다음으로 간다."""

    PAD = 4

    def __init__(self, win):
        super().__init__(win)
        from .common import px
        self.px = px
        self.win = win
        self.step = 0
        self.task_id = None
        self.done = []                  # 이번 단계 퀘스트별 완료 여부
        self.completed = set()          # 끝낸 단계 (클리어 화면 배지)
        self.clear = False
        self._burst = None              # 완료 연출 시작 시각(ms) — 그동안은 입력 대신 연출
        self._t = 0                     # 애니메이션 시계(ms)
        self._line, self._typed = '', 0
        self._hole = QRect()
        self._card = QRect()
        self._yuju_n = None
        self._yuju_pm = QPixmap()
        self._attach_base = 0
        self._bottom_seen_open = False
        self._menu_seen = False
        self._confetti = []

        self.btn_skip = _LinkButton('이 단계 건너뛰기', self)
        self.btn_quit = _LinkButton('그만두기 (Esc)', self)
        self.btn_action = _LinkButton('', self, primary=True)
        self.btn_later = _LinkButton('', self)
        self.btn_skip.clicked.connect(self.skip_step)
        self.btn_quit.clicked.connect(lambda: self.finish())
        self.btn_action.clicked.connect(self._on_action)
        self.btn_later.clicked.connect(self.skip_step)

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(33)
        self._esc = QShortcut(QKeySequence(Qt.Key_Escape), win)
        self._esc.activated.connect(lambda: self.finish())
        win.installEventFilter(self)
        QApplication.instance().installEventFilter(self)      # ⋮ · 「…」 메뉴가 열렸는지
        self.setGeometry(win.rect())
        self.raise_()
        self.show()

    # ── 대상 ──────────────────────────────────────────────────
    def _rect(self, w):
        if w is None:
            return QRect()
        try:
            if not w.isVisible():
                return QRect()
            return QRect(w.mapTo(self.win, QPoint(0, 0)), w.size())
        except RuntimeError:
            return QRect()

    def _row(self):
        from .task_row import TaskRow
        rows = [r for r in self.win.findChildren(TaskRow) if r.isVisible()]
        for r in rows:
            if r.task.get('id') == self.task_id:
                return r
        if rows:
            rows.sort(key=lambda r: r.mapTo(self.win, QPoint(0, 0)).y())
            return rows[0]
        return None

    def _key(self):
        return STEPS[self.step]

    def _bottom_first(self):
        from .bottom_bar import load_order, LABELS
        k = load_order()[0]
        return k, LABELS.get(k, k)

    def _bottom_open(self, key):
        w = self.win
        if key == 'ai':
            return bool(w.btn_ai_bottom.property('active'))
        bw = w._bottom_window(key)
        try:
            return bw is not None and bw.isVisible() and not bw.isMinimized()
        except RuntimeError:
            return False

    def _target(self):
        if self.clear:
            return QRect()
        key, w = self._key(), self.win
        if key == 'add':
            r = self._rect(w.input_name)
            dr = getattr(w, 'date_row_widget', None)
            if dr is not None and dr.isVisible():
                r = r.united(self._rect(dr))
            return r
        if key in ('row', 'file'):
            row = self._row()
            sc = getattr(w, 'task_scroll', None)
            if row is not None and sc is not None and getattr(self, '_scrolled', None) != self.step:
                sc.ensureWidgetVisible(row, 0, 8)
                self._scrolled = self.step
            r = self._rect(row)
            ne = getattr(row, 'note_editor', None) if row is not None else None
            if ne is not None and ne.isVisible():
                r = r.united(self._rect(ne))
            return r
        if key == 'bottom':
            return self._rect(w.btn_ai_bottom.parentWidget())
        if key == 'menu':
            return self._rect(w.title_bar.btn_menu)
        return QRect()

    # ── 퀘스트 ────────────────────────────────────────────────
    def quests(self):
        key = self._key()
        if key == 'add':
            return ['할 일 쓰기', 'Tab 눌러 마감일 칸으로 가기', 'Enter 눌러 업무 넣기']
        if key == 'row':
            return ['☆ 눌러 중요 표시하기', '연필 아이콘 눌러 메모 열기', '⋮ 눌러 메뉴 열어 보기']
        if key == 'file':
            return ['관련 파일을 이 줄에 끌어다 놓기']
        if key == 'bottom':
            _k, label = self._bottom_first()
            return [f'Alt+1 눌러 {label} 열기', 'Alt+1 한 번 더 눌러 닫기']
        if key == 'ai':
            return ['AI 연결하기']
        return ['「…」 눌러 메뉴 열어 보기']

    def _lines(self):
        key = self._key()
        if key == 'add':
            return ['할 일을 하나 써 볼까요?', '좋아요! 이제 Tab을 눌러 볼까요?', '마지막! Enter를 눌러 넣어요.']
        if key == 'row':
            return ['별을 눌러 중요 표시를 해 봐요!', '이번엔 연필 아이콘으로 메모를 열어요.', '⋮ 단추도 눌러 볼까요?']
        if key == 'file':
            return ['관련 파일을 이 줄에 끌어다 놓아 봐요!']
        if key == 'bottom':
            _k, label = self._bottom_first()
            return [f'Alt+1을 눌러 {label}을(를) 열어 봐요!', '한 번 더 누르면 닫혀요!']
        if key == 'ai':
            from ai_client import load_ai_mode
            if load_ai_mode() != 'none':
                return ['벌써 AI가 연결돼 있네요! 👍']
            return ['AI를 연결하면 공문 쓰기가 훨씬 편해져요!']
        return ['마지막! 「…」을 눌러 메뉴를 열어 봐요.']

    def _set_line(self, text):
        if text != self._line:
            self._line, self._typed = text, 0

    def show_step(self, i):
        self.step = i
        self.clear = False
        self._burst = None
        self.done = [False] * len(self.quests())
        key = self._key()
        self._scrolled = None
        self._bottom_seen_open = False
        self._menu_seen = False
        if key == 'file':
            row = self._row()
            self._attach_base = len(getattr(row, '_attachments', []) or []) if row else 0
        if key == 'add':
            self.win.input_name.setFocus()
        self._set_yuju(YUJU[key])
        self._line = None
        self._update_line()
        self._layout_buttons()
        self._hole = QRect(-1, -1, 0, 0)
        self._tick()

    def _update_line(self):
        lines = self._lines()
        nxt = next((n for n, d in enumerate(self.done) if not d), len(self.done) - 1)
        self._set_line(lines[min(nxt, len(lines) - 1)])

    def on_task_added(self, task_id):
        """업무창이 업무를 넣은 뒤 부른다."""
        self.task_id = task_id
        if not self.clear and self._key() == 'add':
            self.done[0] = True
            if self.done[1]:
                self.done[2] = True
            else:
                self._set_line('좋아요! 이번엔 Tab도 눌러서 하나 더 넣어 볼까요?')

    def _check(self):
        """지금 단계의 퀘스트를 살펴 ✓를 켠다 (순서대로)."""
        if self.clear or self._burst is not None:
            return
        key, w, d = self._key(), self.win, self.done
        if key == 'add':
            if w.input_name.text().strip():
                d[0] = True
            fw = QApplication.focusWidget()
            date = getattr(w, 'input_date', None)
            if d[0] and fw is not None and date is not None and (fw is date or date.isAncestorOf(fw)):
                d[1] = True
        elif key == 'row':
            row = self._row()
            if row is not None:
                if self._task_priority():
                    d[0] = True
                if d[0] and getattr(row, '_expanded', False):
                    d[1] = True
                if d[1] and self._menu_seen:
                    d[2] = True
        elif key == 'file':
            row = self._row()
            n = len(getattr(row, '_attachments', []) or []) if row else 0
            if n > self._attach_base:
                d[0] = True
        elif key == 'bottom':
            k, _label = self._bottom_first()
            is_open = self._bottom_open(k)
            if is_open:
                d[0] = True
                self._bottom_seen_open = True
            elif self._bottom_seen_open:
                d[1] = True
        elif key == 'ai':
            from ai_client import load_ai_mode
            if load_ai_mode() != 'none':
                d[0] = True
        elif key == 'menu':
            if self._menu_seen:
                d[0] = True
        if not all(d):
            self._update_line()
        else:
            self._complete()

    def _task_priority(self):
        try:
            from db import get_tasks
            for t in get_tasks(self.win.window_id):
                if t['id'] == (self.task_id or (self._row().task.get('id') if self._row() else None)):
                    return bool(t.get('priority'))
        except Exception:
            pass
        return False

    def _complete(self):
        self.completed.add(self._key())
        self._burst = self._t
        self._set_line('완벽해요! 다음으로 가 볼까요?' if self.step < len(STEPS) - 1 else '모두 해냈어요!')
        self._set_yuju(1)

    def skip_step(self):
        self._next()

    def _next(self):
        if self.step >= len(STEPS) - 1:
            self._show_clear()
        else:
            self.show_step(self.step + 1)

    def _show_clear(self):
        self.clear = True
        self._burst = None
        self._set_yuju(YUJU['clear'])
        self._set_line('')
        rnd = random.Random(7)
        self._confetti = [(rnd.random(), rnd.random(), rnd.choice(['#FDE68A', '#A5B4FC', '#F9A8D4', '#86EFAC']),
                           rnd.uniform(1.5, 3.6), rnd.uniform(0, math.tau)) for _ in range(34)]
        self._layout_buttons()
        self._hole = QRect(-1, -1, 0, 0)
        self._tick()

    def _on_action(self):
        if self.clear:
            self.finish()
            return
        if self._key() == 'ai':
            try:
                from .ai_settings_window import open_ai_settings
                open_ai_settings(self.win)
            except Exception:
                pass

    def finish(self, mark=True):
        if mark:
            mark_done()
        self._timer.stop()
        for obj in (self.win, QApplication.instance()):
            try:
                obj.removeEventFilter(self)
            except RuntimeError:
                pass
        try:
            self._esc.setEnabled(False)
            self._esc.deleteLater()
        except RuntimeError:
            pass
        if getattr(self.win, '_tour', None) is self:
            self.win._tour = None
        self.hide()
        self.deleteLater()

    # ── 배치 ──────────────────────────────────────────────────
    def _card_height(self):
        px = self.px
        if self.clear:
            return px(330)
        h = px(78) + len(self.done) * px(28) + px(34)
        if self._key() == 'ai' and not self.done[0]:
            h += len(AI_BENEFITS) * px(22) + px(8)
        return h

    def _layout_buttons(self):
        ai = (not self.clear) and self._key() == 'ai'
        from ai_client import load_ai_mode
        ai_needed = ai and load_ai_mode() == 'none'
        self.btn_skip.setVisible(not self.clear and not ai_needed)
        self.btn_quit.setVisible(not self.clear)
        self.btn_action.setVisible(self.clear or ai_needed)
        self.btn_later.setVisible(ai_needed or ((not self.clear) and self._key() == 'file'))
        self.btn_action.setText('시작하기' if self.clear else 'AI 연결하기')
        self.btn_later.setText('나중에' if ai_needed else '파일이 없어요 · 다음에 할게요')
        for b in (self.btn_skip, self.btn_quit, self.btn_action, self.btn_later):
            b.adjustSize()

    def _place(self):
        px = self.px
        W, H = self.width(), self.height()
        ch = self._card_height()
        cw = min(W - px(24), px(440))
        if self.clear:
            self._card = QRect((W - cw) // 2, max(px(8), (H - ch) // 2), cw, ch)
        else:
            yuju_h = px(84)
            hole = self._hole if not self._hole.isEmpty() else QRect(0, 0, 0, 0)
            below_top = hole.bottom() + px(12) if not hole.isEmpty() else px(8)
            bar = self._rect(self.win.btn_ai_bottom.parentWidget())
            limit = (bar.top() - px(8)) if (not bar.isEmpty() and self._key() != 'bottom') else H - px(8)
            if self._key() == 'bottom':
                limit = hole.top() - px(8)
            y = limit - ch
            if y < below_top + yuju_h:                       # 아래 자리가 모자라면 구멍 위쪽에
                y = max(px(8) + yuju_h, min(y, H - ch - px(8)))
            self._card = QRect((W - cw) // 2, y, cw, ch)
        c = self._card
        m = px(16)
        if self.clear:
            self.btn_action.move(c.center().x() - self.btn_action.width() // 2, c.bottom() - px(14) - self.btn_action.height())
            return
        yb = c.bottom() - px(10) - self.btn_skip.height()
        self.btn_skip.move(c.left() + m, yb)
        self.btn_quit.move(c.left() + m + self.btn_skip.width() + px(10) if self.btn_skip.isVisible() else c.left() + m, yb)
        if self.btn_action.isVisible():
            self.btn_action.move(c.right() - m - self.btn_action.width(), c.bottom() - px(12) - self.btn_action.height())
        if self.btn_later.isVisible():
            if self.btn_action.isVisible():
                self.btn_later.move(self.btn_action.x() - px(8) - self.btn_later.width(), self.btn_action.y() + (self.btn_action.height() - self.btn_later.height()) // 2)
            else:
                self.btn_later.move(c.right() - m - self.btn_later.width(), yb)

    def _set_yuju(self, n):
        size = self.px(130 if self.clear else 88)
        if (n, size) != self._yuju_n:
            self._yuju_n = (n, size)
            dpr = self.devicePixelRatioF() or 1.0
            self._yuju_pm = _yuju_pixmap(n, size, dpr)

    # ── 시계 ──────────────────────────────────────────────────
    def _tick(self):
        if not self.isVisible():
            return
        self._t += 33
        if self.geometry() != self.win.rect():
            self.setGeometry(self.win.rect())
        if self._line and self._typed < len(self._line):
            self._typed = min(len(self._line), self._typed + 1)
        if self._burst is not None and self._t - self._burst >= BURST_MS:
            self._burst = None
            self._next()
            return
        if self._t % 198 < 33:                       # 퀘스트 확인은 0.2초마다
            self._check()
            if self._key() == 'ai':
                self._layout_buttons()
        r = self._target()
        hole = r.adjusted(-self.PAD, -self.PAD, self.PAD, self.PAD) if not r.isEmpty() else QRect()
        if not hole.isEmpty():
            hole = hole.intersected(self.rect().adjusted(1, 1, -1, -1))
        if hole != self._hole:
            self._hole = hole
            region = QRegion(self.rect())
            if not hole.isEmpty():
                region = region.subtracted(QRegion(hole))
            self.setMask(region)
        self._place()
        for b in (self.btn_skip, self.btn_quit, self.btn_action, self.btn_later):
            b.raise_()
        self.raise_()
        self.update()

    # ── 그리기 ────────────────────────────────────────────────
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        px = self.px
        full = QPainterPath()
        full.addRect(QRectF(self.rect()))
        if not self._hole.isEmpty():
            hp = QPainterPath()
            hp.addRoundedRect(QRectF(self._hole), 10, 10)
            full = full.subtracted(hp)
        p.fillPath(full, QColor(8, 10, 30, 188))
        if not self._hole.isEmpty():
            pulse = 0.55 + 0.45 * math.sin(self._t / 420.0)
            for i, a in enumerate((70, 45, 26, 12)):
                p.setPen(QPen(QColor(129, 140, 248, int(a * pulse)), 2 + i * 3))
                p.setBrush(Qt.NoBrush)
                p.drawRoundedRect(QRectF(self._hole).adjusted(-1 - i, -1 - i, 1 + i, 1 + i), 10 + i, 10 + i)
            p.setPen(QPen(_C_LAV, 2))
            p.drawRoundedRect(QRectF(self._hole).adjusted(-1, -1, 1, 1), 10, 10)
        c = QRectF(self._card)
        g = QLinearGradient(c.topLeft(), c.bottomRight())
        g.setColorAt(0, _C_BG1)
        g.setColorAt(1, _C_BG2)
        p.setPen(QPen(_C_EDGE, 1.2))
        p.setBrush(QBrush(g))
        p.drawRoundedRect(c, 14, 14)
        base = self.font()
        if self.clear:
            self._paint_clear(p, c, base)
            return
        m = px(16)
        # 진행 막대
        p.setFont(_font(base, 8, True))
        p.setPen(_C_LAV)
        p.drawText(QRectF(c.left() + m, c.top() + px(10), px(90), px(16)), Qt.AlignLeft | Qt.AlignVCenter,
                   f'STEP {self.step + 1} / {len(STEPS)}')
        sx = c.left() + px(98)
        sw = (c.right() - m - sx - 5 * px(4)) / len(STEPS)
        for i in range(len(STEPS)):
            r = QRectF(sx + i * (sw + px(4)), c.top() + px(15), sw, px(6))
            col = QColor('#818CF8') if i < self.step else (_C_GOLD if i == self.step else QColor(255, 255, 255, 40))
            p.setPen(Qt.NoPen)
            p.setBrush(col)
            p.drawRoundedRect(r, 3, 3)
        p.setFont(_font(base, 13, True))
        p.setPen(QColor('white'))
        p.drawText(QRectF(c.left() + m, c.top() + px(30), c.width() - 2 * m, px(26)), Qt.AlignLeft | Qt.AlignVCenter,
                   STEP_TITLES[self._key()])
        y0 = c.top() + px(64)
        if self._key() == 'ai' and not self.done[0]:
            p.setFont(_font(base, 9.5))
            p.setPen(QColor(224, 231, 255, 230))
            for i, line in enumerate(AI_BENEFITS):
                p.drawText(QRectF(c.left() + m, y0 + i * px(22), c.width() - 2 * m, px(20)),
                           Qt.AlignLeft | Qt.AlignVCenter, line)
            y0 += len(AI_BENEFITS) * px(22) + px(8)
        now = next((n for n, d in enumerate(self.done) if not d), -1)
        for i, (txt, ok) in enumerate(zip(self.quests(), self.done)):
            yy = y0 + i * px(28)
            ctr = QPointF(c.left() + m + px(9), yy + px(10))
            rr = px(8)
            if ok:
                p.setPen(Qt.NoPen)
                p.setBrush(_C_GREEN)
                p.drawEllipse(ctr, rr, rr)
                p.setPen(QPen(QColor('#064E3B'), 2))
                p.drawPolyline(QPointF(ctr.x() - rr * 0.45, ctr.y()), QPointF(ctr.x() - rr * 0.1, ctr.y() + rr * 0.4),
                               QPointF(ctr.x() + rr * 0.5, ctr.y() - rr * 0.38))
                p.setPen(QColor(209, 213, 219, 170))
                p.setFont(_font(base, 10))
            elif i == now:
                p.setPen(QPen(_C_GOLD, 2))
                p.setBrush(QColor(253, 230, 138, int(30 + 30 * (0.5 + 0.5 * math.sin(self._t / 300.0)))))
                p.drawEllipse(ctr, rr, rr)
                p.setPen(_C_GOLD)
                p.setFont(_font(base, 10, True))
            else:
                p.setPen(QPen(QColor(255, 255, 255, 90), 1.5))
                p.setBrush(Qt.NoBrush)
                p.drawEllipse(ctr, rr, rr)
                p.setPen(QColor(255, 255, 255, 140))
                p.setFont(_font(base, 10))
            p.drawText(QRectF(c.left() + m + px(26), yy, c.width() - 2 * m - px(26), px(20)),
                       Qt.AlignLeft | Qt.AlignVCenter, txt)
        self._paint_yuju(p, c, base)
        if self._burst is not None:
            self._paint_burst(p, c, base)

    def _paint_yuju(self, p, c, base):
        px = self.px
        pm = self._yuju_pm
        yh = pm.height() / (pm.devicePixelRatio() or 1) if not pm.isNull() else px(88)
        bounce = 0
        if self._burst is not None:
            bounce = -abs(math.sin((self._t - self._burst) / 120.0)) * px(10)
        if not pm.isNull():
            p.drawPixmap(QPointF(c.left() + px(4), c.top() - yh + px(2) + bounce), pm)
        if self._line:
            bub = QRectF(c.left() + px(84), c.top() - px(62), c.width() - px(96), px(44))
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(255, 255, 255, 250))
            p.drawRoundedRect(bub, 12, 12)
            tail = QPainterPath()
            tail.moveTo(bub.left(), bub.bottom() - px(18))
            tail.lineTo(bub.left() - px(9), bub.bottom() - px(10))
            tail.lineTo(bub.left(), bub.bottom() - px(8))
            p.fillPath(tail, QColor('white'))
            p.setFont(_font(base, 10, True))
            p.setPen(QColor('#1E1B4B'))
            p.drawText(bub.adjusted(px(14), 0, -px(10), 0), Qt.AlignLeft | Qt.AlignVCenter | Qt.TextWordWrap,
                       self._line[:self._typed])

    def _paint_burst(self, p, c, base):
        px = self.px
        k = min(1.0, (self._t - self._burst) / 260.0)            # 0 → 1 (커졌다가)
        scale = 0.6 + 0.5 * k - 0.1 * max(0.0, (self._t - self._burst - 260) / 300.0)
        fade = 1.0 - max(0.0, (self._t - self._burst - 600) / 300.0)
        ctr = QPointF(c.center().x() + px(40), c.top() + px(100))
        rg = QRadialGradient(ctr, px(90))
        rg.setColorAt(0, QColor(52, 211, 153, int(120 * fade)))
        rg.setColorAt(1, QColor(52, 211, 153, 0))
        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(rg))
        p.drawEllipse(ctr, px(90) * scale, px(60) * scale)
        for i in range(12):
            a = i * math.pi / 6
            r1, r2 = px(30) * scale, px(44) * scale
            p.setPen(QPen(QColor(253, 230, 138, int(255 * fade)), 2.5))
            p.drawLine(QPointF(ctr.x() + math.cos(a) * r1 * 1.6, ctr.y() + math.sin(a) * r1),
                       QPointF(ctr.x() + math.cos(a) * r2 * 1.6, ctr.y() + math.sin(a) * r2))
        p.setFont(_font(base, 20 * scale, True))
        p.setPen(QColor(52, 211, 153, int(255 * fade)))
        p.drawText(QRectF(ctr.x() - px(100), ctr.y() - px(24), px(200), px(48)), Qt.AlignCenter, '✓ 완료!')

    def _paint_clear(self, p, c, base):
        px = self.px
        for fx, fy, col, r, ph in self._confetti:
            x = c.left() + px(12) + fx * (c.width() - px(24))
            y = c.top() + px(8) + ((fy * px(120) + self._t * 0.02) % px(130))
            q = QColor(col)
            q.setAlphaF(0.55 + 0.45 * math.sin(self._t / 250.0 + ph))
            p.setPen(Qt.NoPen)
            p.setBrush(q)
            p.drawEllipse(QPointF(x, y), r, r)
        pm = self._yuju_pm
        if not pm.isNull():
            w = pm.width() / (pm.devicePixelRatio() or 1)
            bob = math.sin(self._t / 300.0) * px(3)
            p.drawPixmap(QPointF(c.center().x() - w / 2, c.top() + px(10) + bob), pm)
        p.setFont(_font(base, 17, True))
        p.setPen(_C_GOLD)
        p.drawText(QRectF(c.left(), c.top() + px(142), c.width(), px(32)), Qt.AlignCenter, '튜토리얼 클리어!')
        p.setFont(_font(base, 9.5))
        p.setPen(QColor('#C7D2FE'))
        p.drawText(QRectF(c.left(), c.top() + px(176), c.width(), px(20)), Qt.AlignCenter, '이제 서서니 노트를 마음껏 써 보세요')
        bx, by = c.left() + px(18), c.top() + px(210)
        bw = (c.width() - px(36) - 2 * px(8)) / 3
        for i, (key, name) in enumerate(zip(STEPS, BADGES)):
            r = QRectF(bx + (i % 3) * (bw + px(8)), by + (i // 3) * px(34), bw, px(26))
            got = key in self.completed
            p.setPen(QPen(QColor(253, 230, 138, 160 if got else 60), 1))
            p.setBrush(QColor(253, 230, 138, 40 if got else 10))
            p.drawRoundedRect(r, r.height() / 2, r.height() / 2)
            p.setFont(_font(base, 8.5, True))
            p.setPen(_C_GOLD if got else QColor(253, 230, 138, 110))
            p.drawText(r, Qt.AlignCenter, ('✓ ' if got else '') + name)

    # ── 입력 ──────────────────────────────────────────────────
    def mousePressEvent(self, e):          # 어두운 곳을 눌러도 아무 일 없음
        e.accept()

    def eventFilter(self, obj, event):
        t = event.type()
        if obj is self.win and t == QEvent.Resize:
            QTimer.singleShot(0, self._tick)
        elif t == QEvent.Show and isinstance(obj, QMenu) and not self.clear:
            if self._key() in ('row', 'menu'):
                self._menu_seen = True
        return False
