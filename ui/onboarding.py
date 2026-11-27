"""첫 사용 안내 — 게임 튜토리얼처럼 하나씩 직접 해 보기 (유주가 안내).

- 처음 켰을 때 한 번만 (db 설정 'onboarding_done'). 이미 업무 · 지난 기록이 있는 PC는 건너뜀.
  다시 보기: 「…」 → 도움말 → 사용법 다시 보기.
- 6단계: 업무 추가 · 업무 다루기 · 파일 바로가기 · 아래 줄 버튼 · AI 연결 · 메뉴 → 클리어 화면.
  단계마다 할 일(퀘스트)을 직접 하면 ✓, 다 하면 「✓ 완료!」 뒤 다음 단계로. [다음] 단추는 없다.
- 창을 어둡게 덮고 할 곳만 뚫어 빛나는 테두리로 (뚫린 곳은 그대로 누르고 · 쓰고 · 파일을 놓을 수 있음).
- 「이 단계 건너뛰기」 · 「튜토리얼 끝내기(Esc)」. 연습으로 만든 업무는 그대로 둔다.
"""
import math
import random

from PyQt5.QtCore import QEvent, QPoint, QPointF, QRect, QRectF, QSize, Qt, QTimer
from PyQt5.QtGui import (QBrush, QColor, QFont, QKeySequence, QLinearGradient, QPainter, QPainterPath,
                         QPen, QPixmap, QRadialGradient, QRegion)
from PyQt5.QtWidgets import QApplication, QMenu, QPushButton, QShortcut, QWidget

DONE_KEY = 'onboarding_done'
STEPS = ('add', 'row', 'file', 'bottom', 'keys', 'ai', 'menu')
STEP_TITLES = {'add': '업무 추가하기', 'row': '업무 다루기', 'file': '파일 바로가기',
               'bottom': '아래 줄 버튼', 'keys': '단축키', 'ai': 'AI 연결', 'menu': '메뉴'}
BADGES = ('업무 추가', '업무 다루기', '파일 바로가기', '아래 줄 버튼', '단축키', 'AI 연결', '메뉴')
YUJU = {'add': 6, 'row': 5, 'file': 1, 'bottom': 3, 'keys': 2, 'ai': 4, 'menu': 6, 'clear': 1}
KEY_TIPS = (('Alt+N', '업무 추가 칸으로'), ('Alt+1~4', '아래 줄 버튼'),
            ('F2', '업무 이름 고치기'), ('Ctrl+Shift+F', '항상 위에 고정'),
            ('Ctrl+Shift+R', '업무창 접기'), ('Ctrl+1·2·3', '공문 창 복사'))
_ACTIVE = []          # 지금 떠 있는 안내 (main.py · window.py가 notify로 알려 줌)


def notify(event: str):
    """앱 곳곳에서 튜토리얼에 알려 준다: 'capture'(Ctrl+Shift+X) · 'restore'(Ctrl+Shift+E) · 'palette'(Ctrl+K)."""
    for t in list(_ACTIVE):
        try:
            t.on_event(event)
        except RuntimeError:
            _ACTIVE.remove(t)
AI_BENEFITS = ('📝  제목만 쓰면 공문 초안을 써 줘요 (내 공문 모양대로)',
               '📸  캡처한 공문의 제목 · 번호를 정확하게 읽어요',
               '📊  예산 화면을 캡처하면 항목별로 정리해 줘요',
               '💬  「이번 주 마감 뭐야?」 물어보면 알려 줘요')
AI_NOTE = '구독(Claude · ChatGPT)이 있으면 1분이면 돼요. 없어도 앱은 그대로 쓸 수 있어요.'
AI_WHERE = ('💬  아래 줄 [AI] — 업무를 묻고 시키기 (「이번 주 마감 뭐야?」)',
            '📝  공문 작성 — 제목만 쓰면 초안을 써 줘요',
            '📸  공문 캡처 — 제목 · 번호를 자동으로 읽어요')
ADD_NOTES = ('⌨  Tab 다음 칸 · Shift+Tab 앞 칸 · ↑↓ 값 바꾸기',
             '📅  마감일 칸을 누르면 달력이 떠요 — 날짜를 눌러도 돼요',
             '↩  마감일이 없으면 바로 Enter, 잘못 넣은 날짜는 Esc로 지우기')
ROW_NOTES = ('⋮ 메뉴에서 완료 표시(C) · 삭제(D)를 해요',)
FILE_NOTES = ('🪟  업무창을 옆으로 옮겨 바탕화면이 보이게 해도 돼요',
              '📂  붙은 파일은 더블클릭하면 바로 열려요')
ADD_LINES = {
    'write': '할 일을 하나 써 볼까요?',
    'name': 'Tab을 누르면 마감일 칸으로 가요! (오늘 날짜가 들어가요)',
    'year': '여기는 연도예요. 보통 그대로 두고 Tab!',
    'month': '월이에요. 숫자를 치거나 ↑↓로 바꿔요. 다 했으면 Tab!',
    'day': '일이에요. 바꿨으면 Tab → 시간으로!',
    'hour': '시예요. 24시간으로 써요 (14 = 오후 2시). Tab!',
    'minute': '분이에요. Tab을 누르면 반복 칸으로!',
    'repeat': '반복! ↑↓로 매주 · 격주 · 매월 · 매년. 다 됐으면 Enter!',
    'calendar': '달력에서 날짜를 눌러도 돼요!',
    'enter': '다 됐으면 Enter로 업무를 넣어요!',
}
BURST_MS = 900
TYPE_PAUSE_MS = 700         # 할 일을 다 쓰고 타자를 멈춘 뒤 이만큼 지나야 ✓
INTRO_TEXT = '안녕하세요! 저는 서서니 노트 도우미 유주예요 🍊\n간단한 사용법을 같이 알아보아요~'

# 색 (남보라 유리 카드)
_C_BG1, _C_BG2 = QColor(30, 27, 75, 246), QColor(49, 46, 129, 246)
_C_EDGE = QColor(165, 180, 252, 130)
_C_BUBBLE_EDGE = QColor('#7C3AED')       # 유주 말풍선 테두리 (보라)
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
    win._tour.show_intro()
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


SCALE = 1.3                 # 튜토리얼 전체 크기 (글자 · 그림 · 단추 · 카드)


def _font(base: QFont, pt, bold=False):
    f = QFont(base)
    f.setPointSizeF(pt * SCALE)
    f.setBold(bold)
    return f


class _LinkButton(QPushButton):
    """튜토리얼 단추. QSS 대신 직접 그린다 — Pretendard 진짜 굵기 + 부드러운 가장자리.
    (QSS 글자는 글꼴이 정해지지 않아 Windows 기본 글꼴의 가짜 굵게로 그려지고,
     투명 바탕 창 위에서는 ClearType도 꺼져 흐리게 보였다.)"""

    def __init__(self, text, parent, primary=False):
        super().__init__(text, parent)
        from .common import pr_font
        self.primary = primary
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.NoFocus)
        self.setAttribute(Qt.WA_Hover)
        self.setFlat(True)
        f = pr_font(10)
        f.setFamilies([f.family(), 'Pretendard', 'Malgun Gothic'])
        if primary:
            f.setPointSizeF(10.5 * SCALE)
            f.setWeight(QFont.DemiBold)
        else:
            f.setPointSizeF(9 * SCALE)
            f.setWeight(QFont.Medium)
        f.setHintingPreference(QFont.PreferNoHinting)
        f.setStyleStrategy(QFont.PreferAntialias)
        self.setFont(f)

    def _pad(self):
        from .common import px
        return (px(round(20 * SCALE)), px(round(8 * SCALE))) if self.primary else (px(4), px(round(4 * SCALE)))

    def sizeHint(self):
        fm = self.fontMetrics()
        px_, py = self._pad()
        return QSize(fm.horizontalAdvance(self.text()) + 2 * px_ + 2, fm.height() + 2 * py)

    def minimumSizeHint(self):
        return self.sizeHint()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        hover = self.underMouse()
        r = QRectF(self.rect())
        if self.primary:
            g = QLinearGradient(r.topLeft(), r.topRight())
            g.setColorAt(0, QColor('#4F46E5' if hover else '#6366F1'))
            g.setColorAt(1, QColor('#9333EA' if hover else '#A855F7'))
            p.setPen(Qt.NoPen)
            p.setBrush(g)
            rad = r.height() / 2.0
            p.drawRoundedRect(r.adjusted(0.5, 0.5, -0.5, -0.5), rad, rad)
            p.setPen(QColor('#FFFFFF'))
        else:
            p.setPen(QColor('#E0E7FF') if hover else QColor(199, 210, 254, 185))
        f = QFont(self.font())
        if hover and not self.primary:
            f.setUnderline(True)
        p.setFont(f)
        p.drawText(r, Qt.AlignCenter, self.text())
        p.end()


class _DropZone(QWidget):
    """파일 단계: 업무 줄 위의 점선 놓는 칸 「📎 여기에 놓기」. 마우스 · 끌어 놓기는 아래 업무 줄로 그대로 통과."""

    def __init__(self, tour):
        super().__init__(tour.win)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.hover = False
        self.t = 0
        self.hide()

    def paintEvent(self, _):
        from .common import pr_font
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(self.rect()).adjusted(1.5, 1.5, -1.5, -1.5)
        pulse = 0.5 + 0.5 * math.sin(self.t / 380.0)
        fill = QColor(253, 230, 138, 150 if self.hover else int(70 + 50 * pulse))
        p.setBrush(fill)
        pen = QPen(QColor('#D97706') if self.hover else QColor(217, 119, 6, int(150 + 100 * pulse)), 2.2 if self.hover else 1.8)
        pen.setStyle(Qt.SolidLine if self.hover else Qt.DashLine)
        p.setPen(pen)
        p.drawRoundedRect(r, 9, 9)
        f = _font(pr_font(10), 9.5, True)
        p.setFont(f)
        txt = '📎 여기에 놓으세요!' if self.hover else '📎 여기에 놓기'
        fm = p.fontMetrics()
        w, h = fm.horizontalAdvance(txt) + 20, min(fm.height() + 8, int(r.height()) - 4)
        lab = QRectF(r.right() - w - 8, r.center().y() - h / 2, w, h)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor('#D97706'))
        p.drawRoundedRect(lab, h / 2, h / 2)
        p.setPen(QColor('#FFFFFF'))
        p.drawText(lab, Qt.AlignCenter, txt)
        p.end()


class _SidePanel(QWidget):
    """업무창 바깥 옆에 뜨는 퀘스트 카드 + 유주 (포커스를 가져가지 않아 업무창에서 계속 타자 · Tab 가능)."""

    def __init__(self, tour):
        super().__init__(tour.win, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowDoesNotAcceptFocus)
        self.tour = tour
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFocusPolicy(Qt.NoFocus)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        self.tour._paint_card(p)


class TourOverlay(QWidget):
    """업무창을 덮는 어두운 막 + 빛나는 구멍 + 퀘스트 카드 + 유주. 직접 해야 다음으로 간다."""

    PAD = 4

    def __init__(self, win):
        super().__init__(win)
        from .common import px
        self.px = lambda n: px(round(n * SCALE))
        self.win = win
        self.step = 0
        self.task_id = None
        self.done = []                  # 이번 단계 퀘스트별 완료 여부
        self.completed = set()          # 끝낸 단계 (클리어 화면 배지)
        self.clear = False
        self.intro = False              # 맨 앞 유주 인사 화면
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
        self._cal_used = False          # 1단계: 마감일 칸 달력으로 날짜를 골랐는지
        self._ai_opened = False         # AI 단계: 연결 창을 열어 봤는지
        self._drag_over = False         # 파일 단계: 파일을 끌고 놓는 칸 위에 올라왔는지
        self._confetti = []
        self._outside = False           # 카드를 업무창 바깥 옆 창에 그리는지
        self._last_text, self._typed_at = '', 0
        self.panel = _SidePanel(self)
        self.drop_zone = _DropZone(self)

        self.btn_skip = _LinkButton('이 단계 건너뛰기', self)
        self.btn_quit = _LinkButton('튜토리얼 끝내기 (Esc)', self)
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
        self._events = set()
        self._closed_editor = False
        _ACTIVE.append(self)
        self.setGeometry(win.rect())
        self.raise_()
        self.show()

    def on_event(self, event):
        if not self.clear:
            self._events.add(event)

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
        if self.clear or self.intro:
            return QRect()
        key, w = self._key(), self.win
        if key == 'add':
            r = self._rect(w.input_name)
            dr = getattr(w, 'date_row_widget', None)
            if dr is not None and dr.isVisible():
                r = r.united(self._rect(dr))
            return r
        if key == 'row' and len(self.done) > 2 and self.done[1] and not self.done[2]:
            row = self._row()                            # ⋮ 단추만 밝게
            if row is not None:
                return self._rect(row.btn_menu)
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
        if key == 'bottom':                              # 첫 번째(Alt+1) 단추만
            k, _label = self._bottom_first()
            return self._rect(w._bottom_btns.get(k, w.btn_ai_bottom))
        if key == 'ai' and self._ai_connected():
            return self._rect(w.btn_ai_bottom)
        if key == 'menu':
            return self._rect(w.title_bar.btn_menu)
        return QRect()

    # ── 퀘스트 ────────────────────────────────────────────────
    def quests(self):
        key = self._key()
        if key == 'add':
            return ['할 일 쓰기', 'Tab → 마감일 (연 · 월 · 일)', 'Tab → 시간 (시 · 분)', 'Tab → 반복', 'Enter 눌러 업무 넣기']
        if key == 'row':
            return ['☆ 눌러 중요 표시하기', '연필 아이콘 눌러 메모 열기', '⋮ 눌러 메뉴 열고 닫기']
        if key == 'file':
            return ['바탕화면의 파일을 노란 칸에 끌어다 놓기']
        if key == 'bottom':
            _k, label = self._bottom_first()
            return [f'Alt+1 눌러 {label} 창 열기', 'Alt+1 한 번 더 눌러 닫기']
        if key == 'keys':
            return ['Ctrl+Shift+X 눌러 캡처 화면 띄우기',
                    '다른 창을 누른 뒤 Ctrl+Shift+E로 불러오기', 'Ctrl+K 눌러 빠른 찾기 열기']
        if key == 'ai':
            return ['아래 줄 [AI] 눌러 대화창 열어 보기'] if self._ai_connected() else ['[AI 연결하기]로 연결 창 열어 보기']
        return ['「…」 눌러 메뉴 열어 보기']

    def _ai_connected(self):
        try:
            from ai_client import load_ai_mode
            return load_ai_mode() != 'none'
        except Exception:
            return False

    def _info(self):
        """카드 위쪽(퀘스트 앞) 설명 줄."""
        if self._key() == 'ai':
            return list(AI_WHERE) if self._ai_connected() else list(AI_BENEFITS)
        return []

    def _notes(self):
        """카드 아래쪽(퀘스트 뒤) 작은 도움말 줄."""
        key = self._key()
        if key == 'add':
            return list(ADD_NOTES)
        if key == 'row':
            return list(ROW_NOTES)
        if key == 'file':
            return list(FILE_NOTES)
        if key == 'ai' and not self._ai_connected():
            return [AI_NOTE]
        return []

    def _lines(self):
        key = self._key()
        if key == 'add':
            return [ADD_LINES['write'], ADD_LINES['name'], ADD_LINES['hour'], ADD_LINES['repeat'], ADD_LINES['enter']]
        if key == 'row':
            return ['별을 눌러 중요 표시를 해 봐요!', '이번엔 연필 아이콘으로 메모를 열어요.',
                    '줄 오른쪽 끝 ⋮ 단추를 눌러 봐요!']
        if key == 'file':
            return ['바탕화면의 아무 파일이나 노란 칸에 끌어다 놓아 보세요!']
        if key == 'bottom':
            _k, label = self._bottom_first()
            return [f'Alt+1을 눌러 {label} 창을 열어 봐요!', '한 번 더 누르면 닫혀요!']
        if key == 'keys':
            return ['Ctrl+Shift+X를 누르면 바로 캡처해요! (Esc로 닫아도 돼요)',
                    '다른 창을 누른 뒤 Ctrl+Shift+E! 어디서든 제가 바로 나와요.',
                    'Ctrl+K를 누르면 기능 이름으로 바로 찾아 줘요.']
        if key == 'ai':
            if self._ai_connected():
                return ['벌써 AI가 연결돼 있네요! 어디서 쓰는지 볼까요?']
            return ['AI를 연결하면 이런 걸 해 줘요!']
        return ['마지막! 「…」을 눌러 메뉴를 열어 봐요.']

    def _set_line(self, text):
        if text != self._line:
            self._line, self._typed = text, 0

    def show_intro(self):
        """유주 인사 → [시작하기]를 누르면 1단계로."""
        self.intro = True
        self.step = 0
        self.done = [False] * len(self.quests())
        self._set_yuju(YUJU['add'])
        self._line, self._typed = INTRO_TEXT, len(INTRO_TEXT)      # 한 글자씩이 아니라 한 번에
        self._layout_buttons()
        self._hole = QRect(-1, -1, 0, 0)
        self._tick()

    def begin(self):
        if self.intro:
            self.intro = False
            self._yuju_n = None
            self.show_step(0)

    def show_step(self, i):
        self.step = i
        self.intro = False
        self.clear = False
        self._burst = None
        self.done = [False] * len(self.quests())
        key = self._key()
        self._scrolled = None
        self._bottom_seen_open = False
        self._menu_seen = False
        self._cal_used = False
        self._ai_opened = False
        self._drag_over = False
        self._events = set()
        self._panel_on_top(key == 'bottom')          # Alt+1로 연 창에 카드가 가려지지 않게
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
            if self.done[1]:                             # 마감일까지 가 봤으면 나머지도 ✓
                self.done = [True] * len(self.done)
            else:
                self._set_line('좋아요! 이번엔 Tab도 눌러서 하나 더 넣어 볼까요?')

    def _check(self):
        """지금 단계의 퀘스트를 살펴 ✓를 켠다 (순서대로)."""
        if self.clear or self.intro or self._burst is not None:
            return
        key, w, d = self._key(), self.win, self.done
        if key == 'add':
            text = w.input_name.text().strip()
            if text != self._last_text:                     # 칠 때마다 0.7초를 처음부터
                self._last_text, self._typed_at = text, self._t
            if text and self._t - self._typed_at >= TYPE_PAUSE_MS:
                d[0] = True
            sec = self._section()
            cal = getattr(w, '_cal_popup', None)
            cal_open = cal is not None and cal.isVisible()
            if cal_open:
                self._cal_used = True
            if d[0]:
                if sec in ('day', 'hour', 'minute', 'repeat') or (self._cal_used and not cal_open and w._date_touched):
                    d[1] = True
                if sec in ('minute', 'repeat'):
                    d[2] = True
                if sec == 'repeat':
                    d[3] = True
            if all(d):                                   # Enter로 넣었으면 완료
                self._complete()
                return
            # 커서가 있는 칸에 맞춰 한 칸씩 설명
            if cal_open:
                self._set_line(ADD_LINES['calendar'])
            elif sec in ADD_LINES and sec != 'name':
                self._set_line(ADD_LINES[sec])
            elif not d[0]:
                self._set_line(ADD_LINES['write'])
            elif not d[1]:
                self._set_line(ADD_LINES['name'])
            else:
                self._set_line(ADD_LINES['enter'])
            return
        elif key == 'row':
            row = self._row()
            if row is not None:
                if self._task_priority():
                    d[0] = True
                if d[0] and getattr(row, '_expanded', False):
                    d[1] = True
                try:                                     # 연필 · ⋮ 단추는 원래 마우스를 올려야 보임 → 이 단계에선 늘 보이게
                    row.btn_note.set_row_hover(True)
                    if d[1]:
                        row.btn_menu.set_row_active(True)
                except (AttributeError, RuntimeError):
                    pass
                menu_open = QApplication.activePopupWidget() is not None
                if d[1] and self._menu_seen and menu_open:
                    self._set_line('여기서 완료 표시(C) · 삭제(D)를 해요. 메뉴를 닫으면 다음!')
                    return
                if d[1] and self._menu_seen and not menu_open:   # 메뉴를 닫은 뒤에 ✓
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
        elif key == 'keys':
            ev = self._events
            if 'capture' in ev:
                d[0] = True
                self._close_capture_editor()
            if d[0] and 'restore' in ev:
                d[1] = True
            if d[1] and 'palette' in ev:
                d[2] = True
        elif key == 'ai':
            if self._ai_connected():
                if self._ai_opened or self._bottom_open('ai'):    # 연결된 PC: [AI] 대화창을 열어 봐야 ✓
                    d[0] = True
            elif self._ai_opened:                                # 연결 안 된 PC: 연결 창을 열어 보기만 해도 ✓
                d[0] = True
        elif key == 'menu':
            if self._menu_seen:
                d[0] = True
        if not all(d):
            self._update_line()
        else:
            self._complete()

    def _focus_field(self):
        """업무 추가 줄에서 지금 커서가 있는 칸: 'date' · 'time' · 'repeat' · ''."""
        fw = QApplication.focusWidget()
        if fw is None:
            return ''
        for name, attr in (('date', 'input_date'), ('time', 'input_time'), ('repeat', 'input_recurrence')):
            box = getattr(self.win, attr, None)
            if box is not None and (fw is box or box.isAncestorOf(fw)):
                return name
        return ''

    def _section(self):
        """업무 추가 줄에서 커서 자리: 'name' · 'year' · 'month' · 'day' · 'hour' · 'minute' · 'repeat' · ''."""
        from PyQt5.QtWidgets import QDateTimeEdit
        fw = QApplication.focusWidget()
        w = self.win
        if fw is None:
            return ''
        if fw is w.input_name:
            return 'name'
        where = self._focus_field()
        if where == 'repeat':
            return 'repeat'
        if where == 'date':
            return {QDateTimeEdit.YearSection: 'year', QDateTimeEdit.MonthSection: 'month',
                    QDateTimeEdit.DaySection: 'day'}.get(w.input_date.currentSection(), 'year')
        if where == 'time':
            return 'minute' if w.input_time.currentSection() == QDateTimeEdit.MinuteSection else 'hour'
        return ''

    def _panel_on_top(self, on):
        """Alt+1 단계: 옆 카드를 공문 창 등 다른 창 위로 계속 올린다 (_tick에서 raise_).
        (창 속성을 '항상 위'로 바꾸면 Qt가 창을 새로 만들다 드물게 앱이 꺼져서, 속성은 건드리지 않는다)"""
        self._keep_panel_top = bool(on)

    def _close_capture_editor(self):
        """Ctrl+Shift+X 연습으로 열린 공문 작성 창: 캡처가 끝나고 쓴 내용이 없으면 닫고 업무창으로 돌아온다."""
        if self._closed_editor:
            return
        try:
            from document_editor import DocumentEditorWindow
            ed = DocumentEditorWindow._instance
            if ed is None or getattr(ed, '_capturing', False) or not ed.isVisible():
                return
            self._closed_editor = True
            if not self.win._doc_has_text(ed):
                ed.close()
            QTimer.singleShot(0, self.win._take_focus_back)
        except Exception:
            pass

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
        if self._key() == 'file':
            self._set_line('붙었어요! 더블클릭하면 바로 열려요.')
        else:
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
        if self.intro:
            self.begin()
            return
        if self.clear:
            self.finish()
            return
        if self._key() == 'ai':
            self._ai_opened = True
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
        if self in _ACTIVE:
            _ACTIVE.remove(self)
        for obj in (self.panel, self.drop_zone):
            try:
                obj.hide()
                obj.deleteLater()
            except RuntimeError:
                pass
        self._restore_row_controls()
        self.hide()
        self.deleteLater()

    def _restore_row_controls(self):
        """튜토리얼에서 늘 보이게 했던 연필 · ⋮ 단추를 원래대로 (마우스를 올렸을 때만)."""
        try:
            from .task_row import TaskRow
            for r in self.win.findChildren(TaskRow):
                r._sync_row_controls()
        except (RuntimeError, AttributeError):
            pass

    # ── 배치 ──────────────────────────────────────────────────
    def _big(self):
        return self.clear or self.intro

    def _card_height(self):
        px = self.px
        if self.intro:
            return px(268)
        if self.clear:
            return px(372)
        h = px(78) + len(self.done) * px(28) + px(34)
        info, notes = self._info(), self._notes()
        if info:
            h += len(info) * px(22) + px(8)
        if notes:
            h += len(notes) * px(20) + px(10)
        if self._key() == 'keys':
            h += px(28) + ((len(KEY_TIPS) + 1) // 2) * px(20)
        return h

    def _layout_buttons(self):
        if self.intro:
            self.btn_skip.hide()
            self.btn_quit.hide()
            self.btn_later.hide()
            self.btn_action.setText('시작하기')
            self.btn_action.show()
            self.btn_action.adjustSize()
            return
        ai = (not self.clear) and self._key() == 'ai'
        ai_needed = ai and not self._ai_connected()
        later = ai_needed or ((not self.clear) and self._key() == 'file')
        self.btn_skip.setVisible(not self.clear and not later)
        self.btn_quit.setVisible(not self.clear)
        self.btn_action.setVisible(self.clear or ai_needed)
        self.btn_later.setVisible(later)
        self.btn_action.setText('시작하기' if self.clear else 'AI 연결하기')
        self.btn_later.setText('다음에 할게요')
        for b in (self.btn_skip, self.btn_quit, self.btn_action, self.btn_later):
            b.adjustSize()

    def _host(self):
        return self.panel if self._outside else self

    def _side_spot(self, pw, ph):
        """업무창 바깥 오른쪽(없으면 왼쪽)에 옆 창을 둘 자리. 둘 다 모자라면 None."""
        w = self.win
        fg = w.frameGeometry()
        scr = None
        try:
            scr = w.screen().availableGeometry()
        except Exception:
            sc = QApplication.screenAt(fg.center()) or QApplication.primaryScreen()
            scr = sc.availableGeometry() if sc else None
        if scr is None:
            return None
        gap = self.px(10)
        if fg.right() + gap + pw <= scr.right():
            x = fg.right() + gap
        elif fg.left() - gap - pw >= scr.left():
            x = fg.left() - gap - pw
        else:
            return None
        y = fg.top() + (fg.height() - ph) // 2
        y = max(scr.top(), min(y, scr.bottom() - ph))
        return QPoint(x, y)

    def _place(self):
        px = self.px
        W, H = self.width(), self.height()
        ch = self._card_height()
        # ① 바깥 옆 창에 둘 수 있으면 거기에
        cw_out = px(440) if self.intro else px(400)
        top = px(8) if self._big() else px(100)         # 위쪽은 유주 · 말풍선 자리
        pw, ph = cw_out + px(8), top + ch + px(8)
        spot = self._side_spot(pw, ph) if self.win.isVisible() else None
        outside = spot is not None
        if outside != self._outside:
            self._outside = outside
            for b in (self.btn_skip, self.btn_quit, self.btn_action, self.btn_later):
                vis = not b.isHidden()
                b.setParent(self._host())
                b.setVisible(vis)
        if outside:
            self.panel.setGeometry(spot.x(), spot.y(), pw, ph)
            if not self.panel.isVisible():
                self.panel.show()
            self._card = QRect(px(4), top, cw_out, ch)
        else:
            self.panel.hide()
            cw = min(W - px(24), px(440))
            if self._big():
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
                if y < below_top + yuju_h:                   # 아래 자리가 모자라면 구멍 위쪽에
                    y = max(px(8) + yuju_h, min(y, H - ch - px(8)))
                self._card = QRect((W - cw) // 2, y, cw, ch)
        c = self._card
        m = px(16)
        if self._big():
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
        size = self.px(130 if self._big() else 88)
        if (n, size) != self._yuju_n:
            self._yuju_n = (n, size)
            dpr = self.devicePixelRatioF() or 1.0
            self._yuju_pm = _yuju_pixmap(n, size, dpr)

    # ── 시계 ──────────────────────────────────────────────────
    def _tick(self):
        if not self.isVisible():
            try:
                self.panel.hide()                    # 업무창이 숨으면 (캡처 중 등) 옆 창도
            except RuntimeError:
                pass
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
        self._update_drop_zone()
        if getattr(self, '_keep_panel_top', False) and self._outside and self.panel.isVisible():
            self.panel.raise_()                      # 포커스는 가져가지 않음 (WA_ShowWithoutActivating)
        for b in (self.btn_skip, self.btn_quit, self.btn_action, self.btn_later):
            b.raise_()
        self.raise_()
        self.update()
        if self._outside:
            self.panel.update()

    def _update_drop_zone(self):
        """파일 단계: 업무 줄 위에 「📎 여기에 놓기」 점선 칸 (마우스 · 끌어 놓기는 아래 업무 줄로 통과)."""
        dz = self.drop_zone
        show = (not self.clear and not self.intro and self._burst is None
                and self._key() == 'file' and not all(self.done))
        row = self._row() if show else None
        r = self._rect(row) if row is not None else QRect()
        if r.isEmpty():
            if dz.isVisible():
                dz.hide()
            return
        r = r.adjusted(2, 2, -2, -2)
        if dz.geometry() != r:
            dz.setGeometry(r)
        dz.hover = self._drag_over
        dz.t = self._t
        if not dz.isVisible():
            dz.show()
        dz.raise_()
        dz.update()

    # ── 그리기 ────────────────────────────────────────────────
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
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
            if self._key() == 'bottom' and not self.clear and not self.intro:
                self._paint_key_badge(p, 'Alt+1')
        if not self._outside:
            self._paint_card(p)

    def _paint_key_badge(self, p, text):
        """구멍(강조한 단추) 바로 위에 「Alt+1」 같은 키 표시."""
        from .common import pr_font
        f = _font(pr_font(10), 9.5, True)
        p.setFont(f)
        fm = p.fontMetrics()
        w, h = fm.horizontalAdvance(text) + self.px(18), fm.height() + self.px(8)
        hole = QRectF(self._hole)
        r = QRectF(hole.center().x() - w / 2, hole.top() - h - self.px(8), w, h)
        bob = math.sin(self._t / 300.0) * self.px(2)
        r.translate(0, bob)
        p.setPen(Qt.NoPen)
        p.setBrush(_C_GOLD)
        p.drawRoundedRect(r, h / 2, h / 2)
        tri = QPainterPath()
        tri.moveTo(r.center().x() - self.px(6), r.bottom() - 1)
        tri.lineTo(r.center().x() + self.px(6), r.bottom() - 1)
        tri.lineTo(r.center().x(), r.bottom() + self.px(6))
        p.fillPath(tri, _C_GOLD)
        p.setPen(QColor('#1E1B4B'))
        p.drawText(r, Qt.AlignCenter, text)

    def _paint_card(self, p):
        """퀘스트 카드 · 유주 · 완료 연출 · 클리어 화면 (창 안이든 바깥 옆 창이든 self._card 자리에)."""
        px = self.px
        c = QRectF(self._card)
        g = QLinearGradient(c.topLeft(), c.bottomRight())
        g.setColorAt(0, _C_BG1)
        g.setColorAt(1, _C_BG2)
        p.setPen(QPen(_C_EDGE, 1.2))
        p.setBrush(QBrush(g))
        p.drawRoundedRect(c, 14, 14)
        from .common import pr_font
        base = pr_font(10)
        base.setStyleStrategy(QFont.PreferAntialias)
        p.setRenderHint(QPainter.TextAntialiasing)
        if self.intro:
            self._paint_intro(p, c, base)
            return
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
        info = self._info()
        if info:
            p.setFont(_font(base, 9.5))
            p.setPen(QColor(224, 231, 255, 230))
            for i, line in enumerate(info):
                p.drawText(QRectF(c.left() + m, y0 + i * px(22), c.width() - 2 * m, px(20)),
                           Qt.AlignLeft | Qt.AlignVCenter, line)
            y0 += len(info) * px(22) + px(8)
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
        notes = self._notes()
        if notes:
            ny = y0 + len(self.done) * px(28) + px(2)
            p.setPen(QPen(QColor(165, 180, 252, 60), 1))
            p.drawLine(QPointF(c.left() + m, ny), QPointF(c.right() - m, ny))
            p.setFont(_font(base, 8.5))
            p.setPen(QColor(199, 210, 254, 215))
            for i, line in enumerate(notes):
                p.drawText(QRectF(c.left() + m, ny + px(6) + i * px(20), c.width() - 2 * m, px(18)),
                           Qt.AlignLeft | Qt.AlignVCenter, line)
        if self._key() == 'keys':
            ty = y0 + len(self.done) * px(28) + px(4)
            p.setPen(QPen(QColor(165, 180, 252, 60), 1))
            p.drawLine(QPointF(c.left() + m, ty), QPointF(c.right() - m, ty))
            p.setFont(_font(base, 8.5, True))
            p.setPen(_C_LAV)
            p.drawText(QRectF(c.left() + m, ty + px(4), c.width(), px(18)), Qt.AlignLeft | Qt.AlignVCenter,
                       '이것도 알아 두면 편해요')
            colw = (c.width() - 2 * m) / 2
            for i, (k, desc) in enumerate(KEY_TIPS):
                r = QRectF(c.left() + m + (i % 2) * colw, ty + px(24) + (i // 2) * px(20), colw, px(18))
                p.setFont(_font(base, 8.5, True))
                p.setPen(_C_GOLD)
                kw = p.fontMetrics().horizontalAdvance(k) + px(6)
                p.drawText(r, Qt.AlignLeft | Qt.AlignVCenter, k)
                p.setFont(_font(base, 8.5))
                p.setPen(QColor(224, 231, 255, 200))
                p.drawText(r.adjusted(kw, 0, 0, 0), Qt.AlignLeft | Qt.AlignVCenter, desc)
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
            # 흰 말풍선 + 보라 테두리 (흰 화면 위에서도 또렷하게) · 꼬리까지 한 덩어리로
            body = QPainterPath()
            body.addRoundedRect(bub, 12, 12)
            tail = QPainterPath()
            tail.moveTo(bub.left() + 1, bub.bottom() - px(20))
            tail.lineTo(bub.left() - px(10), bub.bottom() - px(10))
            tail.lineTo(bub.left() + 1, bub.bottom() - px(7))
            tail.closeSubpath()
            shape = body.united(tail)
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(30, 27, 75, 40))                # 옅은 그림자
            p.drawPath(shape.translated(0, px(2)))
            p.setPen(QPen(_C_BUBBLE_EDGE, 2))
            p.setBrush(QColor(255, 255, 255, 252))
            p.drawPath(shape)
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

    def _paint_intro(self, p, c, base):
        px = self.px
        pm = self._yuju_pm
        if not pm.isNull():
            w = pm.width() / (pm.devicePixelRatio() or 1)
            pop = min(1.0, self._t / 300.0)
            bob = math.sin(self._t / 300.0) * px(3) + (1 - pop) * px(18)
            p.drawPixmap(QPointF(c.center().x() - w / 2, c.top() + px(8) + bob), pm)
        lines = (self._line or '')[:self._typed].split('\n')
        y = c.top() + px(146)
        side = px(24)                                     # 양옆 여백 — 글이 길면 줄을 바꿔 잘리지 않게
        for i, line in enumerate(lines):
            f = _font(base, 12.5 if i == 0 else 11, False)
            f.setWeight(QFont.DemiBold if i == 0 else QFont.Normal)
            p.setFont(f)
            p.setPen(_C_GOLD if i == 0 else QColor('#E0E7FF'))
            box = QRectF(c.left() + side, y, c.width() - 2 * side, px(80))
            r = p.boundingRect(box, Qt.AlignHCenter | Qt.AlignTop | Qt.TextWordWrap, line)
            p.drawText(box, Qt.AlignHCenter | Qt.AlignTop | Qt.TextWordWrap, line)
            y += r.height() + px(6)

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
        p.setFont(_font(base, 8.5))
        p.setPen(QColor(199, 210, 254, 170))
        p.drawText(QRectF(c.left(), by + 3 * px(34) - px(2), c.width(), px(18)), Qt.AlignCenter,
                   '단축키는 빠른 찾기(Ctrl+K)에서도 볼 수 있어요')

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
        elif t in (QEvent.DragEnter, QEvent.DragMove, QEvent.DragLeave, QEvent.Drop) and not self.clear:
            if self._key() == 'file':
                try:
                    row = self._row()
                    inside = row is not None and isinstance(obj, QWidget) and (obj is row or row.isAncestorOf(obj))
                except RuntimeError:
                    inside = False
                if inside:
                    self._drag_over = t in (QEvent.DragEnter, QEvent.DragMove)
        return False
