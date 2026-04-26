import os
import ctypes
from PyQt5.QtSvg import QSvgRenderer
from datetime import date, datetime, timedelta
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QDateTimeEdit, QMenu, QAction, QWidgetAction,
    QMessageBox, QDialog, QGridLayout, QCalendarWidget, QToolButton,
    QPlainTextEdit, QSizePolicy, QGraphicsColorizeEffect, QTimeEdit,
    QListWidget, QAbstractItemView, QComboBox, QListWidgetItem
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QDateTime, QPoint, QPointF, QSize, QSettings, pyqtSignal, QPropertyAnimation, QEasingCurve, QRectF
from PyQt5.QtGui import QFont, QFontMetrics, QColor, QPainter, QTextCharFormat, QPalette, QTextOption, QTextLayout, QIcon, QPixmap, QFontDatabase, QPen
from db import (update_window, delete_window, get_tasks, add_task, delete_task, update_task,
                add_task_history, get_task_history, delete_task_history,
                set_task_priority, set_task_recurrence, get_task_notes, set_task_notes,
                search_tasks_all,
                get_documents, add_document, update_document, delete_document,
                get_official_documents, delete_official_document,
                get_window_memo, set_window_memo, get_all_windows,
                set_window_merge_group, clear_window_merge_group, get_next_merge_group_id)
from autostart import is_enabled as autostart_is_enabled, set_enabled as autostart_set
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, OcrWorker, _normalize_doc_number
import updater

TITLE_BAR_HEIGHT = 33
TITLE_COLOR      = '#f7c948'
SNAP_THRESHOLD   = 12  # 완전히 붙는 거리(px)
SNAP_ZONE        = 30  # 당기기 시작하는 거리(px)
ALIGN_THRESHOLD  = 12  # 정렬 스냅 확정 거리(px)
ALIGN_ZONE       = 30  # 정렬 스냅 인력 시작 거리(px)
SNAP_RESIST      = 18  # 정렬 스냅 탈출 거리(px)
BREAKAWAY_DIST   = 25  # 인접 스냅 탈출 거리(px) - 유클리드 거리 기준

PALETTE = [
    '#FDD663', '#FEFFA7', '#CCFF90', '#A8F0E8',
    '#AECBFA', '#D7AEFB', '#FDCFE8', '#E6C9A8', '#AAAAAA', '#DDDDDD', '#FFFFFF',
]

# ── Fluent System Icons 폰트 로드 ────────────────────────────────
_FI_FONT_ID        = -1
_FI_FILLED_FONT_ID = -1
_MAT_FONT_ID       = -1
_PR_FONT_ID        = -1

def _load_fluent_icons():
    global _FI_FONT_ID, _FI_FILLED_FONT_ID, _MAT_FONT_ID
    base = os.path.dirname(os.path.abspath(__file__))
    if _FI_FONT_ID == -1:
        _FI_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'FluentSystemIcons-Regular.ttf'))
    if _FI_FILLED_FONT_ID == -1:
        _FI_FILLED_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'FluentSystemIcons-Filled.ttf'))
    if _MAT_FONT_ID == -1:
        _MAT_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'MaterialIcons-Regular.ttf'))

def mi_font(size=14):
    """Fluent System Icons Regular QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_FI_FONT_ID)
    return QFont(families[0] if families else 'FluentSystemIcons-Regular', size)

def mi_font_filled(size=14):
    """Fluent System Icons Filled QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_FI_FILLED_FONT_ID)
    return QFont(families[0] if families else 'FluentSystemIcons-Filled', size)

def mat_font(size=14):
    """Material Icons QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_MAT_FONT_ID)
    return QFont(families[0] if families else 'Material Icons', size)

def _load_pretendard():
    global _PR_FONT_ID
    if _PR_FONT_ID == -1:
        base = os.path.dirname(os.path.abspath(__file__))
        _PR_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'Pretendard-Regular.ttf'))

def pr_font(size=11):
    """Pretendard Regular QFont 반환."""
    _load_pretendard()
    families = QFontDatabase.applicationFontFamilies(_PR_FONT_ID)
    return QFont(families[0] if families else 'Pretendard', size)

def mi_icon(codepoint, size=18, color='#555555'):
    """Fluent Icons 코드포인트를 QIcon으로 변환."""
    _load_fluent_icons()
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setFont(mi_font(int(size * 0.82)))
    painter.setPen(QColor(color))
    painter.drawText(pix.rect(), Qt.AlignCenter, codepoint)
    painter.end()
    return QIcon(pix)


# Fluent System Icons 코드포인트 (20px Regular 기준)
class MI:
    STAR          = '\uebaa'   # star_emphasis_20_regular (삐침 별) — Regular 폰트용
    STAR_BORDER   = '\uf70f'   # star_20_regular (아웃라인 별) — Regular 폰트용
    CHECK         = '\uf294'   # checkmark

# Material Icons 코드포인트
class MAT:
    REPEAT        = '\ue040'   # repeat


class AddDateButton(QLabel):
    """deadline이 없는 태스크에서 날짜 자리에 표시되는 클릭 가능한 레이블."""
    clicked = pyqtSignal()

    _STYLE_NORMAL = (
        'color: #bbb; background: transparent;'
        'border: 1px dashed #ccc; border-radius: 3px;'
        'padding: 1px 5px;'
    )
    _STYLE_HOVER = (
        'color: #888; background: rgba(212,184,0,0.1);'
        'border: 1px dashed #d4b800; border-radius: 3px;'
        'padding: 1px 5px;'
    )

    def __init__(self, *args, **kwargs):
        super().__init__('＋날짜', *args, **kwargs)
        self.setCursor(Qt.PointingHandCursor)
        self.setFont(pr_font(10))
        self.setStyleSheet(self._STYLE_NORMAL)

    def enterEvent(self, event):
        self.setStyleSheet(self._STYLE_HOVER)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setStyleSheet(self._STYLE_NORMAL)
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class ClickableLabel(QLabel):
    def __init__(self, callback, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._callback = callback

    def mousePressEvent(self, event):
        self._callback()
        super().mousePressEvent(event)


class DdayLabel(QLabel):
    """D‑day 표시 라벨. 텍스트를 2px 위로 당겨서 다른 행 요소들과 시각 중심을 맞춘다."""
    SHIFT = 2

    def __init__(self, dday_text, date_text, *args, **kwargs):
        super().__init__(dday_text, *args, **kwargs)
        self._dday_text = dday_text
        self._date_text = date_text

    def paintEvent(self, event):
        p = QPainter(self)
        p.setFont(self.font())
        p.setPen(self.palette().color(self.foregroundRole()))
        rect = self.rect()
        rect.translate(0, -self.SHIFT)
        p.drawText(rect, self.alignment(), self.text())
        p.end()

    def enterEvent(self, event):
        self.setText(self._date_text)
        f = self.font()
        f.setPointSizeF(f.pointSizeF() - 0.3)
        self.setFont(f)
        self._original_style = self.styleSheet()
        if '#e74c3c' in self._original_style:
            self.setStyleSheet(self._original_style.replace('#e74c3c', '#e96060'))
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setText(self._dday_text)
        f = self.font()
        f.setPointSizeF(f.pointSizeF() + 0.3)
        self.setFont(f)
        if hasattr(self, '_original_style'):
            self.setStyleSheet(self._original_style)
        super().leaveEvent(event)



def _svg_icon(svg_path, color, size) -> 'QIcon':
    with open(svg_path, 'r', encoding='utf-8') as f:
        data = f.read()
    data = data.replace('#000000', color)
    renderer = QSvgRenderer(data.encode('utf-8'))
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    renderer.render(p)
    p.end()
    return QIcon(pm)


def calc_dday(deadline_str):
    try:
        if len(deadline_str) > 10:  # 'yyyy-MM-dd HH:MM'
            deadline_dt = datetime.strptime(deadline_str, '%Y-%m-%d %H:%M')
            diff_days = (deadline_dt.date() - date.today()).days
            diff_seconds = (deadline_dt - datetime.now()).total_seconds()
        else:  # 'yyyy-MM-dd'
            diff_days = (date.fromisoformat(deadline_str) - date.today()).days
            diff_seconds = None
    except ValueError:
        return '날짜오류'

    if diff_days > 0:
        return f'D-{diff_days}'
    elif diff_days < 0:
        return f'D+{abs(diff_days)}'
    else:
        return 'D-day'


def _next_recurrence_deadline(deadline_str, recurrence):
    import calendar as _cal
    try:
        base = date.fromisoformat(deadline_str[:10])
    except ValueError:
        return ''

    next_date = None
    if recurrence == 'weekly':
        next_date = base + timedelta(weeks=1)
    elif recurrence == 'biweekly':
        next_date = base + timedelta(weeks=2)
    elif recurrence == 'monthly':
        year = base.year + (base.month // 12)
        month = (base.month % 12) + 1
        day = min(base.day, _cal.monthrange(year, month)[1])
        next_date = date(year, month, day)
    elif recurrence == 'yearly':
        year = base.year + 1
        day = min(base.day, _cal.monthrange(year, base.month)[1])
        next_date = date(year, base.month, day)
    elif recurrence.startswith('custom:'):
        # custom:month|week|weekday (month 0=매월, 1~12 / week 1~4 / weekday 0=월~6=일)
        try:
            _, parts = recurrence.split(':')
            m_target, w_target, wd_target = map(int, parts.split('|'))
            
            # 기준일 다음날부터 탐색
            curr = base + timedelta(days=1)
            found = False
            # 최대 12년치 탐색 (특정 월 반복의 경우 고려)
            for _ in range(365 * 12):
                # 월 체크
                if m_target != 0 and curr.month != m_target:
                    # 월이 다르면 다음 달 1일로 점프 (최적화)
                    if curr.month == 12:
                        curr = date(curr.year + 1, 1, 1)
                    else:
                        curr = date(curr.year, curr.month + 1, 1)
                    continue
                
                # 요일 체크
                if curr.weekday() == wd_target:
                    # 몇째주인지 계산 (해당 월의 n번째 요일)
                    # 1일의 요일을 구해서 계산
                    first_day_of_month = date(curr.year, curr.month, 1)
                    first_wd = first_day_of_month.weekday()
                    # curr.day가 해당 월의 몇 번째 wd_target 인지 계산
                    # 예: 1일이 월(0)이고 8일도 월(0)이면, (8-1)//7 + 1 = 2번째
                    nth = (curr.day - 1) // 7 + 1
                    
                    if nth == w_target:
                        next_date = curr
                        found = True
                        break
                
                curr += timedelta(days=1)
            if not found:
                return ''
        except:
            return ''
    else:
        return ''
    
    if not next_date:
        return ''
        
    time_suffix = deadline_str[10:] if len(deadline_str) > 10 else ''
    return next_date.isoformat() + time_suffix


class CustomRecurrenceDialog(QDialog):
    """월/주/요일 사용자 설정 다이얼로그"""
    def __init__(self, parent=None, current=''):
        super().__init__(parent)
        self.setWindowTitle('사용자 설정 반복')
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setMinimumWidth(300)
        self.setStyleSheet("font-family: 'Malgun Gothic'; font-size: 10pt; background-color: #ffffff; color: #111111;")
        
        layout = QVBoxLayout(self)
        
        # 월 선택
        row_m = QHBoxLayout()
        row_m.addWidget(QLabel('반복 월:'))
        self.combo_m = QComboBox()
        self.combo_m.addItem('매월', 0)
        for i in range(1, 13):
            self.combo_m.addItem(f'{i}월', i)
        row_m.addWidget(self.combo_m, 1)
        layout.addLayout(row_m)
        
        # 주 선택
        row_w = QHBoxLayout()
        row_w.addWidget(QLabel('반복 주:'))
        self.combo_w = QComboBox()
        for i in range(1, 5):
            self.combo_w.addItem(f'{i}째주', i)
        row_w.addWidget(self.combo_w, 1)
        layout.addLayout(row_w)
        
        # 요일 선택
        row_wd = QHBoxLayout()
        row_wd.addWidget(QLabel('반복 요일:'))
        self.combo_wd = QComboBox()
        weekdays = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
        for i, name in enumerate(weekdays):
            self.combo_wd.addItem(name, i)
        row_wd.addWidget(self.combo_wd, 1)
        layout.addLayout(row_wd)
        
        # 현재 값 적용
        if current.startswith('custom:'):
            try:
                _, parts = current.split(':')
                m, w, wd = map(int, parts.split('|'))
                idx_m = self.combo_m.findData(m)
                if idx_m >= 0: self.combo_m.setCurrentIndex(idx_m)
                idx_w = self.combo_w.findData(w)
                if idx_w >= 0: self.combo_w.setCurrentIndex(idx_w)
                idx_wd = self.combo_wd.findData(wd)
                if idx_wd >= 0: self.combo_wd.setCurrentIndex(idx_wd)
            except: pass

        # 버튼
        btns = QHBoxLayout()
        btn_ok = QPushButton('확인')
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton('취소')
        btn_cancel.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)

    def get_value(self):
        m = self.combo_m.currentData()
        w = self.combo_w.currentData()
        wd = self.combo_wd.currentData()
        return f'custom:{m}|{w}|{wd}'


class EdgeHandle(QWidget):
    EDGE  = 6   # 가장자리 감지 두께(px)
    CORNER = 16  # 모서리 감지 크기(px)
    MIN_W = 320
    MIN_H = 100

    _CURSORS = {
        'left':         Qt.SizeHorCursor,
        'right':        Qt.SizeHorCursor,
        'bottom':       Qt.SizeVerCursor,
        'bottom-left':  Qt.SizeBDiagCursor,
        'bottom-right': Qt.SizeFDiagCursor,
        'top':          Qt.SizeVerCursor,
        'top-left':     Qt.SizeFDiagCursor,
        'top-right':    Qt.SizeBDiagCursor,
    }

    def __init__(self, parent, edge):
        super().__init__(parent)
        self.edge         = edge
        self._drag_start  = None
        self._start_geom  = None
        self.setCursor(self._CURSORS[edge])
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.setStyleSheet('background: transparent;')

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_start = e.globalPos()
            self._start_geom = self.window().geometry()
            g = self._start_geom
            self._right_anchor = g.x() + g.width()   # 왼쪽 드래그 시 고정되는 오른쪽 끝
            self._bottom_anchor = g.y() + g.height()  # 위쪽 드래그 시 고정되는 아래쪽 끝
            self.grabKeyboard()
            # 리사이즈 중 합체 파트너가 딸려오지 않도록 양쪽 동기화 일시 정지
            win = self.window()
            if getattr(win, '_link_group', None):
                for m in win._link_group:
                    m._is_syncing_move = True

    def _anchor_partner(self, win):
        """파트너가 리사이즈로 움직이는 쪽 모서리에 붙어 있을 때만 따라붙게 한다.
        고정된 모서리 쪽에 파트너가 있으면 아무것도 하지 않는다."""
        if not win._link_group:
            return
        # 리사이즈 방향에 인접한 그룹 멤버를 파트너로 선택
        partner = None
        if self.edge in ('right', 'top-right', 'bottom-right'):
            cands = [m for m in win._link_group if m is not win and m.x() > win.x()]
            if cands: partner = min(cands, key=lambda m: m.x())
        elif self.edge in ('left', 'top-left', 'bottom-left'):
            cands = [m for m in win._link_group if m is not win and m.x() < win.x()]
            if cands: partner = max(cands, key=lambda m: m.x())
        elif self.edge in ('bottom', 'bottom-left', 'bottom-right'):
            cands = [m for m in win._link_group if m is not win and m.y() > win.y()]
            if cands: partner = min(cands, key=lambda m: m.y())
        elif self.edge in ('top', 'top-left', 'top-right'):
            cands = [m for m in win._link_group if m is not win and m.y() < win.y()]
            if cands: partner = max(cands, key=lambda m: m.y())
        if partner is None:
            return
        partner_on_right = win.x() < partner.x()
        partner_on_bottom = win.y() < partner.y()
        nx, ny = partner.x(), partner.y()
        need_move = False
        if self.edge in ('right', 'top-right', 'bottom-right') and partner_on_right:
            nx = win.x() + win.width()
            need_move = True
        elif self.edge in ('left', 'top-left', 'bottom-left') and not partner_on_right:
            nx = win.x() - partner.width()
            need_move = True
        if self.edge in ('bottom', 'bottom-left', 'bottom-right') and partner_on_bottom:
            ny = win.y() + win.height()
            need_move = True
        elif self.edge in ('top', 'top-left', 'top-right') and not partner_on_bottom:
            ny = win.y() - partner.height()
            need_move = True
        if need_move:
            partner.move(nx, ny)

    def mouseMoveEvent(self, e):
        if not (e.buttons() == Qt.LeftButton and self._drag_start):
            return
        d = e.globalPos() - self._drag_start
        g = self._start_geom
        win = self.window()

        if self.edge == 'left':
            temp_x = g.x() + d.x()
            max_x = self._right_anchor - self.MIN_W
            new_x = min(temp_x, max_x)
            new_w = self._right_anchor - new_x
            win.setGeometry(new_x, g.y(), new_w, g.height())
            self._anchor_partner(win)

        elif self.edge == 'right':
            win.resize(max(self.MIN_W, g.width() + d.x()), g.height())
            self._anchor_partner(win)

        elif self.edge == 'bottom':
            win.resize(g.width(), max(self.MIN_H, g.height() + d.y()))
            self._anchor_partner(win)

        elif self.edge == 'bottom-right':
            win.resize(
                max(self.MIN_W, g.width() + d.x()),
                max(self.MIN_H, g.height() + d.y()),
            )
            self._anchor_partner(win)

        elif self.edge == 'bottom-left':
            temp_x = g.x() + d.x()
            max_x = self._right_anchor - self.MIN_W
            new_x = min(temp_x, max_x)
            new_w = self._right_anchor - new_x
            new_h = max(self.MIN_H, g.height() + d.y())
            win.setGeometry(new_x, g.y(), new_w, new_h)
            self._anchor_partner(win)

        elif self.edge == 'top':
            temp_y = g.y() + d.y()
            max_y = self._bottom_anchor - self.MIN_H
            new_y = min(temp_y, max_y)
            new_h = self._bottom_anchor - new_y
            win.setGeometry(g.x(), new_y, g.width(), new_h)
            self._anchor_partner(win)

        elif self.edge == 'top-left':
            temp_x = g.x() + d.x()
            max_x = self._right_anchor - self.MIN_W
            new_x = min(temp_x, max_x)
            new_w = self._right_anchor - new_x
            temp_y = g.y() + d.y()
            max_y = self._bottom_anchor - self.MIN_H
            new_y = min(temp_y, max_y)
            new_h = self._bottom_anchor - new_y
            win.setGeometry(new_x, new_y, new_w, new_h)
            self._anchor_partner(win)

        elif self.edge == 'top-right':
            new_w = max(self.MIN_W, g.width() + d.x())
            temp_y = g.y() + d.y()
            max_y = self._bottom_anchor - self.MIN_H
            new_y = min(temp_y, max_y)
            new_h = self._bottom_anchor - new_y
            win.setGeometry(g.x(), new_y, new_w, new_h)
            self._anchor_partner(win)

    def mouseReleaseEvent(self, e):
        self.releaseKeyboard()
        self._drag_start = None
        self._start_geom = None
        # 리사이즈 완료 후 offset 재계산 후 동기화 재개
        win = self.window()
        if getattr(win, '_link_group', None):
            for m in win._link_group:
                m._is_syncing_move = False
        win.save_state()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape and self._drag_start:
            self.window().setGeometry(self._start_geom)
            self.releaseKeyboard()
            self._drag_start = None
            self._start_geom = None


class MergeButton(QWidget):
    """두 창 경계선에 떠있는 합체/분리 토글 버튼."""
    SIZE = 56

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFixedSize(self.SIZE, self.SIZE)
        self._merged = False
        self._win_a  = None
        self._win_b  = None
        self._raise_timer = QTimer(self)
        self._raise_timer.setInterval(150)
        self._raise_timer.timeout.connect(self._do_raise)

    def set_windows(self, win_a, win_b):
        self._win_a = win_a
        self._win_b = win_b

    def set_merged(self, merged: bool):
        self._merged = merged
        self.update()

    def update_position(self, win_a, win_b):
        ga = win_a.geometry()
        gb = win_b.geometry()
        top_y = min(ga.top(), gb.top())
        if ga.right() <= gb.left() + SNAP_THRESHOLD:
            bx = ga.right()
        else:
            bx = gb.right()
        self.move(bx - self.SIZE // 2, top_y + TITLE_BAR_HEIGHT - self.SIZE // 2 + 12)

    def _do_raise(self):
        try:
            ctypes.windll.user32.SetWindowPos(
                int(self.winId()), 0, 0, 0, 0, 0,
                0x0010 | 0x0001 | 0x0002)
        except Exception:
            pass

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        s = self.SIZE
        bw, bh = 20, 16
        by = (s - bh) // 2
        f = QFont('Arial', 11, QFont.Bold)
        p.setFont(f)
        if self._merged:
            nx, sx = 8, 28
        else:
            nx, sx = 3, 33
        p.setPen(Qt.NoPen)
        p.setBrush(QColor('#e74c3c'))
        p.drawRoundedRect(nx, by, bw, bh, 4, 4)
        p.setBrush(QColor('#3498db'))
        p.drawRoundedRect(sx, by, bw, bh, 4, 4)
        if not self._merged:
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(nx + bw, by, sx - (nx + bw), bh)
        p.setPen(QColor(255, 255, 255))
        p.drawText(QRectF(nx, by, bw, bh), Qt.AlignCenter, 'N')
        p.drawText(QRectF(sx, by, bw, bh), Qt.AlignCenter, 'S')
        p.end()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._win_a and self._win_b:
            self._win_a.toggle_merge(self._win_b)

    def enterEvent(self, event):
        self.setCursor(Qt.PointingHandCursor)
        super().enterEvent(event)

    def showEvent(self, event):
        super().showEvent(event)

    def hideEvent(self, event):
        super().hideEvent(event)
        self._raise_timer.stop()


class TitleBar(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent    = parent
        self._drag_pos = None
        self.setFixedHeight(TITLE_BAR_HEIGHT)
        self.setStyleSheet(f"""
            QWidget {{
                background: {TITLE_COLOR};
                border-radius: 0;
            }}
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
                padding: 0;
            }}
            QPushButton:hover {{ background: rgba(0,0,0,0.12); border-radius: 3px; }}
            QPushButton:pressed {{ background: rgba(0,0,0,0.22); border-radius: 3px; }}
            QPushButton#btn_mode {{
                background: transparent;
                border: 1px solid transparent;
                border-radius: 6px;
                padding: 0px 6px;
                color: #2979ff;
            }}
            QPushButton#btn_mode:hover {{ background: rgba(0,0,0,0.12); border-color: rgba(0,0,0,0.12); }}
            QPushButton#btn_mode:pressed {{ background: rgba(0,0,0,0.20); border-color: rgba(0,0,0,0.20); }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 8, 0)
        layout.setSpacing(2)

        self.label = QLabel('서서니 노트')
        self.label.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        self.label.setStyleSheet('color: #5a4000; background: transparent;')

        btn_menu  = QPushButton('…')
        btn_menu.setToolTip('메뉴')
        btn_menu.setStyleSheet('font-size: 19px; letter-spacing: 0px; padding: 0px 6px 0px 6px;')
        self.btn_pin   = QPushButton('📌')
        self.btn_pin.setToolTip('항상 위 꺼짐')
        self.btn_pin.setFixedSize(26, 26)
        self.btn_pin.setFont(QFont('Segoe UI Emoji', 12))
        _gray = QGraphicsColorizeEffect()
        _gray.setColor(QColor('#888888'))
        _gray.setStrength(1.0)
        self.btn_pin.setGraphicsEffect(_gray)
        _x_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', '엑스아이콘.png')
        self.btn_close = QPushButton()
        self.btn_close.setIcon(QIcon(QPixmap(_x_icon_path).scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation)))
        self.btn_close.setIconSize(QSize(18, 18))
        self.btn_close.setFixedSize(26, 26)
        self.btn_close.setToolTip('닫기')
        self.btn_close.setStyleSheet("""
            QPushButton { background: transparent; border: none; }
            QPushButton:hover { background: rgba(0,0,0,0.12); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.22); border-radius: 3px; padding: 2px 0 0 2px; }
        """)

        _edit_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', '수정 아이콘.png')
        self.btn_mode = QPushButton()
        self.btn_mode.setObjectName('btn_mode')
        self.btn_mode.setIcon(QIcon(QPixmap(_edit_icon_path).scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation)))
        self.btn_mode.setIconSize(QSize(18, 18))
        self.btn_mode.setFixedSize(26, 26)
        self.btn_mode.setToolTip('메모 모드로 전환')
        _gray_mode = QGraphicsColorizeEffect()
        _gray_mode.setColor(QColor('#888888'))
        _gray_mode.setStrength(1.0)
        self.btn_mode.setGraphicsEffect(_gray_mode)
        self.btn_mode.clicked.connect(parent._toggle_memo_mode)

        layout.addWidget(self.label)
        layout.addStretch()
        layout.addWidget(self.btn_mode, 0, Qt.AlignVCenter)
        layout.addWidget(btn_menu)
        layout.addWidget(self.btn_pin)
        layout.addWidget(self.btn_close)

        # 가운데 오버레이 상태 메시지 (전체 너비, 마우스 이벤트 통과)
        self.status_label = QLabel('', self)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont('Malgun Gothic', 9))
        self.status_label.setStyleSheet('color: #5a4000; background: transparent;')
        self.status_label.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.status_label.setGeometry(0, 0, self.width(), TITLE_BAR_HEIGHT)

        btn_menu.clicked.connect(lambda: parent.show_menu(btn_menu))

        self.btn_pin.clicked.connect(parent.toggle_always_on_top)
        self.btn_close.clicked.connect(parent.close)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.status_label.setGeometry(0, 0, self.width(), self.height())

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPos() - self.parent.frameGeometry().topLeft()
            self.parent._snapped_pos = None
            self.parent._current_snap_edges = []
            self.parent._snap_is_adj = False

    def mouseMoveEvent(self, e):
        if not (e.buttons() == Qt.LeftButton and self._drag_pos):
            return
        raw_pos = e.globalPos() - self._drag_pos
        p = self.parent
        new_pos, snap_edges, adj_snapped, x_snap_partner = p._snap_full(raw_pos)

        if p._snapped_pos is not None:
            threshold = BREAKAWAY_DIST if p._snap_is_adj else SNAP_RESIST
            dx = raw_pos.x() - p._snapped_pos.x()
            dy = raw_pos.y() - p._snapped_pos.y()
            if (dx*dx + dy*dy) ** 0.5 < threshold:
                if adj_snapped:
                    p._snapped_pos = new_pos
                    p._current_snap_edges = snap_edges
                    p._snap_is_adj = True
                else:
                    new_pos = p._snapped_pos
                    snap_edges = p._current_snap_edges
                    adj_snapped = p._snap_is_adj
            else:
                p._snapped_pos = None
                p._current_snap_edges = []
                p._snap_is_adj = False

        if p._snapped_pos is None:
            if adj_snapped:
                p._snapped_pos = new_pos
                p._current_snap_edges = snap_edges
                p._snap_is_adj = True
            elif snap_edges:
                p._snapped_pos = new_pos
                p._current_snap_edges = snap_edges
                p._snap_is_adj = False
            else:
                p._snapped_pos = None
                p._current_snap_edges = []

        p.move(new_pos)

        # 합체 버튼 노출 / 갱신 (좌우 모서리 + 상단 또는 하단 라인 일치)
        y_aligned = (x_snap_partner is not None and (
            abs(p.y() - x_snap_partner.y()) < SNAP_THRESHOLD or
            abs((p.y() + p.height()) - (x_snap_partner.y() + x_snap_partner.height())) < SNAP_THRESHOLD
        ))
        already_linked = p._link_group is not None and x_snap_partner in p._link_group
        if x_snap_partner is not None and not already_linked and y_aligned:
            left_win = p if p.x() <= x_snap_partner.x() else x_snap_partner
            right_win = x_snap_partner if left_win is p else p
            if left_win._merge_btn is None or left_win._merge_btn._win_b is not right_win:
                if left_win._merge_btn is not None and not left_win._merge_btn._merged:
                    left_win._merge_btn.hide()
                btn = MergeButton()
                btn.set_windows(left_win, right_win)
                left_win._merge_btn = btn
            btn = left_win._merge_btn
            btn.set_merged(False)
            btn.update_position(btn._win_a, btn._win_b)
            btn.show()
            btn._raise_timer.start()
        elif x_snap_partner is None:
            if p._merge_btn is not None and not p._merge_btn._merged:
                p._merge_btn.hide()
                p._merge_btn = None

    def mouseReleaseEvent(self, e):
        self._drag_pos = None
        self.parent._snapped_pos = None
        self.parent._current_snap_edges = []
        self.parent._snap_is_adj = False
        self.parent.save_state()
        if self.parent._link_group:
            for member in self.parent._link_group:
                if member is not self.parent:
                    member.save_state()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.parent.toggle_shade()


class RotatedMenuButton(QPushButton):
    """'…' 을 90도 회전해서 세로로 표시하는 버튼"""
    hovered = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__('…', parent)
        self._hovered = False

    def sizeHint(self):
        fm = QFontMetrics(self.font())
        char_w = fm.horizontalAdvance('…')
        side = char_w + 12
        return QSize(side, side)

    def enterEvent(self, e):
        self._hovered = True
        self.hovered.emit(True)
        self.update()

    def leaveEvent(self, e):
        self._hovered = False
        self.hovered.emit(False)
        self.update()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        if self._hovered:
            p.setBrush(QColor(0, 0, 0, 30))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(self.rect(), 4, 4)

        p.translate(self.width() / 2.0, self.height() / 2.0)
        p.rotate(90)
        p.setFont(self.font())
        p.setPen(QColor(0, 0, 0))
        p.drawText(
            -self.height() // 2, -self.width() // 2,
            self.height(), self.width(),
            Qt.AlignCenter, self.text()
        )
        p.end()


class _AutoHeightEdit(QPlainTextEdit):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.document().setDocumentMargin(0)
        opt = QTextOption()
        opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        self.document().setDefaultTextOption(opt)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._fixed_height = 28
        self.document().contentsChanged.connect(self._update_height)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.clearFocus()
        else:
            super().keyPressEvent(e)

    def sizeHint(self):
        sh = super().sizeHint()
        return QSize(sh.width(), self._fixed_height)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._update_height()

    def _update_height(self):
        width = self.viewport().width()
        if width <= 0:
            return
        text = self.toPlainText() or ' '
        layout = QTextLayout(text, self.font())
        opt = QTextOption()
        opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        layout.setTextOption(opt)
        layout.beginLayout()
        y = 0.0
        while True:
            line = layout.createLine()
            if not line.isValid():
                break
            line.setLineWidth(width)
            line.setPosition(QPointF(0, y))
            y += line.height()
        layout.endLayout()
        doc_margin = self.document().documentMargin()
        h = int(y + 2 * doc_margin) + 4
        fixed_h = max(h, 28)
        self._fixed_height = fixed_h
        self.setMinimumHeight(fixed_h)
        self.setMaximumHeight(fixed_h)
        v_pad = max(0, (fixed_h - int(y)) // 2)
        self.setViewportMargins(0, v_pad, 0, v_pad)


class _NoteButton(QPushButton):
    """chat_20_regular 아이콘을 x축 대칭(상하 반전)으로 그리는 메모 버튼."""
    _CODEPOINT = ''

    def __init__(self, scale=1.0):
        super().__init__()
        self._color_normal = QColor('#cccccc')
        self._color_hover  = QColor('#d4b800')
        self._hovered      = False
        self.setFlat(True)
        sz = int(22 * scale)
        self.setFixedSize(sz, sz)
        self.setFont(mi_font(int(14 * scale)))
        self.setCursor(Qt.PointingHandCursor)

    def set_note_color(self, has_note: bool):
        self._color_normal = QColor('#c8a800' if has_note else '#cccccc')
        self.update()

    def enterEvent(self, e):
        self._hovered = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hovered = False
        self.update()
        super().leaveEvent(e)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(self._color_hover if self._hovered else self._color_normal)
        p.setFont(self.font())
        p.drawText(self.rect(), Qt.AlignCenter, self._CODEPOINT)
        p.end()


class TaskNotePopup(QWidget):
    """업무별 메모 플로팅 창."""

    def __init__(self, task, scale=1.0, on_note_changed=None):
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint)
        self._task_id = task['id']
        self._on_note_changed = on_note_changed
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setObjectName('TaskNotePopup')
        self.setStyleSheet('QWidget#TaskNotePopup { background: #fffacd; border: 1px solid #d4b800; }')

        vlay = QVBoxLayout(self)
        vlay.setContentsMargins(8, 6, 8, 8)
        vlay.setSpacing(4)

        title_font = pr_font(int(10 * scale))
        title_font.setBold(True)
        title = QLabel()
        title.setFont(title_font)
        title.setStyleSheet('color: #444; background: transparent; border: none;')
        fm = QFontMetrics(title_font)
        title.setText(fm.elidedText(task['name'], Qt.ElideRight, int(150 * scale)))

        btn_done = QPushButton('ctrl+enter')
        btn_done.setFont(pr_font(int(9 * scale)))
        btn_done.setFixedHeight(int(20 * scale))
        btn_done.setCursor(Qt.PointingHandCursor)
        btn_done.setStyleSheet("""
            QPushButton { background: #d4b800; color: white; border: none; border-radius: 4px; padding: 0 8px; }
            QPushButton:hover { background: #b89e00; }
            QPushButton:pressed { background: #a08800; }
        """)
        btn_done.clicked.connect(self.close)

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(6)
        title_row.addWidget(title, 1)
        title_row.addWidget(btn_done, 0)
        vlay.addLayout(title_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet('border: none; border-top: 1px solid #e0c800; background: transparent;')
        sep.setFixedHeight(1)
        vlay.addWidget(sep)

        self._editor = QPlainTextEdit()
        self._editor.setPlaceholderText('업무 관련 메모를 적으세요...')
        self._editor.setFont(pr_font(int(10 * scale)))
        self._editor.setStyleSheet(
            'QPlainTextEdit { background: transparent; border: none; color: #333; }'
        )
        vlay.addWidget(self._editor)

        self.setFixedSize(int(240 * scale), int(160 * scale))
        self._editor.setPlainText(get_task_notes(self._task_id))

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.timeout.connect(self._flush_save)
        self._editor.textChanged.connect(lambda: self._save_timer.start(500))

        QApplication.instance().focusChanged.connect(self._on_focus_changed)

    def _on_focus_changed(self, old, new):
        if new is None:
            return
        w = new
        while w is not None:
            if w is self:
                return
            w = w.parent()
        QTimer.singleShot(100, self._close_if_unfocused)

    def _close_if_unfocused(self):
        if not self.isVisible():
            return
        focused = QApplication.focusWidget()
        if focused is not None:
            w = focused
            while w is not None:
                if w is self:
                    return
                w = w.parent()
        self.close()

    def _flush_save(self):
        text = self._editor.toPlainText()
        set_task_notes(self._task_id, text)
        if self._on_note_changed:
            self._on_note_changed(text)

    def closeEvent(self, event):
        try:
            QApplication.instance().focusChanged.disconnect(self._on_focus_changed)
        except Exception:
            pass
        self._flush_save()
        super().closeEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        elif event.modifiers() & Qt.ControlModifier and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.close()
        else:
            super().keyPressEvent(event)


class TaskRow(QWidget):
    def __init__(self, task, on_delete, on_update, scale=1.0):
        super().__init__()
        self.task      = task
        self.on_update = on_update
        self._scale    = scale
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setObjectName('TaskRow')
        self.setStyleSheet('QWidget#TaskRow { background: transparent; }')

        layout = QHBoxLayout(self)
        layout.setContentsMargins(9, 0, 8, 0)
        layout.setSpacing(4)

        # 별 버튼 (우선순위)
        _is_starred = bool(task.get('priority', 0))
        self._star_sz = int(22 * scale)
        _star_pt = int(15 * scale)
        self._star_svg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'ic_fluent_star_emphasis_20_filled.svg')
        self.btn_star = QPushButton()
        self.btn_star.setFixedSize(self._star_sz, self._star_sz)
        self.btn_star.setFlat(True)
        self.btn_star.setStyleSheet("""
            QPushButton { border: none; padding: 0; }
            QPushButton:pressed { padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }
        """)
        if _is_starred:
            self.btn_star.setIcon(_svg_icon(self._star_svg, '#D4A800', self._star_sz))
            self.btn_star.setIconSize(QSize(self._star_sz, self._star_sz))
        else:
            self.btn_star.setText(MI.STAR_BORDER)
            self.btn_star.setFont(mi_font(_star_pt))
            self.btn_star.setStyleSheet("""
                QPushButton { color: #AAAAAA; border: none; padding: 0; }
                QPushButton:pressed { padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }
            """)
        self.btn_star.clicked.connect(self._toggle_priority)
        layout.addWidget(self.btn_star, 0, Qt.AlignVCenter)
        layout.setAlignment(self.btn_star, Qt.AlignVCenter)

        deadline = task['deadline']
        overdue = False
        urgent = False
        if deadline:
            dday = calc_dday(deadline)
            if dday.startswith('D+'):
                overdue = True
                dday_color = '#aaa'
            elif dday == 'D-day' or '시간' in dday:
                dday_color = '#e74c3c'
                urgent = True
            elif dday == '날짜오류':
                dday_color = '#aaa'
            else:
                days_left = int(dday[2:])
                if days_left <= 5:
                    dday_color = '#e74c3c'
                    urgent = True
                else:
                    dday_color = '#333'
            suffix = f'({dday})'
        else:
            dday_color = '#333'
            suffix = ''

        self._strikethrough = bool(task.get('strikethrough', 0))
        name_color = '#aaa' if overdue else '#111'

        self.name_edit = _AutoHeightEdit(task['name'])
        _base_pt = 12 * scale
        name_font = pr_font(12)
        name_font.setPointSizeF((11 / 12) * _base_pt if overdue else _base_pt)
        if overdue:
            pass
        elif urgent:
            name_font.setBold(True)
        elif _is_starred:
            name_font.setBold(True)
        self.name_edit.setFont(name_font)
        self.name_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background: transparent;
                border: none;
                color: {name_color};
                padding: 0px;
            }}
            QPlainTextEdit:focus {{
                background: rgba(255,255,255,0.6);
                border-bottom: 1px solid #d4b800;
            }}
        """)
        if self._strikethrough:
            font = self.name_edit.font()
            font.setStrikeOut(True)
            self.name_edit.setFont(font)
        self.name_edit.installEventFilter(self)

        if suffix:
            try:
                d = date.fromisoformat(deadline[:10])
                weekday = ['월','화','수','목','금','토','일'][d.weekday()]
                base_date_text = f'{d.year}. {d.month}. {d.day}.({weekday})'
                time_suffix = f' {deadline[11:]}' if len(deadline) > 10 else ''
                if dday == 'D-day':
                    date_text = f'{deadline[11:]} 까지' if len(deadline) > 10 else '언능 하소!'
                else:
                    date_text = f'{base_date_text}{time_suffix}'
            except ValueError:
                date_text = suffix
            dday_lbl = DdayLabel(suffix, date_text)
            dday_font = pr_font(12)
            dday_font.setBold(urgent or _is_starred)
            dday_font.setPointSizeF((11 / 12) * _base_pt if overdue else _base_pt)
            dday_lbl.setFont(dday_font)
            dday_lbl.setStyleSheet(f'color: {dday_color}; background: transparent;')
        else:
            dday_lbl = None
            self._add_date_btn = AddDateButton()
            self._add_date_btn.clicked.connect(self._show_date_picker)

        btn_menu = RotatedMenuButton()
        _f = btn_menu.font()
        _f.setPixelSize(19)
        btn_menu.setFont(_f)
        btn_menu.setStyleSheet('background: transparent; border: none; padding: 4px 6px; border-radius: 4px;')
        btn_menu.clicked.connect(lambda: self._open_task_menu(btn_menu, task, on_delete))
        btn_menu.hovered.connect(self._set_row_highlight)

        _has_note = bool(task.get('notes', ''))
        self.btn_note = _NoteButton(scale=scale)
        self.btn_note.set_note_color(_has_note)
        self.btn_note.setToolTip('업무 메모')
        self.btn_note.clicked.connect(self._open_note_popup)

        layout.addWidget(self.name_edit, 1, Qt.AlignVCenter)
        if task.get('recurrence', ''):
            lbl_recur = QLabel(MAT.REPEAT)
            _recur_font = mat_font(int(13 * scale))
            lbl_recur.setFont(_recur_font)
            lbl_recur.setStyleSheet('color: #333; background: transparent;')
            _fm = QFontMetrics(_recur_font)
            lbl_recur.setFixedWidth(_fm.horizontalAdvance(MAT.REPEAT) + 4)
            icon_dday_layout = QHBoxLayout()
            icon_dday_layout.setContentsMargins(0, 0, 0, 0)
            icon_dday_layout.setSpacing(0)
            icon_dday_layout.addWidget(lbl_recur, 0, Qt.AlignVCenter)
            if dday_lbl:
                icon_dday_layout.addWidget(dday_lbl, 0, Qt.AlignVCenter)
            elif hasattr(self, '_add_date_btn'):
                icon_dday_layout.addWidget(self._add_date_btn, 0, Qt.AlignVCenter)
            icon_dday_widget = QWidget()
            icon_dday_widget.setLayout(icon_dday_layout)
            layout.addWidget(icon_dday_widget, 0, Qt.AlignVCenter)
        else:
            if dday_lbl:
                layout.addWidget(dday_lbl, 0, Qt.AlignVCenter)
            elif hasattr(self, '_add_date_btn'):
                layout.addWidget(self._add_date_btn, 0, Qt.AlignVCenter)
        layout.addWidget(self.btn_note, 0, Qt.AlignTop)
        layout.addWidget(btn_menu, 0, Qt.AlignVCenter)

    def _set_row_highlight(self, on: bool):
        if on:
            self.setStyleSheet('QWidget#TaskRow { background: rgba(0,0,0,18); }')
        else:
            self.setStyleSheet('QWidget#TaskRow { background: transparent; }')

    def eventFilter(self, watched, event):
        if watched is self.name_edit and event.type() == QEvent.FocusOut:
            self._save()
        return super().eventFilter(watched, event)

    def _save(self):
        name = self.name_edit.toPlainText().strip()
        if not name:
            self.name_edit.setPlainText(self.task['name'])
            return
        if name != self.task['name']:
            self.task['name'] = name
            update_task(self.task['id'], name, self.task['deadline'],
                        strikethrough=int(self._strikethrough),
                        priority=self.task.get('priority', 0),
                        recurrence=self.task.get('recurrence', ''))
            self.on_update()

    def _open_task_menu(self, btn, task, on_delete):
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background: #fffacd;
                border: 1px solid #d4b800;
                border-radius: 6px;
                padding: 4px 0;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                color: #333;
                min-width: 120px;
            }
            QMenu::item { padding: 6px 16px 6px 12px; }
            QMenu::item:selected { background: rgba(212,184,0,0.3); }
        """)
        sub_recur = QMenu('반복 설정', self)
        sub_recur.setFont(mi_font(10))
        sub_recur.setStyleSheet(menu.styleSheet().replace('min-width: 120px', 'min-width: 0px').replace(
            'QMenu::item { padding: 6px 16px 6px 12px; }',
            'QMenu::item { padding: 6px 16px 6px 16px; text-align: center; }'
        ))
        cur_recur = self.task.get('recurrence', '')
        for _label, _val in [('없음', ''), ('매주', 'weekly'), ('격주', 'biweekly'), ('매월', 'monthly'), ('매년', 'yearly')]:
            _act = QAction((MI.CHECK + ' ' if cur_recur == _val else '    ') + _label, self)
            _act.triggered.connect(lambda _, v=_val: self._set_recurrence(v))
            sub_recur.addAction(_act)
        _is_custom = cur_recur.startswith('custom:')
        _act_custom = QAction((MI.CHECK + ' ' if _is_custom else '    ') + '사용자 설정...', self)
        _act_custom.triggered.connect(self._open_custom_recurrence)
        sub_recur.addAction(_act_custom)
        menu.addMenu(sub_recur)

        menu.addSeparator()
        act_strike = QAction('가운데 선(C)', self)
        act_delete = QAction('삭제(Delete/D)', self)
        menu.addAction(act_strike)
        menu.addAction(act_delete)

        act_strike.triggered.connect(self._toggle_strikethrough)
        act_delete.triggered.connect(lambda: on_delete(task))

        def _key_press(e):
            if e.key() == Qt.Key_C:
                act_strike.trigger()
                menu.close()
            elif e.key() in (Qt.Key_Delete, Qt.Key_D):
                act_delete.trigger()
                menu.close()
            else:
                QMenu.keyPressEvent(menu, e)

        menu.keyPressEvent = _key_press

        pos = btn.mapToGlobal(btn.rect().bottomLeft())
        menu.exec_(pos)

    def _toggle_strikethrough(self):
        self._strikethrough = not self._strikethrough
        font = self.name_edit.font()
        font.setStrikeOut(self._strikethrough)
        self.name_edit.setFont(font)
        self.name_edit._update_height()
        update_task(self.task['id'], self.task['name'], self.task['deadline'],
                    strikethrough=int(self._strikethrough),
                    priority=self.task.get('priority', 0),
                    recurrence=self.task.get('recurrence', ''))

    def _open_note_popup(self):
        if getattr(self, '_note_popup', None) is not None:
            self._note_popup.close()
            return
        popup = TaskNotePopup(self.task, scale=self._scale, on_note_changed=self._update_note_btn)
        popup.destroyed.connect(lambda: setattr(self, '_note_popup', None))
        pos = self.btn_note.mapToGlobal(QPoint(0, self.btn_note.height()))
        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        sr = screen.availableGeometry()
        x, y = pos.x(), pos.y()
        pw, ph = popup.width(), popup.height()
        if x + pw > sr.right():
            x = sr.right() - pw
        if y + ph > sr.bottom():
            y = self.btn_note.mapToGlobal(QPoint(0, -ph)).y()
        popup.move(x, y)
        popup.show()
        popup._editor.setFocus()
        self._note_popup = popup

    def _update_note_btn(self, notes_text):
        self.btn_note.set_note_color(bool(notes_text.strip()))

    def _show_date_picker(self):
        if not hasattr(self, '_cal_popup'):
            self._cal_popup = CustomCalendarWidget()
            self._cal_popup.setWindowFlags(Qt.Popup)
            self._cal_popup.setStyleSheet("""
                QCalendarWidget { background-color: white; }
                QCalendarWidget QAbstractItemView {
                    background-color: white; color: black;
                    font-family: 'Malgun Gothic'; font-size: 10pt;
                    selection-background-color: #d4b800; selection-color: white;
                }
                QCalendarWidget QWidget { background-color: white; }
                QCalendarWidget QToolButton {
                    background-color: white; color: black;
                    font-family: 'Malgun Gothic'; font-size: 10pt;
                }
                QCalendarWidget QWidget#qt_calendar_navigationbar { background-color: white; }
                QCalendarWidget QSpinBox { background-color: white; color: black; }
            """)
            fmt_sat = QTextCharFormat()
            fmt_sat.setForeground(QColor('#0055cc'))
            self._cal_popup.setWeekdayTextFormat(Qt.Saturday, fmt_sat)
            fmt_sun = QTextCharFormat()
            fmt_sun.setForeground(QColor('#cc0000'))
            self._cal_popup.setWeekdayTextFormat(Qt.Sunday, fmt_sun)
            self._cal_popup.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
            self._cal_popup.clicked.connect(self._on_date_selected)

        today = QDate.currentDate()
        self._cal_popup.setCurrentPage(today.year(), today.month())
        self._cal_popup.setSelectedDate(today)

        btn = self._add_date_btn
        cal_size = self._cal_popup.sizeHint()
        pos = btn.mapToGlobal(QPoint(0, btn.height()))

        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        screen_rect = screen.availableGeometry()
        if pos.x() + cal_size.width() > screen_rect.right():
            pos.setX(screen_rect.right() - cal_size.width())
        if pos.y() + cal_size.height() > screen_rect.bottom():
            pos = btn.mapToGlobal(QPoint(0, -cal_size.height()))

        self._cal_popup.move(pos)
        self._cal_popup.show()

    def _on_date_selected(self, qdate):
        self._cal_popup.hide()
        deadline_str = qdate.toString('yyyy-MM-dd')
        self.task['deadline'] = deadline_str
        update_task(self.task['id'], self.task['name'], deadline_str,
                    strikethrough=int(self._strikethrough),
                    priority=self.task.get('priority', 0),
                    recurrence=self.task.get('recurrence', ''))
        self.on_update()

    def _toggle_priority(self):
        new = 0 if self.task.get('priority', 0) else 1
        self.task['priority'] = new
        if new:
            self.btn_star.setText('')
            self.btn_star.setIcon(_svg_icon(self._star_svg, '#D4A800', self._star_sz))
            self.btn_star.setIconSize(QSize(self._star_sz, self._star_sz))
            self.btn_star.setStyleSheet("""
                QPushButton { border: none; padding: 0; }
                QPushButton:pressed { padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }
            """)
        else:
            self.btn_star.setIcon(QIcon())
            self.btn_star.setText(MI.STAR_BORDER)
            self.btn_star.setFont(mi_font(int(15 * self._scale)))
            self.btn_star.setStyleSheet("""
                QPushButton { color: #AAAAAA; border: none; padding: 0; }
                QPushButton:pressed { padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }
            """)
        set_task_priority(self.task['id'], new)
        self.on_update()

    def _set_recurrence(self, recurrence):
        self.task['recurrence'] = recurrence
        set_task_recurrence(self.task['id'], recurrence)
        self.on_update()

    def _open_custom_recurrence(self):
        dlg = CustomRecurrenceDialog(self, current=self.task.get('recurrence', ''))
        if dlg.exec_() == QDialog.Accepted:
            self._set_recurrence(dlg.get_value())


class _FixedDotBar(QWidget):
    """드래그 불가 고정 구분선 (황금색 바 + 점 장식)."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(10)
        self._color = '#d4b800'

    def set_color(self, hex_color):
        self._color = hex_color
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect().adjusted(1, 0, -1, 0), QColor(self._color))
        painter.end()


class CustomCalendarWidget(QCalendarWidget):
    def paintCell(self, painter, rect, date):
        if date.month() != self.monthShown():
            return
        super().paintCell(painter, rect, date)
        if date == QDate.currentDate():
            painter.save()
            painter.setPen(QColor('#d4b800'))
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(rect.adjusted(1, 1, -2, -2))
            painter.restore()


class GlobalSearchDialog(QDialog):
    def __init__(self, open_windows, parent=None):
        super().__init__(parent, Qt.WindowStaysOnTopHint)
        self.setWindowTitle('통합 검색')
        self.setMinimumSize(440, 380)
        self._open_windows = open_windows

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        self._input = QLineEdit()
        self._input.setPlaceholderText('검색어 입력…')
        self._input.setFont(QFont('Malgun Gothic', 11))
        self._input.setStyleSheet(
            "QLineEdit { border: 1px solid #d4b800; border-radius: 4px; padding: 4px 8px; }"
        )
        layout.addWidget(self._input)

        self._list = QListWidget()
        self._list.setFont(QFont('Malgun Gothic', 10))
        self._list.setAlternatingRowColors(True)
        self._list.setStyleSheet(
            "QListWidget { border: 1px solid #ddd; border-radius: 4px; }"
            "QListWidget::item { padding: 5px 8px; }"
            "QListWidget::item:selected { background: rgba(212,184,0,0.4); color: #333; }"
        )
        layout.addWidget(self._list)

        self._timer = QTimer()
        self._timer.setSingleShot(True)
        self._timer.setInterval(250)
        self._input.textChanged.connect(lambda: self._timer.start())
        self._timer.timeout.connect(self._do_search)
        self._list.itemDoubleClicked.connect(self._jump_to)

    def _do_search(self):
        q = self._input.text().strip()
        self._list.clear()
        if not q:
            return
        for r in search_tasks_all(q):
            dday = calc_dday(r['deadline']) if r.get('deadline') else ''
            text = f"[{dday}]  {r['name']}" if dday else r['name']
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, r)
            if r.get('color'):
                item.setBackground(QColor(r['color']).lighter(140))
            self._list.addItem(item)

    def _jump_to(self, item):
        r = item.data(Qt.UserRole)
        for win in self._open_windows:
            if win.window_id == r['window_id']:
                win.show()
                win.raise_()
                win.activateWindow()
                win._highlight_task(r['id'])
                break
        self.accept()


class TimePickerPopup(QFrame):
    time_selected = pyqtSignal(QTime)

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.setStyleSheet("""
            QFrame {
                background: #fffbe6;
                border: 1px solid #d4b800;
                border-radius: 6px;
            }
            QListWidget {
                background: transparent;
                border: none;
                font-family: 'Malgun Gothic';
                font-size: 11pt;
                color: #333;
                outline: 0;
            }
            QListWidget::item { padding: 3px 10px; border-radius: 3px; }
            QListWidget::item:selected { background: #d4b800; color: white; }
            QListWidget::item:hover { background: #fff3a0; }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        self._hour_list = QListWidget()
        self._hour_list.setFixedWidth(52)
        self._hour_list.setFixedHeight(180)
        for h in range(24):
            _item = QListWidgetItem(f'{h:02d}')
            _item.setTextAlignment(Qt.AlignCenter)
            self._hour_list.addItem(_item)

        self._min_list = QListWidget()
        self._min_list.setFixedWidth(52)
        self._min_list.setFixedHeight(180)
        for m in range(0, 60, 5):
            _item = QListWidgetItem(f'{m:02d}')
            _item.setTextAlignment(Qt.AlignCenter)
            self._min_list.addItem(_item)

        layout.addWidget(self._hour_list)
        layout.addWidget(self._min_list)

        self._selected_hour = 9
        self._hour_list.itemClicked.connect(self._on_hour_clicked)
        self._min_list.itemClicked.connect(self._on_min_clicked)

    def set_current_time(self, qtime):
        self._selected_hour = qtime.hour()
        self._hour_list.setCurrentRow(qtime.hour())
        self._min_list.setCurrentRow(qtime.minute() // 5)
        self._hour_list.scrollToItem(self._hour_list.currentItem(), QAbstractItemView.PositionAtCenter)
        self._min_list.scrollToItem(self._min_list.currentItem(), QAbstractItemView.PositionAtCenter)

    def _on_hour_clicked(self, item):
        self._selected_hour = int(item.text())

    def _on_min_clicked(self, item):
        self.time_selected.emit(QTime(self._selected_hour, int(item.text())))
        self.hide()


class ClickToCopyLabel(QLabel):
    """클릭하면 텍스트를 클립보드에 복사하고 '복사됨!' 툴팁을 표시하는 레이블."""

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.text():
            QApplication.clipboard().setText(self.text())
            self._show_copied_popup(event.globalPos())
        super().mousePressEvent(event)

    def _show_copied_popup(self, global_pos):
        popup = QLabel('복사됨!', self, Qt.ToolTip | Qt.FramelessWindowHint)
        popup.setStyleSheet("""
            QLabel {
                background: #5a4000;
                color: white;
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """)
        popup.adjustSize()
        popup.move(global_pos + QPoint(4, -popup.height() - 8))
        popup.show()
        QTimer.singleShot(500, popup.deleteLater)


class SmartDocLineEdit(QLineEdit):
    """텍스트가 있으면 클릭 시 복사, 비어있으면 _on_empty_click 호출."""

    def __init__(self, text=''):
        super().__init__(text)
        self._on_empty_click = None
        self._editing = False
        self._cancel_edit = None
        self.setCursorPosition(0)

    def mousePressEvent(self, event):
        if self._editing:
            super().mousePressEvent(event)
            return
        if event.button() == Qt.LeftButton:
            if self.text():
                QApplication.clipboard().setText(self.text())
                self._show_popup(event.globalPos())
            elif self._on_empty_click:
                self._on_empty_click()
        super().mousePressEvent(event)

    def keyPressEvent(self, e):
        if self._editing and e.key() == Qt.Key_Escape:
            if self._cancel_edit:
                self._cancel_edit()
            return
        super().keyPressEvent(e)

    def _show_popup(self, global_pos):
        popup = QLabel('복사됨!', self, Qt.ToolTip | Qt.FramelessWindowHint)
        popup.setStyleSheet("""
            QLabel {
                background: #5a4000;
                color: white;
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """)
        popup.adjustSize()
        popup.move(global_pos + QPoint(4, -popup.height() - 8))
        popup.show()
        QTimer.singleShot(500, popup.deleteLater)


class ClickToCopyLineEdit(QLineEdit):
    """클릭하면 텍스트를 클립보드에 복사하고 '복사됨!' 툴팁을 표시하는 읽기전용 필드."""

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.text():
            QApplication.clipboard().setText(self.text())
            self._show_copied_popup(event.globalPos())
        super().mousePressEvent(event)

    def _show_copied_popup(self, global_pos):
        popup = QLabel('복사됨!', self, Qt.ToolTip | Qt.FramelessWindowHint)
        popup.setStyleSheet("""
            QLabel {
                background: #5a4000;
                color: white;
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """)
        popup.adjustSize()
        popup.move(global_pos + QPoint(4, -popup.height() - 8))
        popup.show()
        QTimer.singleShot(500, popup.deleteLater)



class DocumentRow(QWidget):
    """공문번호 한 항목 행: [공문제목 입력] [공문번호(읽기전용)] [삭제메뉴]"""

    def __init__(self, doc: dict, on_delete, on_paste=None):
        super().__init__()
        self._doc = doc
        self._on_delete = on_delete
        self.setStyleSheet('background: transparent;')

        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 2, 8, 2)
        lay.setSpacing(0)

        field_style = """
            QLineEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """
        readonly_style = """
            QLineEdit {
                background: rgba(255,255,255,0.25);
                border: 1px solid #c0a800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #5a4000;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """

        editing_style = """
            QLineEdit {
                background: rgba(255,255,255,0.9);
                border: 1px solid #a08800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #111;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """

        self.title_edit = SmartDocLineEdit(doc.get('title', ''))
        self.title_edit.setPlaceholderText('공문 제목 붙여넣는 곳')
        self.title_edit.setReadOnly(True)
        self.title_edit.setStyleSheet(field_style)
        self.title_edit.setCursor(Qt.PointingHandCursor)
        lay.addWidget(self.title_edit, 11)
        lay.addSpacing(0)

        import os
        from PyQt5.QtGui import QIcon, QPixmap
        from PyQt5.QtCore import QSize
        _edit_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', '수정 아이콘.png')
        btn_edit = QPushButton()
        btn_edit.setIcon(QIcon(QPixmap(_edit_icon_path).scaled(15, 16, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)))
        btn_edit.setIconSize(QSize(15, 16))
        btn_edit.setFixedSize(17, 18)
        btn_edit.setToolTip('')
        btn_edit.setCursor(Qt.PointingHandCursor)
        btn_edit.setStyleSheet("""
            QPushButton { background: transparent; border: none; padding: 0; }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; }
        """)
        lay.addWidget(btn_edit)
        lay.addSpacing(1)

        self.num_edit = SmartDocLineEdit(doc.get('doc_number', ''))
        self.num_edit.setReadOnly(True)
        self.num_edit.setPlaceholderText('공문 번호 붙여넣는 곳')
        self.num_edit.setStyleSheet(readonly_style)
        self.num_edit.setCursor(Qt.PointingHandCursor)
        lay.addWidget(self.num_edit, 11)

        btn_edit_num = QPushButton()
        btn_edit_num.setIcon(QIcon(QPixmap(_edit_icon_path).scaled(15, 16, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)))
        btn_edit_num.setIconSize(QSize(15, 16))
        btn_edit_num.setFixedSize(17, 18)
        btn_edit_num.setToolTip('')
        btn_edit_num.setCursor(Qt.PointingHandCursor)
        btn_edit_num.setStyleSheet("""
            QPushButton { background: transparent; border: none; padding: 0; }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; }
        """)
        lay.addWidget(btn_edit_num)
        lay.addSpacing(4)

        def _enter_edit():
            self.title_edit._editing = True
            self.title_edit.setReadOnly(False)
            self.title_edit.setStyleSheet(editing_style)
            self.title_edit.setCursor(Qt.IBeamCursor)
            self.title_edit.setFocus()
            self.title_edit.setCursorPosition(0)

        def _commit_edit():
            self.title_edit._editing = False
            self.title_edit.setReadOnly(True)
            self.title_edit.setStyleSheet(field_style)
            self.title_edit.setCursor(Qt.PointingHandCursor)
            self.title_edit.setCursorPosition(0)
            if doc.get('id') is not None:
                self._save()

        _original_title = doc.get('title', '')

        def _cancel_edit():
            self.title_edit.setText(_original_title)
            _commit_edit()

        self.title_edit._cancel_edit = _cancel_edit
        self.title_edit.returnPressed.connect(_commit_edit)
        btn_edit.clicked.connect(lambda: _commit_edit() if self.title_edit._editing else _enter_edit())

        def _enter_edit_num():
            self.num_edit._editing = True
            self.num_edit.setReadOnly(False)
            self.num_edit.setStyleSheet(editing_style)
            self.num_edit.setCursor(Qt.IBeamCursor)
            self.num_edit.setFocus()
            self.num_edit.setCursorPosition(0)

        def _commit_edit_num():
            self.num_edit._editing = False
            self.num_edit.setReadOnly(True)
            self.num_edit.setStyleSheet(readonly_style)
            self.num_edit.setCursor(Qt.PointingHandCursor)
            self.num_edit.setCursorPosition(0)
            if doc.get('id') is not None:
                self._save()

        _original_num = doc.get('doc_number', '')

        def _cancel_edit_num():
            self.num_edit.setText(_original_num)
            _commit_edit_num()

        self.num_edit._cancel_edit = _cancel_edit_num
        self.num_edit.returnPressed.connect(_commit_edit_num)
        btn_edit_num.clicked.connect(lambda: _commit_edit_num() if self.num_edit._editing else _enter_edit_num())

        if on_paste:
            self.title_edit._on_empty_click = lambda: on_paste('title', self.title_edit, self.num_edit)
            self.num_edit._on_empty_click = lambda: on_paste('number', self.title_edit, self.num_edit)

        if doc.get('id') is None:
            lay.addSpacing(22)  # 삭제 버튼(22px) 자리 확보 → 위 행과 폭 통일
            return

        import os
        from PyQt5.QtGui import QIcon, QPixmap
        from PyQt5.QtCore import QSize
        _del_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', '엑스아이콘.png')
        btn_del = QPushButton()
        btn_del.setIcon(QIcon(QPixmap(_del_icon_path).scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation)))
        btn_del.setIconSize(QSize(18, 18))
        btn_del.setFixedSize(22, 26)
        btn_del.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                padding: 0;
            }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; padding: 2px 0 0 2px; }
        """)
        btn_del.clicked.connect(lambda: self._on_delete(self._doc['id']))
        lay.addWidget(btn_del)

    def _save(self):
        update_document(self._doc['id'], self.title_edit.text(), self.num_edit.text())


class UrgentToast(QWidget):
    """마감 임박 태스크를 우하단에 10초간 표시하는 토스트 팝업."""

    def __init__(self, tasks):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_ShowWithoutActivating)

        self.setStyleSheet("""
            QWidget#toast {
                background: #fff0e6;
            }
            QLabel { background: transparent; border: none; }
        """)
        self.setObjectName('toast')

        # 팝업 이미지 로드 (assets/popup_images/ 폴더에서 랜덤 선택)
        import random as _random
        _base = os.path.dirname(os.path.abspath(__file__))
        _popup_dir = os.path.join(_base, 'assets', 'popup_images')
        _exts = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        _candidates = sorted([
            os.path.join(_popup_dir, f)
            for f in os.listdir(_popup_dir)
            if os.path.splitext(f)[1].lower() in _exts
        ]) if os.path.isdir(_popup_dir) else []
        _yuju_path = _random.choice(_candidates) if _candidates else None
        face_pix = None
        try:
            if not _yuju_path:
                raise FileNotFoundError
            from PIL import Image, ImageEnhance
            import io
            _pil = Image.open(_yuju_path).convert('RGBA')
            face_h_px = int(_pil.height * 0.72)
            _pil = _pil.crop((0, 0, _pil.width, face_h_px))
            _pil = ImageEnhance.Brightness(_pil).enhance(1.2)
            _pil = ImageEnhance.Color(_pil).enhance(1.1)
            buf = io.BytesIO()
            _pil.save(buf, format='PNG')
            _yuju_pix = QPixmap()
            _yuju_pix.loadFromData(buf.getvalue())
            face_pix = _yuju_pix.scaledToHeight(130, Qt.SmoothTransformation)
        except Exception:
            if _yuju_path:
                _yuju_pix = QPixmap(_yuju_path)
                if not _yuju_pix.isNull():
                    face_h = int(_yuju_pix.height() * 0.72)
                    _tmp = _yuju_pix.copy(0, 0, _yuju_pix.width(), face_h)
                    face_pix = _tmp.scaledToHeight(130, Qt.SmoothTransformation)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(27, 18, 27, 22)
        layout.setSpacing(11)

        header_row = QHBoxLayout()
        header_row.setSpacing(9)
        header = QLabel('마감 임박')
        header.setFont(QFont('Malgun Gothic', 16, QFont.Bold))
        header.setStyleSheet('color: #7a3000;')
        header_row.addWidget(header)
        header_row.addStretch()
        layout.addLayout(header_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet('background: #e8a070; max-height: 2px; border: none;')
        layout.addWidget(sep)

        for task in tasks:
            dday = task.get('notice') or calc_dday(task['deadline'])
            name = task['name']
            if len(name) > 30:
                name = name[:29] + '…'
            row = QLabel(f'• {name}  <b>{dday}</b>')
            row.setFont(QFont('Malgun Gothic', 14))
            row.setStyleSheet('color: #3a2000;')
            layout.addWidget(row)

        self.setMinimumWidth(300)
        self.adjustSize()
        screen = QApplication.primaryScreen().availableGeometry()
        tx = screen.right() - self.width() - 24
        ty = screen.bottom() - self.height() - 52
        self.move(tx, ty)

        # 유자 이미지: 팝업 우상단 밖으로 튀어나오는 별도 위젯
        self._face_widget = None
        if face_pix:
            fw = QLabel(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
            fw.setAttribute(Qt.WA_TranslucentBackground)
            fw.setAttribute(Qt.WA_ShowWithoutActivating)
            fw.setPixmap(face_pix)
            fw.setStyleSheet('background: transparent; border: none;')
            fw.resize(face_pix.size())
            # 팝업 가로 3/4 지점, 이미지 절반이 위로 튀어나오도록 배치
            fw.move(tx + self.width() * 3 // 4 - face_pix.width() // 2,
                    ty - face_pix.height() // 2)
            fw.show()
            fw.setCursor(Qt.PointingHandCursor)
            fw.mousePressEvent = lambda e: self._click_fade()
            self._face_widget = fw
            # 토스트가 show()된 뒤 이미지 위젯을 맨 앞으로
            QTimer.singleShot(0, fw.raise_)

        self.setCursor(Qt.PointingHandCursor)

        # 10초 후 페이드 아웃
        self._anim = QPropertyAnimation(self, b'windowOpacity')
        self._anim.setDuration(600)
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.setEasingCurve(QEasingCurve.InQuad)
        self._anim.finished.connect(self._on_fade_done)
        self._fade_timer = QTimer(self)
        self._fade_timer.setSingleShot(True)
        self._fade_timer.timeout.connect(self._anim.start)
        self._fade_timer.start(10000)

    def _on_fade_done(self):
        if self._face_widget:
            self._face_widget.close()
        self.close()

    def _click_fade(self):
        self._fade_timer.stop()
        self._anim.stop()
        self._anim.setDuration(600)
        self._anim.setStartValue(self.windowOpacity())
        self._anim.start()

    def paintEvent(self, event):
        from PyQt5.QtGui import QPainter, QColor
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        pen = painter.pen()
        pen.setColor(QColor('#d4637a'))
        pen.setWidth(5)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        # 안쪽으로 2px 들어와서 테두리가 잘리지 않게
        painter.drawRect(2, 2, self.width() - 4, self.height() - 4)
        painter.end()

    def mousePressEvent(self, event):
        self._click_fade()


class SnapGuideOverlay(QWidget):
    """드래그 중 정렬 가이드라인을 전체 화면 위에 그리는 투명 오버레이."""

    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self._lines = []
        self._fade_timer = QTimer(self)
        self._fade_timer.setSingleShot(True)
        self._fade_timer.timeout.connect(self._start_fade)
        self._anim = QPropertyAnimation(self, b'windowOpacity')
        self._anim.setDuration(350)
        self._anim.finished.connect(self.hide)
        screens = QApplication.screens()
        total = screens[0].geometry()
        for s in screens[1:]:
            total = total.united(s.geometry())
        self.setGeometry(total)

    def show_lines(self, lines):
        self._lines = lines
        if not lines:
            return
        self._fade_timer.stop()
        self._anim.stop()
        self.setWindowOpacity(1.0)
        self.update()
        self.show()
        self.raise_()

    def hide_lines(self):
        self._lines = []
        self._fade_timer.start(250)

    def _start_fade(self):
        self._anim.setStartValue(self.windowOpacity())
        self._anim.setEndValue(0.0)
        self._anim.start()

    def paintEvent(self, event):
        if not self._lines:
            return
        painter = QPainter(self)
        pen = painter.pen()
        pen.setColor(QColor(41, 121, 255, 180))
        pen.setWidth(1)
        painter.setPen(pen)
        for x1, y1, x2, y2 in self._lines:
            painter.drawLine(x1, y1, x2, y2)
        painter.end()


_snap_guide_overlay = None


def _get_snap_guide_overlay():
    global _snap_guide_overlay
    if _snap_guide_overlay is None:
        _snap_guide_overlay = SnapGuideOverlay()
    return _snap_guide_overlay


class MemoWindow(QMainWindow):
    def __init__(self, window_id, on_new=None, open_windows=None, on_toggle_hotkey=None,
                 on_alarm_interval_change=None, get_alarm_interval=None,
                 on_timed_alarm_change=None, get_timed_alarm_enabled=None,
                 on_shortcut_change=None, get_shortcut_enabled=None):
        super().__init__()
        self.window_id       = window_id
        self.on_new          = on_new or (lambda **kw: None)
        self._open_windows   = open_windows if open_windows is not None else []
        self._on_toggle_hotkey = on_toggle_hotkey
        self._on_alarm_interval_change = on_alarm_interval_change
        self._get_alarm_interval = get_alarm_interval
        self._on_timed_alarm_change = on_timed_alarm_change
        self._get_timed_alarm_enabled = get_timed_alarm_enabled
        self._on_shortcut_change = on_shortcut_change
        self._get_shortcut_enabled = get_shortcut_enabled
        self.collapsed       = False
        self.expanded_height = 400
        self.pin_active      = False
        self._bg_color       = '#FEFFA7'  # 기본값
        self._capture_hint_shown = False
        self._scale          = 1.0
        self._memo_mode      = False
        self._memo_save_timer = None
        self._snapped_pos        = None
        self._current_snap_edges = []
        self._snap_is_adj        = False  # 현재 인접 스냅 상태
        self._link_group         = None   # 합체된 창들의 set (자신 포함)
        self._is_syncing_move    = False  # moveEvent 무한루프 방지
        self._merge_btn          = None   # MergeButton 위젯 참조

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMinimumWidth(EdgeHandle.MIN_W)
        self._build_ui()
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self._pin_shortcut = QShortcut(QKeySequence('Ctrl+Shift+F'), self)
        self._pin_shortcut.activated.connect(self.toggle_always_on_top)
        self._shade_shortcut = QShortcut(QKeySequence('Ctrl+Shift+R'), self)
        self._shade_shortcut.activated.connect(self.toggle_shade)
        self._shortcuts = [self._pin_shortcut, self._shade_shortcut]
        # 저장된 단축키 상태 적용
        if get_shortcut_enabled is not None and not get_shortcut_enabled():
            for sc in self._shortcuts:
                sc.setEnabled(False)

    def apply_state(self, x, y, width, height, collapsed, color='', scale=1.0):
        self.expanded_height = height
        self.collapsed       = collapsed
        self.move(x, y)
        self.resize(max(width, EdgeHandle.MIN_W), TITLE_BAR_HEIGHT if collapsed else height)
        if collapsed:
            self.content.hide()
            self.setFixedHeight(TITLE_BAR_HEIGHT)
        if color:
            self._apply_color(color)
        if scale != 1.0:
            self._scale = scale
            self._apply_scale_to_inputs()
            self._refresh_tasks()

    def _build_ui(self):
        central = QWidget()
        central.setObjectName('central')
        central.setStyleSheet("""
            QWidget#central {
                background: #fffacd;
                border-radius: 0;
                border: 1px solid rgba(0,0,0,0.4);
            }
        """)
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.title_bar = TitleBar(self)
        root.addWidget(self.title_bar)

        self._title_sep = QFrame()
        self._title_sep.setFrameShape(QFrame.HLine)
        self._title_sep.setFixedHeight(1)
        self._title_sep.setStyleSheet('background: #D4AF37; border: none;')
        self._title_sep.hide()
        root.addWidget(self._title_sep)

        self.content = QWidget()
        self.content.setStyleSheet('background: transparent;')
        _outer_layout = QVBoxLayout(self.content)
        _outer_layout.setContentsMargins(0, 0, 0, 0)
        _outer_layout.setSpacing(0)

        _splitter = QWidget()
        _splitter_layout = QVBoxLayout(_splitter)
        _splitter_layout.setContentsMargins(0, 0, 0, 0)
        _splitter_layout.setSpacing(0)
        _outer_layout.addWidget(_splitter)

        _top_panel = QWidget()
        _top_panel.setStyleSheet('background: transparent;')
        content_layout = QVBoxLayout(_top_panel)
        content_layout.setContentsMargins(0, 6, 0, 0)
        content_layout.setSpacing(4)

        field_style_line = """
            QLineEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
            }
        """
        _date_base = """
            QDateEdit::drop-down { width: 0px; border: none; background: transparent; }
            QDateEdit::up-button { width: 0px; }
            QDateEdit::down-button { width: 0px; }
        """
        self._field_style_date_empty = """
            QDateEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #aaa;
                qproperty-alignment: AlignCenter;
            }
        """ + _date_base
        self._field_style_date_filled = """
            QDateEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
                qproperty-alignment: AlignCenter;
            }
        """ + _date_base
        field_style_date = self._field_style_date_empty
        lbl_font = QFont('Malgun Gothic', 11)
        lbl_css  = 'color: #5a4000; background: transparent;'

        # 1행: 업무
        self.task_input_row = QWidget()
        self.task_input_row.setStyleSheet('background: transparent;')
        row1 = QHBoxLayout(self.task_input_row)
        row1.setContentsMargins(16, 0, 8, 0)
        row1.setSpacing(4)
        lbl_name = QLabel('업무:')
        lbl_name.setFont(lbl_font)
        lbl_name.setStyleSheet(lbl_css)
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText('업무명')
        self.input_name.setFont(QFont('Malgun Gothic', 11))
        self.input_name.setStyleSheet(field_style_line)
        self._date_touched = False
        self._time_touched = False
        self.input_name.installEventFilter(self)
        self.input_name.returnPressed.connect(self._add_task)
        self.input_name.textChanged.connect(self._on_name_cleared)
        row1.addWidget(lbl_name)
        row1.addWidget(self.input_name, 1)
        content_layout.addWidget(self.task_input_row)

        # 2행: 마감일 (기본 숨김)
        self.date_row_widget = QWidget()
        self.date_row_widget.setStyleSheet('background: transparent;')
        row2 = QHBoxLayout(self.date_row_widget)
        row2.setContentsMargins(16, 0, 8, 0)
        row2.setSpacing(4)
        lbl_date = QLabel('마감일:')
        lbl_date.setFont(lbl_font)
        lbl_date.setStyleSheet(lbl_css)
        self.input_date = QDateEdit()
        self.input_date.setDisplayFormat('yyyy-MM-dd')
        self.input_date.setCalendarPopup(False)
        self.input_date.setMinimumDate(QDate(1900, 1, 1))
        self.input_date.setSpecialValueText(' ')
        self.input_date.setDate(QDate(1900, 1, 1))  # 빈 칸으로 표시
        self.input_date.setFixedWidth(115)
        self.input_date.setFont(QFont('Malgun Gothic', 11))
        self.input_date.setStyleSheet(field_style_date)
        self.input_date.dateChanged.connect(self._on_date_changed)
        self.input_date.installEventFilter(self)
        for child in self.input_date.findChildren(QWidget):
            child.installEventFilter(self)
        self.input_time = QTimeEdit(QTime(9, 0))
        self.input_time.setDisplayFormat('HH:mm')
        self.input_time.setFixedWidth(68)
        self.input_time.setFont(QFont('Malgun Gothic', 11))
        self.input_time.setEnabled(False)
        self.input_time.setStyleSheet("""
            QTimeEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
                qproperty-alignment: AlignCenter;
            }
            QTimeEdit:disabled {
                background: rgba(255,255,255,0.2);
                border: 1px solid #ddd;
                color: #bbb;
            }
            QTimeEdit::up-button { width: 0px; }
            QTimeEdit::down-button { width: 0px; }
        """)
        self.input_time.timeChanged.connect(lambda: setattr(self, '_time_touched', True))
        self.input_time.installEventFilter(self)
        for child in self.input_time.findChildren(QWidget):
            child.installEventFilter(self)
        self._cal_popup = CustomCalendarWidget()
        self._cal_popup.setWindowFlags(Qt.Popup)
        self._cal_popup.setStyleSheet("""
            QCalendarWidget {
                background-color: white;
            }
            QCalendarWidget QAbstractItemView {
                background-color: white;
                color: black;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                selection-background-color: #d4b800;
                selection-color: white;
            }
            QCalendarWidget QWidget {
                background-color: white;
            }
            QCalendarWidget QToolButton {
                background-color: white;
                color: black;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: white;
            }
            QCalendarWidget QSpinBox {
                background-color: white;
                color: black;
            }
        """)
        fmt_sat = QTextCharFormat()
        fmt_sat.setForeground(QColor('#0055cc'))
        self._cal_popup.setWeekdayTextFormat(Qt.Saturday, fmt_sat)
        fmt_sun = QTextCharFormat()
        fmt_sun.setForeground(QColor('#cc0000'))
        self._cal_popup.setWeekdayTextFormat(Qt.Sunday, fmt_sun)
        self._cal_popup.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self._cal_popup.clicked.connect(self._on_cal_date_selected)
        btn_history = QPushButton('지난 기록')
        btn_history.setFont(QFont('Malgun Gothic', 10))
        btn_history.setStyleSheet("""
            QPushButton {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #5a4000;
            }
            QPushButton:hover { background: rgba(212,184,0,0.25); }
        """)
        btn_history.clicked.connect(self._show_history)

        self.input_recurrence = QComboBox()
        self.input_recurrence.addItems(['반복 없음', '매주', '격주', '매월', '매년', '사용자 설정...'])
        for _i in range(self.input_recurrence.count()):
            self.input_recurrence.setItemData(_i, Qt.AlignCenter, Qt.TextAlignmentRole)
        self.input_recurrence.installEventFilter(self)
        self.input_recurrence.setFixedWidth(90)
        self.input_recurrence.setFont(QFont('Malgun Gothic', 10))
        self.input_recurrence.setStyleSheet(
            "QComboBox { border: 1px solid #d4b800; border-radius: 4px; padding: 0 4px; "
            "background: #fffde7; color: #333; }"
            "QComboBox:focus { border: 2px solid #4a90d9; background: rgba(255,255,255,0.85); }"
            "QComboBox::drop-down { border: none; width: 0px; }"
            "QComboBox QAbstractItemView { font-family: 'Malgun Gothic'; font-size: 10pt; }"
        )
        self._custom_recurrence_val = ''
        self.input_recurrence.currentIndexChanged.connect(self._on_recurrence_index_changed)

        _row_h = self.input_name.sizeHint().height()
        self.input_date.setFixedHeight(_row_h)
        self.input_time.setFixedHeight(_row_h)
        self.input_recurrence.setFixedHeight(_row_h)

        row2.addWidget(lbl_date)
        row2.addWidget(self.input_date)
        row2.addWidget(self.input_time)
        row2.addWidget(self.input_recurrence)
        row2.addStretch()
        row2.addWidget(btn_history)
        content_layout.addWidget(self.date_row_widget)

        self._sep_date = QFrame()
        self._sep_date.setFrameShape(QFrame.HLine)
        self._sep_date.setFixedHeight(1)
        self._sep_date.setStyleSheet('background: #D4AF37; border: none;')
        content_layout.addWidget(self._sep_date)

        # 업무 목록
        self.task_list_widget = QWidget()
        self.task_list_widget.setStyleSheet('background: transparent;')
        self.task_list_layout = QVBoxLayout(self.task_list_widget)
        self.task_list_layout.setContentsMargins(0, 0, 0, 0)
        self.task_list_layout.setSpacing(0)
        self.task_list_layout.addStretch()

        self.task_scroll = QScrollArea()
        self.task_scroll.setWidgetResizable(True)
        self.task_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.task_scroll.setWidget(self.task_list_widget)
        self.task_scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical { width: 6px; background: transparent; }
            QScrollBar::handle:vertical { background: #d4b800; border-radius: 3px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)
        content_layout.addWidget(self.task_scroll)

        # 메모 모드 자유 편집기 (기본 숨김)
        self.memo_editor = QPlainTextEdit()
        self.memo_editor.setPlaceholderText('여기에 바로 메모를 적으세요...')
        self.memo_editor.setFont(pr_font(12))
        self.memo_editor.setStyleSheet("""
            QPlainTextEdit {
                background: transparent;
                border: none;
                color: #333;
                padding: 4px 8px;
            }
        """)
        self.memo_editor.textChanged.connect(self._on_memo_text_changed)
        self.memo_editor.hide()
        content_layout.addWidget(self.memo_editor)
        _splitter_layout.addWidget(_top_panel, stretch=1)

        # ── 하단 패널: 공문 작성 버튼 ──────────────────────────────
        _bottom_panel = QWidget()
        _bottom_panel.setStyleSheet('background: transparent;')
        bottom_layout = QHBoxLayout(_bottom_panel)
        bottom_layout.setContentsMargins(8, 6, 8, 6)
        bottom_layout.setSpacing(6)

        _btn_style = """
            QPushButton {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 6px;
                padding: 4px 8px;
                color: #5a4000;
            }
            QPushButton:hover { background: rgba(212,184,0,0.25); }
        """

        btn_doc_write = QPushButton('공문 작성 (AI)')
        btn_doc_write.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        btn_doc_write.setStyleSheet(_btn_style)
        btn_doc_write.clicked.connect(self._open_document_editor)
        bottom_layout.addWidget(btn_doc_write)

        btn_doc_history = QPushButton('저장 기록')
        btn_doc_history.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        btn_doc_history.setStyleSheet(_btn_style)
        btn_doc_history.clicked.connect(self._show_official_doc_history)
        bottom_layout.addWidget(btn_doc_history)

        # 기존 capture 메서드 호환용 더미 위젯 (숨김)
        self.lbl_capture_result = QLabel('')
        self.lbl_capture_result.hide()
        self.lbl_capture_status = QLabel('')
        self.lbl_capture_status.hide()
        self.doc_list_widget = QWidget()
        self.doc_list_layout = QVBoxLayout(self.doc_list_widget)
        self.doc_list_layout.addStretch()
        self.doc_list_widget.hide()

        _sep_bottom = QFrame()
        _sep_bottom.setFrameShape(QFrame.HLine)
        _sep_bottom.setFixedHeight(1)
        _sep_bottom.setStyleSheet('background: #D4AF37; border: none;')
        _splitter_layout.addWidget(_sep_bottom)
        _splitter_layout.addWidget(_bottom_panel, stretch=0)

        _all_ids = [w['id'] for w in get_all_windows()]
        if _all_ids and self.window_id != min(_all_ids):
            _bottom_panel.hide()

        root.addWidget(self.content)
        self._apply_color(self._bg_color)
        self._renew_overdue_recurring_tasks()
        self._refresh_tasks()
        self._refresh_documents()
        self._schedule_midnight_refresh()
        self._setup_resize_handles()

        # 메모 저장 타이머 (디바운스 500ms)
        self._memo_save_timer = QTimer(self)
        self._memo_save_timer.setSingleShot(True)
        self._memo_save_timer.timeout.connect(self._flush_memo_save)

        # 저장된 메모 모드 및 텍스트 복원
        _saved_mode, _saved_text = get_window_memo(self.window_id)
        self.memo_editor.setPlainText(_saved_text)
        if _saved_mode:
            self._apply_memo_mode(True)

    def eventFilter(self, obj, event):
        if hasattr(self, '_cal_popup') and (obj is self.input_date or self.input_date.isAncestorOf(obj)) and event.type() == QEvent.MouseButtonPress:
            self._toggle_cal_popup()
            return True
        if event.type() == QEvent.KeyPress:
            if obj is self.input_name and event.key() == Qt.Key_Tab:
                if self.input_date.date() == QDate(1900, 1, 1):
                    self.input_date.setDate(QDate.currentDate())
                self.input_date.setFocus()
                return True
            if obj is self.input_date:
                if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                    self._add_task()
                    return True
                if event.key() == Qt.Key_Escape:
                    self.input_date.setDate(QDate(1900, 1, 1))
                    self.input_time.setEnabled(False)
                    self.input_time.setTime(QTime(9, 0))
                    self._date_touched = False
                    self._time_touched = False
                    self.input_name.setFocus()
                    return True
                if event.key() == Qt.Key_Tab:
                    if self.input_date.currentSection() == QDateTimeEdit.DaySection:
                        if self.input_time.isEnabled():
                            self.input_time.setFocus()
                        else:
                            self.input_name.setFocus()
                        return True
            if obj is self.input_time or (hasattr(self, 'input_time') and self.input_time.isAncestorOf(obj)):
                if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                    self._add_task()
                    return True
                if event.key() == Qt.Key_Tab:
                    if self.input_time.currentSection() == QDateTimeEdit.MinuteSection:
                        self.input_recurrence.setFocus()
                        return True
            if obj is self.input_recurrence:
                if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                    self._add_task()
                    return True
        return super().eventFilter(obj, event)

    def _on_date_changed(self, qdate):
        if self.input_recurrence.currentIndex() == 5:
            self._date_touched = False
            self.input_time.setEnabled(False)
            self.input_time.setTime(QTime(9, 0))
            self._time_touched = False
            self.input_date.setStyleSheet(self._field_style_date_empty)
            return
        if qdate != QDate(1900, 1, 1):
            self._date_touched = True
            self.input_time.setEnabled(True)
            self.input_date.setStyleSheet(self._field_style_date_filled)
        else:
            self._date_touched = False
            self.input_time.setEnabled(False)
            self.input_time.setTime(QTime(9, 0))
            self._time_touched = False
            self.input_date.setStyleSheet(self._field_style_date_empty)

    def _on_name_cleared(self, text):
        if not text:
            self.input_date.setDate(QDate(1900, 1, 1))
            self.input_time.setEnabled(False)
            self.input_time.setTime(QTime(9, 0))
            self._date_touched = False
            self._time_touched = False

    def _on_recurrence_index_changed(self, idx):
        if idx == 5:
            dlg = CustomRecurrenceDialog(self, current=self._custom_recurrence_val)
            if dlg.exec_() == QDialog.Accepted:
                self._custom_recurrence_val = dlg.get_value()
            else:
                self.input_recurrence.blockSignals(True)
                self.input_recurrence.setCurrentIndex(0)
                self.input_recurrence.blockSignals(False)
        self._update_deadline_enabled_state()

    def _update_deadline_enabled_state(self):
        is_custom = self.input_recurrence.currentIndex() == 5
        self.input_date.setEnabled(not is_custom)
        self.input_time.setEnabled((not is_custom) and self._date_touched)
        if is_custom:
            self.input_date.setDate(QDate(1900, 1, 1))
            self.input_time.setTime(QTime(9, 0))
            self._date_touched = False
            self._time_touched = False
            self.input_date.setStyleSheet(self._field_style_date_empty)

    def _toggle_memo_mode(self):
        self._apply_memo_mode(not self._memo_mode)
        self._flush_memo_save()

    def _apply_memo_mode(self, on: bool):
        self._memo_mode = on
        self._title_sep.setVisible(on)
        self._sep_date.setVisible(not on)
        self.task_input_row.setVisible(not on)
        self.date_row_widget.setVisible(not on)
        self.task_scroll.setVisible(not on)
        self.memo_editor.setVisible(on)
        btn = self.title_bar.btn_mode
        if on:
            self.memo_editor.setFocus()
            btn.setToolTip('업무 모드로 전환')
            btn.setGraphicsEffect(None)
        else:
            self.input_name.setFocus()
            btn.setToolTip('메모 모드로 전환')
            _gray = QGraphicsColorizeEffect()
            _gray.setColor(QColor('#888888'))
            _gray.setStrength(1.0)
            btn.setGraphicsEffect(_gray)

    def _on_memo_text_changed(self):
        if self._memo_save_timer is not None:
            self._memo_save_timer.start(500)

    def _flush_memo_save(self):
        text = self.memo_editor.toPlainText()
        set_window_memo(self.window_id, 1 if self._memo_mode else 0, text)

    def _add_task(self):
        name = self.input_name.text().strip()
        if not name:
            return
        if self._date_touched:
            date_str = self.input_date.date().toString('yyyy-MM-dd')
            if self._time_touched:
                deadline = f'{date_str} {self.input_time.time().toString("HH:mm")}'
            else:
                deadline = date_str
        else:
            deadline = ''
        _RECUR_MAP = {0: '', 1: 'weekly', 2: 'biweekly', 3: 'monthly', 4: 'yearly'}
        idx = self.input_recurrence.currentIndex()
        if idx == 5:
            recurrence = self._custom_recurrence_val if self._custom_recurrence_val.startswith('custom:') else ''
            if recurrence and not self._date_touched:
                # 날짜 미지정 시 오늘 기준 다음 해당일을 자동 계산
                yesterday = (date.today() - timedelta(days=1)).isoformat()
                deadline = _next_recurrence_deadline(yesterday, recurrence) or ''
        else:
            recurrence = _RECUR_MAP.get(idx, '')
        add_task(self.window_id, name, deadline, recurrence=recurrence)
        self.input_name.clear()
        self.input_date.setDate(QDate(1900, 1, 1))
        self.input_time.setEnabled(False)
        self.input_time.setTime(QTime(9, 0))
        self.input_recurrence.setCurrentIndex(0)
        self._custom_recurrence_val = ''
        self._date_touched = False
        self._time_touched = False
        self.input_name.setFocus()
        self._refresh_tasks()

    def _delete_task(self, task):
        add_task_history(self.window_id, task['name'], task['deadline'],
                         strikethrough=task.get('strikethrough', 0),
                         priority=task.get('priority', 0),
                         recurrence=task.get('recurrence', ''))
        delete_task(task['id'])
        self._refresh_tasks()

    def _open_global_search(self):
        dlg = GlobalSearchDialog(self._open_windows, parent=self)
        dlg.exec_()

    def _highlight_task(self, task_id):
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            if item and item.widget() and hasattr(item.widget(), 'task'):
                row = item.widget()
                if row.task.get('id') == task_id:
                    row.setStyleSheet('background: rgba(212,184,0,0.35); border-radius: 4px;')
                    QTimer.singleShot(1500, lambda r=row: r.setStyleSheet(''))
                    break

    def _schedule_midnight_refresh(self):
        if hasattr(self, '_midnight_timer'):
            self._midnight_timer.stop()
            self._midnight_timer.deleteLater()
        now = QDateTime.currentDateTime()
        midnight = QDateTime(now.date().addDays(1), QTime(0, 0, 0))
        ms_until_midnight = now.msecsTo(midnight)
        self._midnight_timer = QTimer(self)
        self._midnight_timer.setSingleShot(True)
        self._midnight_timer.timeout.connect(self._on_midnight)
        self._midnight_timer.start(ms_until_midnight)

    def _renew_overdue_recurring_tasks(self):
        today = date.today().isoformat()
        for task in get_tasks(self.window_id):
            if not task.get('recurrence') or not task.get('deadline'):
                continue
            dl = task['deadline'][:10]
            if dl >= today:
                continue
            next_dl = task['deadline']
            for _ in range(366):
                next_dl = _next_recurrence_deadline(next_dl, task['recurrence'])
                if not next_dl:
                    break
                if next_dl[:10] >= today:
                    update_task(task['id'], task['name'], next_dl,
                                strikethrough=0,
                                priority=task.get('priority', 0),
                                recurrence=task['recurrence'])
                    break

    def _on_midnight(self):
        self._renew_overdue_recurring_tasks()
        self._refresh_tasks()
        self._schedule_midnight_refresh()  # 다음 자정 예약

    def _refresh_tasks(self):
        while self.task_list_layout.count() > 1:
            item = self.task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for task in get_tasks(self.window_id):
            row = TaskRow(task, self._delete_task, self._refresh_tasks, scale=self._scale)
            self.task_list_layout.insertWidget(self.task_list_layout.count() - 1, row)

    def _refresh_documents(self):
        while self.doc_list_layout.count() > 1:
            item = self.doc_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for doc in get_documents(self.window_id):
            row = DocumentRow(doc, self._delete_document, on_paste=self._make_paste_cb(doc['id']))
            self.doc_list_layout.insertWidget(self.doc_list_layout.count() - 1, row)
        blank = DocumentRow({'id': None, 'title': '', 'doc_number': ''}, self._delete_document, on_paste=self._make_paste_cb(None))
        self.doc_list_layout.insertWidget(self.doc_list_layout.count() - 1, blank)

    def _make_paste_cb(self, doc_id):
        def cb(field, title_edit, num_edit):
            text = self.lbl_capture_result.text()
            if not text:
                return
            if doc_id is None:
                if field == 'title':
                    add_document(self.window_id, text, '')
                else:
                    add_document(self.window_id, '', _normalize_doc_number(text))
            else:
                if field == 'title':
                    update_document(doc_id, text, num_edit.text())
                else:
                    update_document(doc_id, title_edit.text(), _normalize_doc_number(text))
            self._refresh_documents()
        return cb

    def _delete_document(self, doc_id):
        delete_document(doc_id)
        self._refresh_documents()

    def _open_document_editor(self):
        try:
            from document_editor import DocumentEditorWindow
            self._doc_editor = DocumentEditorWindow(parent=None)
            self._doc_editor.show()
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(self, '오류', f'공문 작성 창 열기 실패:\n{e}')

    def _show_official_doc_history(self):
        from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                      QListWidget, QListWidgetItem, QPushButton,
                                      QLabel, QTextEdit, QSplitter, QMessageBox)
        docs = get_official_documents()

        dlg = QDialog(self)
        dlg.setWindowTitle('저장 기록')
        dlg.setMinimumSize(560, 420)
        dlg.setStyleSheet("font-family: 'Malgun Gothic'; font-size: 10pt;")
        _qs = QSettings('SSNnote', 'SSNnote')
        geo = _qs.value('history_dlg/geometry')
        if geo:
            dlg.restoreGeometry(geo)
        dlg.finished.connect(lambda: QSettings('SSNnote', 'SSNnote').setValue('history_dlg/geometry', dlg.saveGeometry()))

        vlay = QVBoxLayout(dlg)
        vlay.setContentsMargins(10, 10, 10, 10)

        if not docs:
            vlay.addWidget(QLabel('저장된 공문이 없습니다.'))
            btn_close = QPushButton('닫기')
            btn_close.clicked.connect(dlg.accept)
            vlay.addWidget(btn_close)
            dlg.exec_()
            return

        splitter = QSplitter(Qt.Horizontal)

        list_widget = QListWidget()
        list_widget.setMaximumWidth(200)
        for doc in docs:
            item = QListWidgetItem(doc['title'] or '(제목 없음)')
            item.setData(Qt.UserRole, doc)
            list_widget.addItem(item)
        splitter.addWidget(list_widget)

        right = QWidget()
        rlay = QVBoxLayout(right)
        rlay.setContentsMargins(6, 0, 0, 0)
        lbl_title = QLabel()
        lbl_title.setStyleSheet('font-weight: bold; font-size: 11pt;')
        lbl_date = QLabel()
        lbl_date.setStyleSheet('color: #888; font-size: 9pt;')
        text_view = QTextEdit()
        text_view.setReadOnly(True)
        rlay.addWidget(lbl_title)
        rlay.addWidget(lbl_date)
        rlay.addWidget(text_view)
        splitter.addWidget(right)
        splitter.setStretchFactor(1, 1)
        vlay.addWidget(splitter)

        def on_select():
            item = list_widget.currentItem()
            if not item:
                return
            doc = item.data(Qt.UserRole)
            lbl_title.setText(doc['title'] or '(제목 없음)')
            lbl_date.setText(f"저장일: {doc.get('created_at', '')}")
            text_view.setPlainText(doc.get('content', ''))

        list_widget.currentItemChanged.connect(lambda *_: on_select())
        if list_widget.count() > 0:
            list_widget.setCurrentRow(0)

        btn_row = QHBoxLayout()
        btn_del = QPushButton('선택 삭제')
        btn_del.setStyleSheet("""
            QPushButton { background: #fff0f0; border: 1px solid #e08080;
                          border-radius: 4px; padding: 4px 10px; color: #a00; }
            QPushButton:hover { background: #ffd8d8; }
        """)
        def on_delete():
            item = list_widget.currentItem()
            if not item:
                return
            doc = item.data(Qt.UserRole)
            reply = QMessageBox.question(dlg, '삭제 확인',
                f'「{doc["title"] or "(제목 없음)"}」을(를) 삭제할까요?',
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                delete_official_document(doc['id'])
                list_widget.takeItem(list_widget.row(item))
                lbl_title.clear(); lbl_date.clear(); text_view.clear()
        btn_del.clicked.connect(on_delete)
        btn_close = QPushButton('닫기')
        btn_close.setStyleSheet("""
            QPushButton { background: rgba(255,255,255,0.8); border: 1px solid #aaa;
                          border-radius: 4px; padding: 4px 10px; }
            QPushButton:hover { background: #eee; }
        """)
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        vlay.addLayout(btn_row)

        dlg.exec_()

    def _start_capture(self):
        # 모든 메모 창 숨기고 화면 캡처 후 오버레이 표시
        for win in self._open_windows:
            win.hide()
        QTimer.singleShot(150, self._show_capture_overlay)

    def _show_capture_overlay(self):
        screenshot = grab_fullscreen()
        self._overlay = ScreenCaptureOverlay(screenshot)
        self._overlay.region_captured.connect(self._on_capture_complete)
        self._overlay.cancelled.connect(self._restore_windows)
        self._overlay.show()
        self._overlay.activateWindow()
        self._overlay.setFocus()

    def _restore_windows(self):
        for win in self._open_windows:
            win.show()

    def _on_capture_complete(self, pixmap):
        self._overlay.region_captured.disconnect()  # 중복 호출 방지
        self._restore_windows()

        def _after_ocr(text):
            if text:
                self.lbl_capture_result.setText(text)
                if not self._capture_hint_shown:
                    self._capture_hint_shown = True
                    self.lbl_capture_status.setText('캡쳐 완료! 아래 빈 칸을 클릭하세요.')
                    self.lbl_capture_status.show()
                    QTimer.singleShot(3000, self.lbl_capture_status.hide)
                hwnd = int(self.winId())
                ctypes.windll.user32.ShowWindow(hwnd, 9)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                self.raise_()
                self.activateWindow()
            else:
                self.lbl_capture_result.setText('')
                QMessageBox.information(self, '알림', 'OCR 텍스트를 인식하지 못했습니다.')

        def _on_error(msg):
            QMessageBox.warning(self, 'OCR 오류', msg)

        self._ocr_worker = run_ocr(pixmap, _after_ocr, _on_error)

    @staticmethod
    def _make_check_icon(checked, size=20):
        from PyQt5.QtGui import QPixmap, QPainter, QColor, QPen, QIcon
        px = QPixmap(size, size)
        px.fill(QColor(0, 0, 0, 0))
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        if checked:
            p.setPen(Qt.NoPen)
            p.setBrush(QColor('#2ecc71'))
            p.drawRoundedRect(1, 1, size - 2, size - 2, 4, 4)
            p.setPen(QPen(QColor('#ffffff'), 2.5, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            bx, by = int(size * 0.40), int(size * 0.62)
            p.drawLine(int(size * 0.18), int(size * 0.50), bx, by)
            p.drawLine(bx, by, int(size * 0.82), int(size * 0.25))
        else:
            p.setPen(QPen(QColor('#aaaaaa'), 1.5))
            p.setBrush(QColor(0, 0, 0, 0))
            p.drawRoundedRect(2, 2, size - 4, size - 4, 3, 3)
        p.end()
        return QIcon(px)

    def _make_check_widget_action(self, parent_menu, text, checked, on_trigger):
        from PyQt5.QtWidgets import QWidgetAction, QToolButton, QSizePolicy
        wa = QWidgetAction(parent_menu)
        btn = QToolButton()
        btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        btn.setIcon(self._make_check_icon(checked))
        btn.setIconSize(QSize(20, 20))
        btn.setText(text)
        btn.setFont(QFont('Malgun Gothic', 11))
        btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        btn.setStyleSheet("""
            QToolButton {
                background: transparent;
                border: none;
                padding: 7px 16px 7px 4px;
                color: #333;
                text-align: left;
            }
            QToolButton:hover {
                background: rgba(212,184,0,0.3);
                border-radius: 4px;
            }
        """)
        def _clicked():
            on_trigger()
            QTimer.singleShot(0, parent_menu.close)
        btn.pressed.connect(_clicked)
        wa.setDefaultWidget(btn)
        return wa

    def _make_menu_style(self):
        return """
            QMenu {
                background: #fffacd;
                border: 1px solid #d4b800;
                border-radius: 6px;
                padding: 4px 0;
                font-family: 'Malgun Gothic';
                font-size: 11pt;
                color: #333;
                min-width: 176px;
            }
            QMenu::item { padding: 7px 16px 7px 4px; }
            QMenu::item:selected { background: rgba(212,184,0,0.3); }
        """

    def show_menu(self, btn):
        menu = QMenu(self)
        menu.setStyleSheet(self._make_menu_style())

        act_new    = QAction('🆕 새 메모장', self)
        act_search = QAction('🔍 통합 검색', self)
        autostart_on = autostart_is_enabled()
        act_auto = QAction(('✅ ' if autostart_on else '☐ ') + '시작 시 자동실행', self)
        act_auto.triggered.connect(lambda: autostart_set(not autostart_on))
        act_help   = QAction('💡 단축키', self)
        act_update = QAction('⬇️ 업데이트 확인', self)
        act_delete = QAction('🗑️ 메모장 삭제', self)

        # 배경색 서브메뉴
        color_menu = QMenu('🎨 배경색 바꾸기', self)
        color_menu.setStyleSheet(self._make_menu_style())

        color_widget = QWidget()
        color_widget.setStyleSheet('background: transparent;')
        grid = QGridLayout(color_widget)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setSpacing(6)
        for i, color in enumerate(PALETTE):
            border = '2px solid #333' if color == '#FFFFFF' else '1px solid rgba(0,0,0,0.15)'
            cb = QPushButton()
            cb.setFixedSize(30, 30)
            cb.setStyleSheet(f"""
                QPushButton {{
                    background: {color};
                    border-radius: 15px;
                    border: {border};
                }}
                QPushButton:hover {{ border: 2px solid rgba(0,0,0,0.5); }}
            """)
            cb.clicked.connect(lambda _, c=color: (self._apply_color(c), self.save_state(), menu.close()))
            grid.addWidget(cb, i // 7, i % 7)

        wa = QWidgetAction(color_menu)
        wa.setDefaultWidget(color_widget)
        color_menu.addAction(wa)

        # 마감 알림 메뉴
        alarm_menu = QMenu('⏰ 마감 알림', self)
        alarm_menu.setStyleSheet(self._make_menu_style())
        alarm_menu.setToolTipsVisible(True)

        # ── D-day 5일 이내 알림 서브메뉴 ──
        dday_menu = QMenu('D-day 5일 이내 알림', self)
        dday_menu.setStyleSheet(self._make_menu_style())

        current_min = self._get_alarm_interval() if self._get_alarm_interval else 180
        alarm_options = [
            ('잡도리 모드 (30분마다)', 30),
            ('긴장 모드 (1시간마다)', 60),
            ('기본 설정 (3시간마다)', 180),
        ]

        off_check = '✅ ' if current_min == 0 else '　 '
        act_off = QAction(off_check + '사용 안함', self)
        act_off.triggered.connect(lambda: self._set_alarm_interval(0, menu))
        dday_menu.addAction(act_off)

        dday_menu.addSeparator()
        for label, minutes in alarm_options:
            check = '✅ ' if current_min == minutes else '　 '
            act = QAction(check + label, self)
            act.triggered.connect(lambda _, m=minutes: self._set_alarm_interval(m, menu))
            dday_menu.addAction(act)

        dday_menu.addSeparator()
        custom_label = '✅ 사용자 설정' if current_min not in [m for _, m in alarm_options] and current_min != 0 else '　 사용자 설정'
        act_custom = QAction(custom_label, self)
        act_custom.triggered.connect(lambda: self._set_alarm_interval_custom(menu))
        dday_menu.addAction(act_custom)

        alarm_menu.addMenu(dday_menu)

        # ── 당일 3시간 이내 알림 토글 ──
        timed_enabled = self._get_timed_alarm_enabled() if self._get_timed_alarm_enabled else True
        act_timed = QAction(('✅ ' if timed_enabled else '☐ ') + '당일 3시간 알림', self)
        act_timed.setToolTip('3시간, 2시간, 1시간, 30분 전에 팝업 알림이 뜹니다.')
        act_timed.triggered.connect(lambda: self._set_timed_alarm_enabled(not timed_enabled, menu))
        alarm_menu.addAction(act_timed)

        # 단축키 사용 체크 액션
        sc_enabled = self._get_shortcut_enabled() if self._get_shortcut_enabled else True
        act_sc = QAction(('✅ ' if sc_enabled else '☐ ') + '단축키 사용', self)
        act_sc.triggered.connect(lambda: self._toggle_shortcut(not sc_enabled))

        # 글자 크기 서브메뉴
        size_menu = QMenu('🔠 글자 크기', self)
        size_menu.setStyleSheet(self._make_menu_style().replace('min-width: 176px', 'min-width: 100px'))
        for label, val in [('기본', 1.0), ('크게', 1.1), ('더 크게', 1.2)]:
            check = '✅ ' if abs(self._scale - val) < 0.01 else '　 '
            act_sz = QAction(check + label, self)
            act_sz.triggered.connect(lambda _, v=val: self._set_scale(v, menu))
            size_menu.addAction(act_sz)

        menu.addAction(act_new)
        menu.addAction(act_search)
        menu.addMenu(color_menu)
        menu.addMenu(size_menu)
        menu.addMenu(alarm_menu)
        menu.addAction(act_auto)
        menu.addAction(act_sc)
        menu.addAction(act_help)
        menu.addAction(act_update)
        _all_ids = [w['id'] for w in get_all_windows()]
        if not _all_ids or self.window_id != min(_all_ids):
            menu.addSeparator()
            menu.addAction(act_delete)

        act_search.triggered.connect(self._open_global_search)
        act_new.triggered.connect(lambda: self.on_new(offset_from=self, on_toggle_hotkey=self._on_toggle_hotkey,
                                                       on_shortcut_change=self._on_shortcut_change,
                                                       get_shortcut_enabled=self._get_shortcut_enabled))
        act_help.triggered.connect(self.show_help)
        act_update.triggered.connect(lambda: updater.check_for_update_manual(self))
        act_delete.triggered.connect(self.delete_memo)

        pos = btn.mapToGlobal(btn.rect().bottomLeft())
        menu.exec_(pos)

    def _apply_color(self, hex_color):
        self._bg_color = hex_color
        title_color = QColor(hex_color).darker(115).name()
        self.centralWidget().setStyleSheet(f"""
            QWidget#central {{
                background: {hex_color};
                border-radius: 0;
                border: 1px solid rgba(0,0,0,0.4);
            }}
        """)
        self.title_bar.setStyleSheet(f"""
            QWidget {{
                background: {title_color};
                border-radius: 0;
            }}
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
                padding: 0;
            }}
            QPushButton:hover {{ background: rgba(0,0,0,0.12); }}
            QPushButton#btn_mode {{
                background: transparent;
                border: 1px solid transparent;
                border-radius: 6px;
                padding: 4px 8px;
                color: #5a4000;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                font-weight: bold;
            }}
            QPushButton#btn_mode:hover {{ background: rgba(0,0,0,0.12); border-color: rgba(0,0,0,0.12); }}
            QPushButton#btn_mode:pressed {{ background: rgba(0,0,0,0.20); border-color: rgba(0,0,0,0.20); }}
        """)

    def _set_alarm_interval(self, minutes, menu=None):
        if self._on_alarm_interval_change:
            self._on_alarm_interval_change(minutes)
        if menu:
            menu.close()

    def _set_alarm_interval_custom(self, menu=None):
        from PyQt5.QtWidgets import QInputDialog
        current = self._get_alarm_interval() if self._get_alarm_interval else 3
        current_h = max(1, round(current / 60))
        hours, ok = QInputDialog.getInt(
            self, '사용자 설정', '알림 주기를 입력하세요 (시간 단위):',
            value=current_h, min=1, max=24
        )
        if ok:
            self._set_alarm_interval(hours * 60, menu)

    def _set_timed_alarm_enabled(self, enabled, menu=None):
        if self._on_timed_alarm_change:
            self._on_timed_alarm_change(enabled)
        if menu:
            menu.close()

    def _set_local_shortcuts_enabled(self, enabled):
        for sc in self._shortcuts:
            sc.setEnabled(enabled)

    def _toggle_shortcut(self, enabled, menu=None):
        if self._on_shortcut_change:
            self._on_shortcut_change(enabled)
        if menu:
            menu.close()

    def _show_history(self):
        from collections import defaultdict
        records = get_task_history(self.window_id)
        current_ym = date.today().strftime('%Y-%m')

        dlg = QDialog(self)
        dlg.setWindowTitle('지난 기록')
        dlg.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint)
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet("font-family: 'Malgun Gothic'; font-size: 11pt; background: #fffef0;")

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        if not records:
            lbl = QLabel('기록이 없습니다.')
            lbl.setStyleSheet('color: #999;')
            layout.addWidget(lbl)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setStyleSheet('QScrollArea { border: none; }')
            scroll.setMaximumHeight(440)

            inner = QWidget()
            inner.setStyleSheet('background: transparent;')
            vbox = QVBoxLayout(inner)
            vbox.setContentsMargins(0, 0, 0, 0)
            vbox.setSpacing(4)

            def refresh_folder(btn_toggle, container, year, month):
                layout = container.layout()
                visible = sum(
                    1 for i in range(layout.count())
                    if layout.itemAt(i).widget() and layout.itemAt(i).widget().isVisible()
                )
                if visible == 0:
                    btn_toggle.hide()
                    container.hide()
                else:
                    arrow = '▶' if container.isHidden() else '▼'
                    btn_toggle.setText(f'{arrow}  {year}년 {int(month)}월  ({visible}건)')

            def restore_task(r, row_widget, btn_toggle, container, year, month):
                add_task(self.window_id, r['name'], r['deadline'],
                         strikethrough=r.get('strikethrough', 0),
                         priority=r.get('priority', 0),
                         recurrence=r.get('recurrence', ''))
                delete_task_history(r['id'])
                self._refresh_tasks()
                row_widget.hide()
                refresh_folder(btn_toggle, container, year, month)

            def delete_history(r, row_widget, btn_toggle, container, year, month):
                delete_task_history(r['id'])
                row_widget.hide()
                refresh_folder(btn_toggle, container, year, month)

            def make_row(r, btn_toggle, container, year, month):
                name     = r['name']
                deadline = r['deadline']
                cleared  = r['cleared_at'][:10]

                overdue = False
                if deadline:
                    dday = calc_dday(deadline)
                    overdue = dday.startswith('D+')
                    line = f"{name}  ({deadline} / {dday})"
                else:
                    line = name

                row_widget = QWidget()
                row_widget.setStyleSheet('background: rgba(255,255,255,0.6); border-radius: 4px;')
                rh = QHBoxLayout(row_widget)
                rh.setContentsMargins(20, 4, 8, 4)
                rh.setSpacing(6)

                lbl_task = QLineEdit(line)
                lbl_task.setReadOnly(True)
                lbl_task.setFrame(False)
                lbl_task.setStyleSheet("""
                    QLineEdit {
                        color: #333;
                        background: transparent;
                        border: none;
                        padding: 0;
                        font-family: 'Malgun Gothic';
                        font-size: 11pt;
                    }
                """)
                lbl_task.setCursorPosition(0)
                lbl_cleared = QLabel(f'삭제: {cleared}')
                lbl_cleared.setStyleSheet('color: #999; font-size: 9pt; background: transparent;')
                lbl_cleared.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

                rh.addWidget(lbl_task, 1)
                rh.addSpacing(11)
                rh.addWidget(lbl_cleared)

                if not overdue:
                    btn_restore = QPushButton('복원')
                    btn_restore.setFixedHeight(22)
                    btn_restore.setStyleSheet("""
                        QPushButton {
                            background: #d4f0d4; border: 1px solid #7cc87c;
                            border-radius: 4px; color: #2a6e2a;
                            font-size: 9pt; padding: 0 8px;
                        }
                        QPushButton:hover { background: #b6e8b6; }
                    """)
                    btn_restore.clicked.connect(lambda _, rec=r, rw=row_widget: restore_task(rec, rw, btn_toggle, container, year, month))
                    rh.addWidget(btn_restore)

                btn_del = QPushButton('삭제')
                btn_del.setFixedHeight(22)
                btn_del.setStyleSheet("""
                    QPushButton {
                        background: #f0d4d4; border: 1px solid #c87c7c;
                        border-radius: 4px; color: #6e2a2a;
                        font-size: 9pt; padding: 0 8px;
                    }
                    QPushButton:hover { background: #e8b6b6; }
                """)
                btn_del.clicked.connect(lambda _, rec=r, rw=row_widget: delete_history(rec, rw, btn_toggle, container, year, month))
                rh.addWidget(btn_del)

                return row_widget

            # 이전 달 기록: 마감일이 이번 달보다 이전인 것 (마감일 기준 년-월로 정리)
            past_groups = defaultdict(list)
            for r in records:
                if r['deadline'] and r['deadline'][:7] < current_ym:
                    past_groups[r['deadline'][:7]].append(r)

            for ym in sorted(past_groups.keys()):
                year, month = ym.split('-')
                grp = past_groups[ym]
                header_text_collapsed = f'▶  {year}년 {int(month)}월  ({len(grp)}건)'
                header_text_expanded  = f'▼  {year}년 {int(month)}월  ({len(grp)}건)'

                btn_toggle = QPushButton(header_text_collapsed)
                btn_toggle.setStyleSheet("""
                    QPushButton {
                        background: rgba(200,190,150,0.25);
                        border: 1px solid #c8b87c;
                        border-radius: 4px;
                        color: #5a4a00;
                        font-size: 10pt;
                        font-weight: bold;
                        text-align: left;
                        padding: 5px 10px;
                    }
                    QPushButton:hover { background: rgba(200,190,150,0.5); }
                """)

                month_container = QWidget()
                month_container.setObjectName('monthContainer')
                month_container.setStyleSheet("""
                    QWidget#monthContainer {
                        background: transparent;
                        border: 1px solid #c8b87c;
                        border-radius: 4px;
                    }
                """)
                mc_vbox = QVBoxLayout(month_container)
                mc_vbox.setContentsMargins(12, 4, 8, 4)
                mc_vbox.setSpacing(4)
                month_container.hide()

                for r in grp:
                    mc_vbox.addWidget(make_row(r, btn_toggle, month_container, year, month))

                def toggle_month(_, container=month_container, btn=btn_toggle,
                                 txt_exp=header_text_expanded, txt_col=header_text_collapsed):
                    if container.isHidden():
                        container.show()
                        btn.setText(txt_exp)
                    else:
                        container.hide()
                        btn.setText(txt_col)

                btn_toggle.clicked.connect(toggle_month)
                vbox.addWidget(btn_toggle)
                vbox.addWidget(month_container)

            # 이번 달 기록: 마감일이 없거나 이번 달 이상인 것 (기본 펼쳐진 헤더)
            current_records = [
                r for r in records
                if not r['deadline'] or r['deadline'][:7] >= current_ym
            ]
            if current_records:
                cur_year, cur_month = current_ym.split('-')
                cur_txt_exp = f'▼  {cur_year}년 {int(cur_month)}월  ({len(current_records)}건)'
                cur_txt_col = f'▶  {cur_year}년 {int(cur_month)}월  ({len(current_records)}건)'

                btn_cur = QPushButton(cur_txt_exp)
                btn_cur.setStyleSheet("""
                    QPushButton {
                        background: rgba(200,190,150,0.25);
                        border: 1px solid #c8b87c;
                        border-radius: 4px;
                        color: #5a4a00;
                        font-size: 10pt;
                        font-weight: bold;
                        text-align: left;
                        padding: 5px 10px;
                    }
                    QPushButton:hover { background: rgba(200,190,150,0.5); }
                """)

                cur_container = QWidget()
                cur_container.setObjectName('monthContainer')
                cur_container.setStyleSheet("""
                    QWidget#monthContainer {
                        background: transparent;
                        border: 1px solid #c8b87c;
                        border-radius: 4px;
                    }
                """)
                cur_vbox = QVBoxLayout(cur_container)
                cur_vbox.setContentsMargins(12, 4, 8, 4)
                cur_vbox.setSpacing(4)

                for r in current_records:
                    cur_vbox.addWidget(make_row(r, btn_cur, cur_container, cur_year, cur_month))

                def toggle_current(_, container=cur_container, btn=btn_cur,
                                   txt_exp=cur_txt_exp, txt_col=cur_txt_col):
                    if container.isHidden():
                        container.show()
                        btn.setText(txt_exp)
                    else:
                        container.hide()
                        btn.setText(txt_col)

                btn_cur.clicked.connect(toggle_current)
                vbox.addWidget(btn_cur)
                vbox.addWidget(cur_container)

            vbox.addStretch()
            scroll.setWidget(inner)
            layout.addWidget(scroll)

        dlg.exec_()

    def _toggle_cal_popup(self):
        if self._cal_popup.isVisible():
            self._cal_popup.hide()
        else:
            cur = self.input_date.date()
            show_date = cur if cur != QDate(1900, 1, 1) else QDate.currentDate()
            self._cal_popup.setSelectedDate(show_date)
            pos = self.input_date.mapToGlobal(QPoint(0, self.input_date.height()))
            self._cal_popup.move(pos)
            self._cal_popup.show()

    def _on_cal_date_selected(self, date):
        self.input_date.setDate(date)
        self._cal_popup.hide()


    def show_help(self):
        dlg = QDialog(self)
        dlg.setWindowTitle('단축키')
        dlg.setMinimumWidth(370)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(12, 12, 12, 12)
        lbl = QLabel(
            '<p style="line-height: 195%; font-family: Malgun Gothic; font-size: 11pt;">'
            '① Tab키, Enter키만 잘 쓰면 편하게 쓰실 수 있습니다.<br>'
            '② 단축키: <b>Ctrl + Shift</b> 를 누른 후<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>X</b> : 화면 캡처<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>S</b> : 포커싱<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>R</b> : 롤업<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>F</b> : 항상 위<br>'
            '③ 캡처는 공문을 전체화면으로 키워야 정확합니다.'
            '</p>'
        )
        lbl.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        layout.addWidget(lbl)
        btn = QPushButton('확인', dlg)
        btn.clicked.connect(dlg.accept)
        layout.addWidget(btn, alignment=Qt.AlignRight)
        dlg.exec_()

    def delete_memo(self):
        confirm = QMessageBox(self)
        confirm.setWindowTitle('메모장 삭제')
        confirm.setText('이 메모장을 삭제할까요?\n업무 목록도 함께 삭제됩니다.')
        confirm.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        confirm.setDefaultButton(QMessageBox.No)
        confirm.setStyleSheet("font-family: 'Malgun Gothic'; font-size: 11pt;")
        if confirm.exec_() == QMessageBox.Yes:
            delete_window(self.window_id)
            self._deleted = True
            self.close()

    def toggle_shade(self):
        if self.collapsed:
            target_height = self.expanded_height   # resizeEvent가 덮어쓰기 전에 저장
            self.collapsed = False
            self.setMinimumHeight(100)
            self.setMaximumHeight(16777215)
            self.resize(self.width(), target_height)
            self.content.show()
        else:
            self.expanded_height = self.height()
            self.collapsed = True           # resizeEvent보다 먼저 설정
            self.content.hide()
            self.setFixedHeight(TITLE_BAR_HEIGHT)
        if hasattr(self, '_handles'):
            self._reposition_handles()
        self.save_state()

    def toggle_always_on_top(self):
        self.pin_active = not self.pin_active
        flags = Qt.FramelessWindowHint | Qt.Tool
        if self.pin_active:
            flags |= Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.show()
        self.raise_()
        self.activateWindow()
        btn = self.title_bar.btn_pin
        if self.pin_active:
            btn.setText('📍')
            btn.setToolTip('항상 위 켜짐')
            btn.setStyleSheet('background: rgba(0,0,0,0.18); border-radius: 4px;')
            btn.setGraphicsEffect(None)
        else:
            btn.setText('📌')
            btn.setToolTip('항상 위 꺼짐')
            btn.setStyleSheet('')
            _gray = QGraphicsColorizeEffect()
            _gray.setColor(QColor('#888888'))
            _gray.setStrength(1.0)
            btn.setGraphicsEffect(_gray)

    def _snap_full(self, pos):
        """스냅 위치와 정렬 엣지 정보를 함께 반환."""
        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        rect = screen.availableGeometry()
        x, y = float(pos.x()), float(pos.y())
        w, h = self.width(), self.height()

        def snap_axis(val, lo, hi, size):
            dist_lo = val - lo
            dist_hi = hi - (val + size)
            if 0 <= dist_lo < SNAP_ZONE:
                if dist_lo < SNAP_THRESHOLD:
                    return float(lo)
                t = 1.0 - dist_lo / SNAP_ZONE
                return val - dist_lo * t * t
            if 0 <= dist_hi < SNAP_ZONE:
                if dist_hi < SNAP_THRESHOLD:
                    return float(hi - size)
                t = 1.0 - dist_hi / SNAP_ZONE
                return val + dist_hi * t * t
            return val

        nx = snap_axis(x, rect.left(), rect.right(), w)
        ny = snap_axis(y, rect.top(),  rect.bottom(), h)
        nx, ny, snap_edges, adj_snapped, x_snap_partner = self._snap_to_windows(nx, ny, w, h)
        # 타이틀바가 화면 밖으로 나가지 않도록 상단 클램핑
        ny = max(float(rect.top()), ny)
        return QPoint(int(nx), int(ny)), snap_edges, adj_snapped, x_snap_partner

    def _snap_to_screen(self, pos):
        result, _, _, _ = self._snap_full(pos)
        return result

    def _snap_to_windows(self, nx, ny, w, h):
        """바깥쪽 인접 스냅 + 상단 정렬 스냅 — AABB 절대 거리, 방향 무관."""
        my_r = nx + w
        my_b = ny + h

        min_dx        = float('inf')
        min_dy        = float('inf')
        min_align_dy  = float('inf')
        result_nx     = nx
        result_ny     = ny
        best_x_guide  = None
        best_y_guide  = None
        align_y_guide  = None
        align_y_result = None
        align_partner  = None
        x_snap_partner = None

        for other in self._open_windows:
            if other is self:
                continue
            r   = other.geometry()
            ox  = float(r.x());   oy  = float(r.y())
            ow  = float(r.width()); oh = float(r.height())
            o_r = ox + ow;  o_b = oy + oh

            # X 스냅 — Y 범위 겹침 필요 (바깥쪽 모서리만)
            if ny < o_b and my_b > oy:
                # 내 오른쪽 → 상대 왼쪽
                d = ox - my_r
                if abs(d) < SNAP_THRESHOLD and abs(d) < abs(min_dx):
                    min_dx = d;  result_nx = ox - w;  best_x_guide = int(ox); x_snap_partner = other
                # 내 왼쪽 → 상대 오른쪽
                d = nx - o_r
                if abs(d) < SNAP_THRESHOLD and abs(d) < abs(min_dx):
                    min_dx = d;  result_nx = o_r;     best_x_guide = int(o_r); x_snap_partner = other

            # Y 스냅 — X 범위 겹침 필요 (바깥쪽 모서리만)
            if nx < o_r and my_r > ox:
                # 내 아래 → 상대 위
                d = oy - my_b
                if abs(d) < SNAP_THRESHOLD and abs(d) < abs(min_dy):
                    min_dy = d;  result_ny = oy - h;  best_y_guide = int(oy)
                # 내 위 → 상대 아래
                d = ny - o_b
                if abs(d) < SNAP_THRESHOLD and abs(d) < abs(min_dy):
                    min_dy = d;  result_ny = o_b;     best_y_guide = int(o_b)

            # 상단/하단 정렬 스냅 — 가로로 인접한 창끼리 상단/하단 라인 맞추기
            x_adjacent = (abs(nx - o_r) < SNAP_ZONE or abs(my_r - ox) < SNAP_ZONE)
            if x_adjacent:
                d_top = abs(ny - oy)
                if d_top < SNAP_THRESHOLD and d_top < min_align_dy:
                    min_align_dy   = d_top
                    align_y_guide  = int(oy)
                    align_y_result = float(oy)
                    align_partner  = other
                d_bot = abs((ny + h) - (oy + oh))
                if d_bot < SNAP_THRESHOLD and d_bot < min_align_dy:
                    min_align_dy   = d_bot
                    align_y_guide  = int(oy + oh)
                    align_y_result = float(oy + oh - h)
                    align_partner  = other

        # 바깥쪽 Y 스냅이 없을 때만 정렬 스냅 적용
        if align_y_guide is not None and best_y_guide is None:
            result_ny = align_y_result
            best_y_guide = align_y_guide

        snap_edges  = []
        adj_snapped = False
        if best_x_guide is not None:
            snap_edges.append(('x', best_x_guide));  adj_snapped = True
        if best_y_guide is not None:
            snap_edges.append(('y', best_y_guide));  adj_snapped = True

        top_aligned_partner = align_partner if (align_y_guide is not None and best_y_guide == align_y_guide) else None
        return result_nx, result_ny, snap_edges, adj_snapped, x_snap_partner

    def _compute_guide_lines(self, snap_edges):
        """스냅 엣지를 전체 화면 폭/높이의 선 좌표로 변환."""
        if not snap_edges:
            return []
        screens = QApplication.screens()
        total = screens[0].geometry()
        for s in screens[1:]:
            total = total.united(s.geometry())
        lines = []
        for axis, coord in snap_edges:
            if axis == 'x':
                lines.append((coord, total.top(), coord, total.bottom()))
            else:
                lines.append((total.left(), coord, total.right(), coord))
        return lines

    def toggle_merge(self, other):
        if self._link_group and other in self._link_group:
            # 해체: 그룹 전체 해산
            group = list(self._link_group)
            left = self if self.x() <= other.x() else other
            clicked_btn = left._merge_btn
            for member in group:
                member._link_group = None
                member._set_inner_handles_enabled(True)
                clear_window_merge_group(member.window_id)
                if member._merge_btn and member._merge_btn is not clicked_btn:
                    member._merge_btn.hide()
                    member._merge_btn = None
            if clicked_btn:
                clicked_btn.set_merged(False)
        else:
            # 합체: 그룹 합치기
            if self._link_group and other._link_group:
                new_group = self._link_group | other._link_group
            elif self._link_group:
                new_group = set(self._link_group) | {other}
            elif other._link_group:
                new_group = set(other._link_group) | {self}
            else:
                new_group = {self, other}
            group_id = get_next_merge_group_id()
            for member in new_group:
                member._link_group = new_group
                member._set_inner_handles_enabled(False)
                set_window_merge_group(member.window_id, group_id)
            left = self if self.x() <= other.x() else other
            if left._merge_btn:
                left._merge_btn.set_merged(True)

    def reestablish_link(self, members):
        """재시작 시 DB merge_group_id로부터 합체 관계 복원."""
        if len(members) < 2:
            return
        new_group = set(members)
        for m in new_group:
            m._link_group = new_group
            m._set_inner_handles_enabled(False)
        sorted_members = sorted(new_group, key=lambda w: w.x())
        for i in range(len(sorted_members) - 1):
            left = sorted_members[i]
            right = sorted_members[i + 1]
            btn = MergeButton()
            btn.set_windows(left, right)
            btn.set_merged(True)
            left._merge_btn = btn
            btn.update_position(left, right)
            btn.show()
            QTimer.singleShot(300, btn._raise_timer.start)

    def _set_inner_handles_enabled(self, enabled):
        """합체된 안쪽 테두리 핸들 활성화/비활성화."""
        if not hasattr(self, '_handles'):
            return
        if self._link_group is None:
            for h in self._handles.values():
                h.setAttribute(Qt.WA_TransparentForMouseEvents, False)
                h.setCursor(EdgeHandle._CURSORS[h.edge])
            return
        inner = []
        for member in self._link_group:
            if member is self:
                continue
            if member.x() > self.x():
                inner += ['right', 'top-right', 'bottom-right']
            elif member.x() < self.x():
                inner += ['left', 'top-left', 'bottom-left']
        for edge, h in self._handles.items():
            if edge in inner:
                h.setAttribute(Qt.WA_TransparentForMouseEvents, not enabled)
                h.setCursor(Qt.ArrowCursor if not enabled else EdgeHandle._CURSORS[edge])

    def moveEvent(self, event):
        super().moveEvent(event)
        if self._link_group and not self._is_syncing_move:
            dx = self.x() - event.oldPos().x()
            dy = self.y() - event.oldPos().y()
            for member in self._link_group:
                if member is self:
                    continue
                member._is_syncing_move = True
                member._snapped_pos = None
                member._current_snap_edges = []
                member._snap_is_adj = False
                member.move(member.x() + dx, member.y() + dy)
                member._is_syncing_move = False
            for member in self._link_group:
                if member._merge_btn and member._merge_btn.isVisible():
                    btn = member._merge_btn
                    btn.update_position(btn._win_a, btn._win_b)

    def _setup_resize_handles(self):
        self._handles = {
            edge: EdgeHandle(self, edge)
            for edge in ('left', 'right', 'bottom', 'bottom-left', 'bottom-right',
                         'top', 'top-left', 'top-right')
        }
        self._reposition_handles()

    def _reposition_handles(self):
        w  = self.width()
        h  = self.height()
        es = EdgeHandle.EDGE
        cs = EdgeHandle.CORNER
        collapsed = self.collapsed

        # 상단 핸들 (좌우 모서리 제외)
        self._handles['top'].setGeometry(cs, 0, max(0, w - 2 * cs), es)
        self._handles['top-left'].setGeometry(0, 0, cs, cs)
        self._handles['top-right'].setGeometry(w - cs, 0, cs, cs)

        # 좌우: 타이틀바 아래부터 하단 모서리 위까지
        side_top = cs
        side_h   = max(0, h - cs * 2)
        self._handles['left'].setGeometry(0, side_top, es, side_h)
        self._handles['right'].setGeometry(w - es, side_top, es, side_h)

        # 하단 및 모서리
        self._handles['bottom'].setGeometry(cs, h - es, max(0, w - 2 * cs), es)
        self._handles['bottom-left'].setGeometry(0, h - cs, cs, cs)
        self._handles['bottom-right'].setGeometry(w - cs, h - cs, cs, cs)

        # 접힌 상태에서는 상하단 핸들 숨김
        for edge in ('top', 'top-left', 'top-right', 'bottom', 'bottom-left', 'bottom-right'):
            self._handles[edge].setVisible(not collapsed)

        for handle in self._handles.values():
            handle.raise_()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if hasattr(self, '_handles'):
            self._reposition_handles()
        if not self.collapsed:
            self.expanded_height = self.height()

    def _apply_scale_to_inputs(self):
        s = self._scale
        self.input_name.setFont(QFont('Malgun Gothic', int(11 * s)))
        self.input_date.setFont(QFont('Malgun Gothic', int(11 * s)))
        self.input_date.setFixedWidth(int(115 * s))
        self.input_time.setFont(QFont('Malgun Gothic', int(11 * s)))
        self.input_time.setFixedWidth(int(68 * s))
        self.input_recurrence.setFont(QFont('Malgun Gothic', int(10 * s)))
        self.input_recurrence.setFixedWidth(int(90 * s))
        _row_h = self.input_name.sizeHint().height()
        self.input_date.setFixedHeight(_row_h)
        self.input_time.setFixedHeight(_row_h)
        self.input_recurrence.setFixedHeight(_row_h)
        self.memo_editor.setFont(pr_font(int(12 * s)))

    def _set_scale(self, scale, menu=None):
        self._scale = scale
        self._apply_scale_to_inputs()
        self._refresh_tasks()
        self.save_state()
        if menu:
            menu.close()

    def save_state(self):
        pos = self.pos()
        update_window(
            self.window_id,
            x=pos.x(), y=pos.y(),
            width=self.width(),
            height=self.expanded_height,
            collapsed=self.collapsed,
            color=self._bg_color,
            scale=self._scale,
        )

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, '_first_shown', False):
            self._first_shown = True
            QTimer.singleShot(0, lambda: (self.activateWindow(), (self.memo_editor if self._memo_mode else self.input_name).setFocus()))

    def _raise_merge_btn(self):
        """Win32 SetWindowPos로 합체 버튼을 노트 창보다 위, 다른 앱보다 아래로."""
        if self._merge_btn is None or not self._merge_btn.isVisible():
            return
        try:
            HWND_TOP = 0
            SWP_NOACTIVATE = 0x0010
            SWP_NOSIZE = 0x0001
            SWP_NOMOVE = 0x0002
            ctypes.windll.user32.SetWindowPos(
                int(self._merge_btn.winId()), HWND_TOP,
                0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)
        except Exception:
            pass


    def closeEvent(self, e):
        if self._link_group:
            self._link_group.discard(self)
            remaining = self._link_group
            self._link_group = None
            clear_window_merge_group(self.window_id)
            if len(remaining) == 1:
                sole = next(iter(remaining))
                sole._link_group = None
                sole._set_inner_handles_enabled(True)
                clear_window_merge_group(sole.window_id)
            elif len(remaining) > 1:
                for m in remaining:
                    m._link_group = remaining
                    m._set_inner_handles_enabled(False)
        if self._merge_btn is not None:
            self._merge_btn.hide()
            self._merge_btn = None
        if not getattr(self, '_deleted', False):
            self.save_state()
            if self._memo_save_timer and self._memo_save_timer.isActive():
                self._memo_save_timer.stop()
                self._flush_memo_save()
        e.accept()
