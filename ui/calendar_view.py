"""ui/calendar_view.py — 월간 달력 / 날짜별 상세 / 업무 편집 대화상자."""
from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QMessageBox, QDialog, QGridLayout, QPlainTextEdit, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate, QEvent, QPoint, QPointF, QRectF, QSettings, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QMouseEvent
from .common import px, icon_png

from .theme import DARK as _DARK


class P:
    """달력 색 (라이트/다크)."""
    if _DARK:
        BG, SURFACE, LINE = '#191919', '#202020', '#2F2F2F'
        TEXT, TEXT2, MUTED, FAINT = '#E3E3E1', '#B5B5B3', '#8A8A88', '#5A5A58'
        HOVER, CHIP, CHIP_HOVER = '#232323', '#2A2A2A', '#333333'
        SUN, SAT, TODAY = '#E06C67', '#5B9BD5', '#EB5757'
        PRIMARY, PRIMARY_TEXT, PRIMARY_HOVER = '#E3E3E1', '#191919', '#FFFFFF'
        DOT, DOT_PRIORITY, DANGER = '#9A7BD8', '#E9A23B', '#FF6B6B'
    else:
        BG, SURFACE, LINE = '#FFFFFF', '#FBFBFA', '#EDEDEC'
        TEXT, TEXT2, MUTED, FAINT = '#37352F', '#5F5E5B', '#9B9A97', '#C7C6C3'
        HOVER, CHIP, CHIP_HOVER = '#F7F7F5', '#F1F1EF', '#E9E9E7'
        SUN, SAT, TODAY = '#D44C47', '#337EA9', '#EB5757'
        PRIMARY, PRIMARY_TEXT, PRIMARY_HOVER = '#2F3437', '#FFFFFF', '#1F2224'
        DOT, DOT_PRIORITY, DANGER = '#9A6DD7', '#E9A23B', '#D93025'


class CR:
    """새 달력 색 (크림): 창 바탕 · 주말 · 옆 패널은 따뜻한 크림, 날짜 칸은 흰색."""
    BG, CELL, CELL_LINE, WKEND = '#FFFDF8', '#FFFFFF', '#F3EADA', '#FFFBF2'
    SIDE, SIDE_LINE, BTN_LINE, HOVER = '#FFF7E6', '#F2E6CC', '#EDE3CF', '#FFF5E1'
    TEXT, TEXT2, MUTED, FAINT = '#2B2A27', '#37352F', '#A39478', '#CFC6B4'
    TODAY, TODAY_BG, SEL_LINE = '#F59E0B', '#FFFBEB', '#E2C98F'
    SUN, SAT, DANGER = '#D44C47', '#337EA9', '#D93025'
    PIN_BG, PIN_LINE, PIN_TEXT = '#EEF2FF', '#C7D2FE', '#4F46E5'
    # 업무 알약: (바탕, 글자, 점)
    CHIP = {
        'norm': ('#F7F1E4', '#37352F', '#A8A08F'),
        'soon': ('#FFF1E6', '#B4540A', '#F76B15'),
        'over': ('#FDECEC', '#C5221F', '#E5484D'),
        'star': ('#FEF6D8', '#8A6100', '#E9A23B'),
        'done': (None, '#B0AFAB', '#D6D5D2'),
    }


def chip_kind(task, today=None) -> str:
    """달력 알약 색 종류: 완료 · 지운 업무 → done, 기한 지남 → over, ★ → star, 3일 안 → soon, 나머지 norm."""
    if task.get('_source') == 'history' or task.get('strikethrough'):
        return 'done'
    today = today or date.today()
    try:
        d = date.fromisoformat((task.get('deadline') or '')[:10])
    except ValueError:
        return 'star' if task.get('priority') else 'norm'
    days = (d - today).days
    if days < 0:
        return 'over'
    if task.get('priority'):
        return 'star'
    return 'soon' if days <= 3 else 'norm'


def _ui_font(pt, weight=QFont.Normal):
    f = QFont('Pretendard GOV', 10)
    f.setFamilies(['Pretendard GOV', 'Pretendard', 'Malgun Gothic'])
    f.setPointSizeF(pt)
    f.setWeight(weight)
    return f


def _icon_button(icon, tip, size=30):
    import qtawesome as qta
    b = QPushButton()
    b.setIcon(qta.icon(icon, color=P.TEXT2))
    b.setIconSize(QSize(16, 16))
    b.setFixedSize(size, size)
    b.setToolTip(tip)
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"QPushButton {{ background: transparent; border: none; border-radius: 6px; }}"
                    f"QPushButton:hover {{ background: {P.CHIP}; }}"
                    f"QPushButton:pressed {{ background: {P.CHIP_HOVER}; }}")
    return b


def _ghost_button(text):
    b = QPushButton(text)
    b.setFont(_ui_font(10, QFont.DemiBold))
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"QPushButton {{ background: transparent; border: 1px solid {P.LINE}; border-radius: 7px;"
                    f" padding: 5px 12px; color: {P.TEXT}; }}"
                    f"QPushButton:hover {{ background: {P.HOVER}; }}"
                    f"QPushButton:pressed {{ background: {P.CHIP}; }}")
    return b


def _primary_button(text):
    b = QPushButton(text)
    b.setFont(_ui_font(10, QFont.DemiBold))
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"QPushButton {{ background: {P.PRIMARY}; border: none; border-radius: 7px;"
                    f" padding: 6px 14px; color: {P.PRIMARY_TEXT}; }}"
                    f"QPushButton:hover {{ background: {P.PRIMARY_HOVER}; }}")
    return b


class TaskButton(QPushButton):
    """칸 안의 할 일 한 줄: 옅은 회색 알약 + 왼쪽 작은 점 (중요는 주황 점)."""

    def __init__(self, task_id, task_text, strikethrough=False, priority=False, past=False, parent=None, kind=None):
        super().__init__(parent)
        self.kind = kind or ('done' if (past or strikethrough) else ('star' if priority else 'norm'))
        self.task_id = task_id
        self.full_text = task_text
        self._strikethrough = strikethrough
        self._priority = priority
        self._past = past                  # 업무 목록에서 지운 업무 (지난 기록) → 회색
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        f = _ui_font(9)
        f.setStrikeOut(strikethrough)
        self.setFont(f)
        self.setFixedHeight(QFontMetrics(f).height() + 8)
        self.setToolTip(f'{task_text}\n(지운 업무 · 지난 기록)' if past else task_text)
        self._hover = False

    def enterEvent(self, e):
        self._hover = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hover = False
        self.update()
        super().leaveEvent(e)

    def paintEvent(self, e):
        from PyQt5.QtGui import QPainter, QColor, QPolygonF
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        bg, fg, dot = CR.CHIP[self.kind]
        p.setPen(Qt.NoPen)
        if bg:
            c = QColor(bg)
            if self._hover:
                c = c.darker(104)
            p.setBrush(c)
            p.drawRoundedRect(r, 6, 6)
        elif self._hover:
            p.setBrush(QColor(CR.HOVER))
            p.drawRoundedRect(r, 6, 6)
        p.setBrush(QColor(dot))
        cy, cx = r.center().y(), 9.0
        if self.kind == 'star':                          # ★ 중요: 작은 마름모
            p.drawPolygon(QPolygonF([QPointF(cx, cy - 4), QPointF(cx + 4, cy), QPointF(cx, cy + 4), QPointF(cx - 4, cy)]))
        else:
            p.drawEllipse(QPointF(cx, cy), 3, 3)
        f = QFont(self.font())
        if self.kind == 'star':
            f.setWeight(QFont.DemiBold)
        f.setStrikeOut(self.kind == 'done')
        p.setFont(f)
        p.setPen(QColor(fg))
        fm = QFontMetrics(f)
        text_rect = self.rect().adjusted(18, 0, -6, 0)
        p.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter,
                   fm.elidedText(self.full_text, Qt.ElideRight, max(10, text_rect.width())))
        p.end()


class TaskIndicatorRow(QWidget):
    """할 일 알약 한 줄 (예전 이름 유지)."""
    def __init__(self, task, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._btn = TaskButton(
            task['id'], task['name'],
            strikethrough=bool(task.get('strikethrough', 0)),
            priority=bool(task.get('priority', 0)),
            past=task.get('_source') == 'history',
            kind=chip_kind(task),
        )
        layout.addWidget(self._btn, 1)

    def button(self):
        return self._btn


class DayCellWidget(QFrame):
    clicked = pyqtSignal(str)
    overflow_clicked = pyqtSignal(str)

    MAX_VISIBLE = 3

    def __init__(self, date_str, day_number, is_today=False, col=None, is_other_month=False):
        super().__init__()
        self.date_str = date_str
        self._hidden_count = 0
        self._overflow_label = None
        self._other = is_other_month
        self._today = is_today
        self._weekend = col in (0, 6)
        self._selected = False
        self.setObjectName('DayCell')
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)   # 내용과 상관없이 칸 크기를 똑같이
        self.setMinimumSize(0, 0)
        self.setCursor(Qt.ArrowCursor if is_other_month else Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(7, 7, 7, 6)
        layout.setSpacing(3)

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        self.day_label = QLabel(str(day_number))
        self.day_label.setFont(_ui_font(10.5, QFont.Bold if is_today else QFont.DemiBold))
        self.day_label.setAlignment(Qt.AlignCenter)
        d = QFontMetrics(self.day_label.font()).height() + 4
        if is_today:                                     # 오늘: 주황 동그라미 안에 흰 숫자
            self.day_label.setFixedSize(d, d)
            self.day_label.setStyleSheet(f"QLabel {{ color: #FFFFFF; background: {CR.TODAY}; border-radius: {d // 2}px; }}")
        else:
            color = (CR.FAINT if is_other_month else CR.SUN if col == 0 else CR.SAT if col == 6 else CR.TEXT2)
            self.day_label.setFixedHeight(d)
            self.day_label.setStyleSheet(f"QLabel {{ color: {color}; background: transparent; padding: 0 3px; }}")
        top.addWidget(self.day_label, 0, Qt.AlignLeft | Qt.AlignTop)
        top.addStretch()
        layout.addLayout(top)

        self.task_layout = QVBoxLayout()
        self.task_layout.setSpacing(3)
        layout.addLayout(self.task_layout)
        layout.addStretch()
        self._apply_style()

    def _apply_style(self):
        """둥근 카드 칸: 흰 바탕 · 옅은 크림 테두리. 주말은 크림 바탕, 오늘은 주황 테두리, 고른 날은 진한 크림 테두리,
        다른 달은 바탕 · 테두리 없이."""
        if self._other:
            self.setStyleSheet("QFrame#DayCell { background: transparent; border: 1px solid transparent; border-radius: 10px; }")
            return
        bg = CR.TODAY_BG if self._today else (CR.WKEND if self._weekend else CR.CELL)
        if self._today:
            border = f'2px solid {CR.TODAY}'
        elif self._selected:
            border = f'2px solid {CR.SEL_LINE}'
        else:
            border = f'1px solid {CR.CELL_LINE}'
        self.setStyleSheet(f"QFrame#DayCell {{ background: {bg}; border: {border}; border-radius: 10px; }}"
                           f"QFrame#DayCell:hover {{ background: {CR.HOVER}; }}")

    def set_selected(self, on: bool):
        if self._selected != on:
            self._selected = on
            self._apply_style()

    def set_edges(self, right: bool, bottom: bool):
        """(예전 표 모양의 칸 선 — 카드 칸으로 바뀌어 이제는 모양만 다시 칠함)"""
        self._apply_style()

    def clear_tasks(self):
        self._tasks = []
        self._clear_widgets()

    def _clear_widgets(self):
        while self.task_layout.count():
            item = self.task_layout.takeAt(0)
            if item.widget():
                item.widget().hide()
                item.widget().deleteLater()
        self._hidden_count = 0
        self._overflow_label = None

    def _visible_btn_count(self):
        return sum(
            1 for i in range(self.task_layout.count())
            if isinstance(self.task_layout.itemAt(i).widget(), TaskIndicatorRow)
        )

    def _make_overflow_label(self, n):
        lbl = QLabel(f'+{n}개 더')
        lbl.setCursor(Qt.PointingHandCursor)
        lbl.setFont(_ui_font(8.5, QFont.DemiBold))
        lbl.setStyleSheet(f"QLabel {{ color: #6B6458; background: #F1EADB; border-radius: 8px; padding: 1px 8px; }}"
                          f"QLabel:hover {{ background: #E9DFCB; color: {CR.TEXT}; }}")
        lbl.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        lbl.mousePressEvent = lambda e: self.overflow_clicked.emit(self.date_str)
        return lbl

    def _capacity(self):
        """지금 칸 높이에 들어가는 알약 수 (최소 1, 최대 MAX_VISIBLE)."""
        chip_h = QFontMetrics(_ui_font(9)).height() + 8
        top = self.day_label.sizeHint().height() + 6 + 6 + 3      # 날짜 + 위아래 여백 + 간격
        avail = self.height() - top
        return max(1, min(self.MAX_VISIBLE, (avail + 3) // (chip_h + 3)))

    def _relayout(self):
        """할 일을 칸 높이에 맞춰 다시 배치: 다 안 들어가면 마지막 자리에 '+N개 더'."""
        self._clear_widgets()
        tasks = getattr(self, '_tasks', [])
        if not tasks:
            return
        cap = self._capacity()
        if len(tasks) <= cap:
            shown = tasks
        else:   # '+N개 더' 글줄은 알약보다 낮으니 그만큼 알약을 더 넣는다
            chip_h = QFontMetrics(_ui_font(9)).height() + 8
            more_h = QFontMetrics(_ui_font(8.5)).height() + 3
            avail = self.height() - (self.day_label.sizeHint().height() + 15)
            n = max(0, min(self.MAX_VISIBLE, (avail - more_h) // (chip_h + 3)))
            shown = tasks[:n]
        for task in shown:
            row = TaskIndicatorRow(task)
            row.button().clicked.connect(lambda checked: self.clicked.emit(self.date_str))
            if self._other:                              # 다른 달 칸은 흐리게
                from PyQt5.QtWidgets import QGraphicsOpacityEffect
                eff = QGraphicsOpacityEffect(row)
                eff.setOpacity(0.55)
                row.setGraphicsEffect(eff)
            self.task_layout.addWidget(row)
        rest = len(tasks) - len(shown)
        if rest:
            self._hidden_count = rest
            self._overflow_label = self._make_overflow_label(rest)
            self.task_layout.addWidget(self._overflow_label, 0, Qt.AlignLeft)
        self._shown_cap = cap

    def add_task(self, task):
        self._tasks = getattr(self, '_tasks', []) + [task]
        self._relayout()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if getattr(self, '_tasks', None) and self._capacity() != getattr(self, '_shown_cap', None):
            self._relayout()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.date_str)
        super().mousePressEvent(event)


class DetailTaskRow(QWidget):
    """옆 패널의 업무 한 줄: 흰 카드 · ★ · 이름 · 시간/메모 · D-day 알약 · (마우스를 올리면) 편집 · 삭제."""
    edit_requested = pyqtSignal()
    delete_requested = pyqtSignal()

    def __init__(self, task, parent=None):
        super().__init__(parent)
        self.setObjectName('DetailTaskRow')
        self.setAttribute(Qt.WA_StyledBackground, True)
        past = task.get('_source') == 'history'
        done = past or bool(task.get('strikethrough'))
        self.setStyleSheet(
            f"QWidget#DetailTaskRow {{ background: {'transparent' if past else CR.CELL}; border-radius: 10px;"
            f" border: 1px {'dashed' if past else 'solid'} {CR.CELL_LINE}; }}"
            f"QWidget#DetailTaskRow:hover {{ border: 1px solid {CR.SEL_LINE}; }}"
        )
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 8, 10)
        layout.setSpacing(9)

        name_font = _ui_font(10.5, QFont.DemiBold)
        star = QLabel('★' if task.get('priority') else '☆')
        star.setFont(_ui_font(11.5))
        star.setFixedHeight(QFontMetrics(name_font).height())
        star.setStyleSheet(f"color: {'#E9A23B' if (task.get('priority') and not done) else CR.FAINT}; background: transparent; border: none;")
        layout.addWidget(star, 0, Qt.AlignTop)

        text_col = QVBoxLayout()
        text_col.setSpacing(3)
        if done:
            name_font.setStrikeOut(True)
        name = task['name']
        name_lbl = QLabel(name if len(name) <= 40 else name[:40] + '…')
        name_lbl.setFont(name_font)
        name_lbl.setToolTip(name)
        name_lbl.setStyleSheet(f"color: {CR.MUTED if done else CR.TEXT}; background: transparent; border: none;")
        name_lbl.setMinimumWidth(1)
        name_lbl.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self._name_lbl = name_lbl
        text_col.addWidget(name_lbl)
        sub = []
        dl = task.get('deadline') or ''
        if len(dl) > 10:
            sub.append(dl[11:16])
        notes = (task.get('notes') or '').strip()
        if notes:
            first = notes.splitlines()[0]
            sub.append(first if len(first) <= 30 else first[:30] + '…')
        if sub:
            notes_lbl = QLabel(' · '.join(sub))
            notes_lbl.setFont(_ui_font(9))
            notes_lbl.setStyleSheet(f"color: {CR.MUTED}; background: transparent; border: none;")
            text_col.addWidget(notes_lbl)
        if past:                                         # 지운 업무: 보기 전용 (고치려면 지난 기록에서)
            tag = QLabel('지난 업무')
            tag.setFont(_ui_font(8.5, QFont.Medium))
            tag.setToolTip('업무 목록에서 지운 업무예요. 지난 기록에서 되살리거나 완전히 지울 수 있어요.')
            tag.setStyleSheet(f"color: {CR.MUTED}; background: #F1EADB; border-radius: 4px; padding: 2px 6px; border: none;")
            layout.addLayout(text_col, 1)
            layout.addWidget(tag, 0, Qt.AlignTop)
            return
        layout.addLayout(text_col, 1)

        right = QVBoxLayout()
        right.setSpacing(4)
        if dl and not done:
            from .dates import calc_dday
            dd = calc_dday(dl)
            kind = chip_kind(task)
            col = {'over': '#E5484D', 'soon': '#F76B15', 'star': '#F76B15' if dd == 'D-day' else '#E9A23B'}.get(kind, '#B8AE9A')
            pill = QLabel(dd.replace('D+', 'D+') + (' 지남' if dd.startswith('D+') else ''))
            pill.setFont(_ui_font(8.5, QFont.Bold))
            ph = QFontMetrics(pill.font()).height() + 4
            pill.setFixedHeight(ph)
            pill.setStyleSheet(f"color: #FFFFFF; background: {col}; border-radius: {ph // 2}px; padding: 0 8px; border: none;")
            right.addWidget(pill, 0, Qt.AlignRight)
        btns = QHBoxLayout()
        btns.setSpacing(0)
        self._btn_edit = _icon_button('ph.pencil-simple', '편집', 26)
        self._btn_edit.clicked.connect(self.edit_requested)
        self._btn_del = _icon_button('ph.trash', '삭제', 26)
        self._btn_del.clicked.connect(self.delete_requested)
        for b_ in (self._btn_edit, self._btn_del):
            b_.setStyleSheet(b_.styleSheet().replace(f"background: {P.CHIP}", f"background: {CR.HOVER}"))
            b_.setVisible(False)                         # 마우스를 올렸을 때만
            btns.addWidget(b_)
        right.addLayout(btns)
        layout.addLayout(right)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        lbl = getattr(self, '_name_lbl', None)
        if lbl is not None:
            full = lbl.toolTip()
            lbl.setText(QFontMetrics(lbl.font()).elidedText(full, Qt.ElideRight, max(20, lbl.width())))

    def enterEvent(self, e):
        for b_ in (getattr(self, '_btn_edit', None), getattr(self, '_btn_del', None)):
            if b_ is not None:
                b_.setVisible(True)
        super().enterEvent(e)

    def leaveEvent(self, e):
        for b_ in (getattr(self, '_btn_edit', None), getattr(self, '_btn_del', None)):
            if b_ is not None:
                b_.setVisible(False)
        super().leaveEvent(e)


class DayDetailPanel(QWidget):
    """달력 오른쪽 옆 패널: 고른 날의 할 일 (추가 · 편집 · 삭제)."""
    task_edited = pyqtSignal()
    task_deleted = pyqtSignal()
    close_requested = pyqtSignal()

    def __init__(self, open_windows, parent=None):
        super().__init__(parent)
        self._open_windows = open_windows
        self.current_date_str = None
        self._build_ui()

    def _build_ui(self):
        self.setObjectName('DayDetailPanel')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(f"QWidget#DayDetailPanel {{ background: {CR.SIDE}; border-left: 1px solid {CR.SIDE_LINE}; }}")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 14, 14)
        layout.setSpacing(8)

        # 맨 위 줄: (창 단추 자리 ─ ✕) — 달력 창이 채운다
        self.top_bar = QHBoxLayout()
        self.top_bar.setContentsMargins(0, 0, 0, 0)
        self.top_bar.setSpacing(2)
        self.top_bar.addStretch()
        layout.addLayout(self.top_bar)

        self._kicker = QLabel()
        self._kicker.setFont(_ui_font(9.5, QFont.DemiBold))
        self._kicker.setStyleSheet(f"color: {CR.TODAY}; background: transparent; border: none;")
        layout.addWidget(self._kicker)
        self._date_label = QLabel()
        self._date_label.setFont(_ui_font(16, QFont.Bold))
        self._date_label.setStyleSheet(f"color: {CR.TEXT}; background: transparent; border: none;")
        layout.addWidget(self._date_label)
        self._count_label = QLabel()
        self._count_label.setFont(_ui_font(9))
        self._count_label.setStyleSheet(f"color: {CR.MUTED}; background: transparent; border: none;")
        layout.addWidget(self._count_label)
        layout.addSpacing(4)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }"
                             "QScrollBar:vertical { width: 6px; background: transparent; }"
                             f"QScrollBar::handle:vertical {{ background: {CR.SIDE_LINE}; border-radius: 3px; min-height: 30px; }}"
                             "QScrollBar::add-line, QScrollBar::sub-line { height: 0; }")
        self._task_container = QWidget()
        self._task_container.setStyleSheet("background: transparent;")
        self._task_list_layout = QVBoxLayout(self._task_container)
        self._task_list_layout.setSpacing(8)
        self._task_list_layout.setContentsMargins(0, 0, 4, 0)
        self._btn_add = QPushButton('＋  이 날에 업무 추가')
        self._btn_add.setFont(_ui_font(10, QFont.Medium))
        self._btn_add.setCursor(Qt.PointingHandCursor)
        self._btn_add.setStyleSheet(
            f"QPushButton {{ text-align: left; padding: 10px 12px; color: {CR.MUTED}; background: transparent;"
            f" border: 1.5px dashed #DCCFB4; border-radius: 10px; }}"
            f"QPushButton:hover {{ color: {CR.TEXT}; background: #FFFDF6; border-color: {CR.SEL_LINE}; }}")
        self._btn_add.clicked.connect(self._on_add_task)
        self._task_list_layout.addWidget(self._btn_add)
        self._task_list_layout.addStretch()
        scroll.setWidget(self._task_container)
        layout.addWidget(scroll, 1)

    def load(self, date_str, tasks):
        self.current_date_str = date_str
        try:
            d = date.fromisoformat(date_str)
            weekday = ['월', '화', '수', '목', '금', '토', '일'][d.weekday()]
            self._date_label.setText(f'{d.month}월 {d.day}일 ({weekday})')
            diff = (d - date.today()).days
            self._kicker.setText({0: '오늘', 1: '내일', -1: '어제'}.get(diff, f'{diff}일 뒤' if diff > 0 else f'{-diff}일 전'))
            self._kicker.setStyleSheet(f"color: {CR.TODAY if diff == 0 else CR.MUTED}; background: transparent; border: none;")
        except (ValueError, AttributeError):
            self._date_label.setText(date_str)
            self._kicker.setText('')
        live = [t for t in tasks if t.get('_source') != 'history']
        past = len(tasks) - len(live)
        parts = [f'할 일 {len(live)}개'] if live else ([] if past else ['할 일 없음'])
        if past:
            parts.append(f'지난 업무 {past}개')
        self._count_label.setText(' · '.join(parts))

        while self._task_list_layout.count() > 2:          # [＋ 추가] 단추와 끝 여백만 남기고 비움
            item = self._task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for i, task in enumerate(tasks):
            row = DetailTaskRow(task)
            row.edit_requested.connect(lambda t=task: self._on_edit_task(t))
            row.delete_requested.connect(lambda t=task: self._on_delete_task(t))
            self._task_list_layout.insertWidget(i, row)

    def _on_add_task(self):
        if not self._open_windows:
            QMessageBox.information(self, '알림', '열린 노트가 없습니다.\n노트를 먼저 열어주세요.')
            return
        dlg = TaskEditDialog(task=None, default_date=self.current_date_str or '',
                             open_windows=self._open_windows, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.task_edited.emit()

    def _on_edit_task(self, task):
        dlg = TaskEditDialog(task=task, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.task_edited.emit()

    def _on_delete_task(self, task):
        """업무 목록에서 지울 때와 같이 지난 기록에 남기고 지운다 (달력에는 회색으로 남음)."""
        import db
        db.add_task_history(task['window_id'], task['name'], task.get('deadline', ''),
                            strikethrough=task.get('strikethrough', 0), priority=task.get('priority', 0),
                            recurrence=task.get('recurrence', ''), notes=task.get('notes', ''),
                            related_no=task.get('related_no', ''))
        db.delete_task(task['id'])
        self.task_deleted.emit()


class TaskEditDialog(QDialog):
    def __init__(self, task=None, default_date='', open_windows=None, parent=None):
        super().__init__(parent, Qt.Dialog | Qt.WindowCloseButtonHint)
        self.setModal(True)
        self._task = task
        self._is_new = (task is None)
        self._open_windows = open_windows or []
        self.setWindowTitle('새 할 일' if task is None else '할 일 편집')
        field = (f"border: 1px solid {P.LINE}; border-radius: 8px; padding: 7px 10px;"
                 f" background: {P.BG}; color: {P.TEXT}; selection-background-color: {P.CHIP_HOVER};"
                 f" selection-color: {P.TEXT};")
        self.setStyleSheet(
            f"QDialog {{ background: {P.BG}; }} QLabel {{ background: transparent; }}"
            f"QLineEdit, QPlainTextEdit, QDateEdit {{ {field} }}"
            f"QLineEdit:focus, QPlainTextEdit:focus, QDateEdit:focus {{ border: 1px solid {P.TEXT2}; }}"
            f"QDateEdit::drop-down {{ border: none; width: 26px; }}"
            f"QDateEdit::down-arrow {{ image: url({icon_png('ph.caret-down', P.MUTED, 14)}); width: 12px; height: 12px; }}"
        )
        self._build_ui(default_date)

    def _build_ui(self, default_date):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(8)

        title = QLabel('새 할 일' if self._is_new else '할 일 편집')
        title.setFont(_ui_font(13, QFont.Bold))
        title.setStyleSheet(f"color: {P.TEXT};")
        layout.addWidget(title)
        layout.addSpacing(4)

        def label(text):
            lb = QLabel(text)
            lb.setFont(_ui_font(9, QFont.DemiBold))
            lb.setStyleSheet(f"color: {P.MUTED};")
            return lb

        layout.addWidget(label('이름'))
        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText('할 일 이름')
        self._name_edit.setFont(_ui_font(10.5))
        if self._task:
            self._name_edit.setText(self._task['name'])
        layout.addWidget(self._name_edit)

        layout.addWidget(label('마감일'))
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        from .widgets import CustomCalendarWidget
        self._date_edit.setCalendarWidget(CustomCalendarWidget())     # 미니멀 날짜 고르기
        self._date_edit.setFont(_ui_font(10.5))
        self._date_edit.setDisplayFormat('yyyy. M. d. (ddd)')
        from PyQt5.QtCore import QLocale
        self._date_edit.setLocale(QLocale(QLocale.Korean, QLocale.SouthKorea))
        if self._task and self._task.get('deadline'):
            self._date_edit.setDate(QDate.fromString(self._task['deadline'][:10], 'yyyy-MM-dd'))
        elif default_date:
            self._date_edit.setDate(QDate.fromString(default_date, 'yyyy-MM-dd'))
        else:
            self._date_edit.setDate(QDate.currentDate())
        layout.addWidget(self._date_edit)

        layout.addWidget(label('메모'))
        self._memo_edit = QPlainTextEdit()
        self._memo_edit.setPlaceholderText('메모 (선택)')
        self._memo_edit.setFont(_ui_font(10))
        self._memo_edit.setFixedHeight(px(90))
        if self._task:
            from db import get_task_notes
            self._memo_edit.setPlainText(get_task_notes(self._task['id']))
        layout.addWidget(self._memo_edit)
        layout.addSpacing(6)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        hint = QLabel('Ctrl+Enter 저장')
        hint.setFont(_ui_font(8.5))
        hint.setStyleSheet(f"color: {P.FAINT};")
        btn_row.addWidget(hint)
        btn_row.addStretch()
        btn_cancel = _ghost_button('취소')
        btn_cancel.clicked.connect(self.reject)
        btn_ok = _primary_button('저장')
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._on_save)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        self.setFixedWidth(px(380))

    def _on_save(self):
        from db import update_task as _upd, set_task_notes, add_task as _add
        name = self._name_edit.text().strip()
        if not name:
            self._name_edit.setFocus()
            return
        deadline = self._date_edit.date().toString('yyyy-MM-dd')
        strike = int(self._task.get('strikethrough', 0)) if self._task else 0
        priority = int(self._task.get('priority', 0)) if self._task else 0
        notes = self._memo_edit.toPlainText()

        if self._is_new:
            from db import get_first_window_id
            window_id = get_first_window_id()
            if window_id is None:
                QMessageBox.information(self, '알림', '노트가 없습니다.\n노트를 먼저 만들어주세요.')
                return
            new_id = _add(window_id, name, deadline, strike, priority)
            if new_id and notes:
                set_task_notes(new_id, notes)
        else:
            _upd(self._task['id'], name, deadline, strike, priority,
                 recurrence=self._task.get('recurrence', ''))
            set_task_notes(self._task['id'], notes)
        self.accept()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.reject()
        elif event.modifiers() & Qt.ControlModifier and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self._on_save()
        else:
            super().keyPressEvent(event)


class DayDetailDialog(QDialog):
    task_edited = pyqtSignal()
    task_deleted = pyqtSignal()

    def __init__(self, open_windows, date_str, tasks, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint | Qt.WindowMinMaxButtonsHint)
        self._date_str = date_str
        self._set_title(date_str)
        self.setMinimumSize(px(340), px(400))
        self.setStyleSheet(f"QDialog {{ background: {P.BG}; }}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._panel = DayDetailPanel(open_windows, self)
        self._panel.load(date_str, tasks)
        self._panel.task_edited.connect(self.task_edited)
        self._panel.task_deleted.connect(self.task_deleted)
        self._panel.close_requested.connect(self.close)
        layout.addWidget(self._panel)

        _qs = QSettings('SSNnote', 'SSNnote')
        geo = _qs.value('day_detail_dlg/geometry')
        if geo:
            self.restoreGeometry(geo)
        self.finished.connect(lambda: QSettings('SSNnote', 'SSNnote').setValue('day_detail_dlg/geometry', self.saveGeometry()))

        QApplication.instance().installEventFilter(self)

    def _set_title(self, date_str):
        try:
            d = date.fromisoformat(date_str)
            weekday = ['월', '화', '수', '목', '금', '토', '일'][d.weekday()]
            self.setWindowTitle(f'{d.month}월 {d.day}일 ({weekday})')
        except (ValueError, AttributeError):
            self.setWindowTitle(date_str)

    def reload(self):
        from db import get_calendar_items
        self._panel.load(self._date_str, get_calendar_items(self._date_str))

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress and self.isVisible():
            modal = QApplication.activeModalWidget()
            if modal is not None and modal is not self:
                return False
            if isinstance(event, QMouseEvent):
                w = QApplication.widgetAt(event.globalPos())
                if w is not None:
                    if w.window() is self:
                        return False
                    anc = w
                    while anc is not None:
                        if isinstance(anc, DayCellWidget):
                            return False
                        anc = anc.parent()
                self.close()
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        QApplication.instance().removeEventFilter(self)
        super().closeEvent(event)


PIN_KEY = 'calendar_pin'                    # 'none' · 'top'(항상 위에) · 'desktop'(바탕화면에 붙이기)
STARTUP_KEY = 'calendar_on_startup'         # '1'이면 서서니 노트를 켤 때 달력도 연다
PIN_MODES = (('none', '고정 안 함', '보통 창처럼'),
             ('top', '항상 위에', '다른 창보다 늘 위에 떠 있어요'),
             ('desktop', '바탕화면에 붙이기', '다른 창 아래에서 위젯처럼 · 위치 잠금'))
PIN_LABEL = {'none': '고정', 'top': '항상 위에', 'desktop': '바탕화면'}


def load_pin() -> str:
    try:
        from db import get_setting
        v = get_setting(PIN_KEY, 'none')
    except Exception:
        v = 'none'
    return v if v in PIN_LABEL else 'none'


def startup_enabled() -> bool:
    try:
        from db import get_setting
        return get_setting(STARTUP_KEY, '0') == '1'
    except Exception:
        return False


def set_startup(on: bool):
    from db import set_setting
    set_setting(STARTUP_KEY, '1' if on else '0')


def _round_button(text, tip='', bold=True):
    b = QPushButton(text)
    b.setFont(_ui_font(10, QFont.DemiBold if bold else QFont.Normal))
    b.setCursor(Qt.PointingHandCursor)
    b.setToolTip(tip)
    h = px(34)
    b.setFixedHeight(h)
    b.setStyleSheet(f"QPushButton {{ background: #FFFFFF; border: 1px solid {CR.BTN_LINE}; border-radius: {h // 2}px;"
                    f" padding: 0 {px(14)}px; color: {CR.TEXT2}; }}"
                    f"QPushButton:hover {{ background: {CR.HOVER}; }}"
                    f"QPushButton:pressed {{ background: #F6EBD3; }}")
    return b


def _arrow_button(icon, tip):
    import qtawesome as qta
    b = QPushButton()
    b.setIcon(qta.icon(icon, color=CR.TEXT2))
    b.setIconSize(QSize(px(16), px(16)))
    h = px(34)
    b.setFixedSize(h, h)
    b.setToolTip(tip)
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"QPushButton {{ background: #FFFFFF; border: 1px solid {CR.BTN_LINE}; border-radius: {h // 2}px; }}"
                    f"QPushButton:hover {{ background: {CR.HOVER}; }}"
                    f"QPushButton:pressed {{ background: #F6EBD3; }}")
    return b


def _window_button(icon, tip, danger=False):
    import qtawesome as qta
    b = QPushButton()
    b.setIcon(qta.icon(icon, color='#6B6458'))
    b.setIconSize(QSize(px(14), px(14)))
    b.setFixedSize(px(38), px(30))
    b.setToolTip(tip)
    b.setCursor(Qt.PointingHandCursor)
    hover = '#E5484D' if danger else '#F1E6CF'
    b.setStyleSheet(f"QPushButton {{ background: transparent; border: none; border-radius: 8px; }}"
                    f"QPushButton:hover {{ background: {hover}; }}")
    if danger:
        b._icons = (qta.icon(icon, color='#6B6458'), qta.icon(icon, color='#FFFFFF'))
        b.enterEvent = lambda e, b=b: b.setIcon(b._icons[1])
        b.leaveEvent = lambda e, b=b: b.setIcon(b._icons[0])
    return b


class _HeaderBar(QWidget):
    """달력 머리줄: 왼쪽(연도 · 오늘 · 업무 수) | 가운데(‹ 11월 ›, 창 가운데에 딱) | 오른쪽(📌 고정)."""

    def __init__(self, left: QWidget, center: QWidget, right: QWidget, parent=None):
        super().__init__(parent)
        self._left, self._center, self._right = left, center, right
        for w in (left, center, right):
            w.setParent(self)
        self.setFixedHeight(max(px(44), center.sizeHint().height()))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        h = self.height()
        lw = self._left.sizeHint()
        self._left.setGeometry(0, (h - lw.height()) // 2, lw.width(), lw.height())
        rw = self._right.sizeHint()
        self._right.setGeometry(self.width() - rw.width(), (h - rw.height()) // 2, rw.width(), rw.height())
        cw = self._center.sizeHint()
        x = (self.width() - cw.width()) // 2
        x = max(x, lw.width() + px(12))                    # 창이 좁으면 왼쪽과 겹치지 않게
        self._center.setGeometry(x, (h - cw.height()) // 2, cw.width(), cw.height())

    def _place(self):
        """글이 바뀐 뒤(달 · 업무 수 · 고정 이름) 다시 배치."""
        from PyQt5.QtGui import QResizeEvent
        self.resizeEvent(QResizeEvent(self.size(), self.size()))


class NativeCalendarDialog(QDialog):
    """월간 달력 창 (크림 디자인). Windows 제목표시줄 없이 머리줄이 대신하고, 오른쪽에 고른 날의 할 일 패널.
    📌 고정: 고정 안 함 · 항상 위에 · 바탕화면에 붙이기(다른 창 아래 · 위치 잠금)."""

    def __init__(self, open_windows, parent=None):
        super().__init__(parent)
        from .native_frame import CAPTION_HIDDEN
        self._caption_hidden = CAPTION_HIDDEN
        self.setWindowTitle('달력')
        self._open_windows = open_windows
        self._current_date = date.today()
        self._selected = date.today().isoformat()
        self._task_data = {}
        self._pin = load_pin()
        self._lock_frame = False
        self.setMinimumSize(px(900), px(600))
        self._build_ui()
        self._apply_pin(self._pin, save=False)
        self._load_tasks()

    # ── 화면 ─────────────────────────────────────────────────
    def _build_ui(self):
        self.setObjectName('CalWin')
        self.setStyleSheet(f"QDialog#CalWin {{ background: {CR.BG}; }}")
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        main = QWidget()
        main.setStyleSheet('background: transparent;')
        mv = QVBoxLayout(main)
        mv.setContentsMargins(px(26), px(18), px(22), px(18))
        mv.setSpacing(px(10))

        # 왼쪽: 2026 [오늘] 업무 10 · 지남 1
        left = QWidget()
        lh = QHBoxLayout(left)
        lh.setContentsMargins(0, 0, 0, 0)
        lh.setSpacing(px(10))
        self._year_label = QLabel()
        self._year_label.setFont(_ui_font(18, QFont.Bold))
        self._year_label.setStyleSheet(f"color: {CR.TEXT2}; background: transparent;")
        lh.addWidget(self._year_label)
        self.btn_today = _round_button('오늘', '이번 달 · 오늘로')
        self.btn_today.clicked.connect(self._go_today)
        lh.addWidget(self.btn_today)
        self._info_label = QLabel()
        self._info_label.setTextFormat(Qt.RichText)
        self._info_label.setFont(_ui_font(10, QFont.Medium))
        self._info_label.setStyleSheet(f"color: {CR.MUTED}; background: transparent;")
        lh.addWidget(self._info_label)

        # 가운데: ‹ 11월 ›
        center = QWidget()
        ch = QHBoxLayout(center)
        ch.setContentsMargins(0, 0, 0, 0)
        ch.setSpacing(px(8))
        self.btn_prev = _arrow_button('ph.caret-left-bold', '지난 달')
        self.btn_prev.clicked.connect(self._prev_month)
        self._month_title = QLabel()
        self._month_title.setFont(_ui_font(22, QFont.Bold))
        self._month_title.setAlignment(Qt.AlignCenter)
        self._month_title.setStyleSheet(f"color: {CR.TEXT}; background: transparent;")
        self._month_title.setFixedWidth(QFontMetrics(self._month_title.font()).horizontalAdvance('12월') + px(24))
        self.btn_next = _arrow_button('ph.caret-right-bold', '다음 달')
        self.btn_next.clicked.connect(self._next_month)
        ch.addWidget(self.btn_prev)
        ch.addWidget(self._month_title)
        ch.addWidget(self.btn_next)

        # 오른쪽: 📌 고정 ▾
        right = QWidget()
        rh = QHBoxLayout(right)
        rh.setContentsMargins(0, 0, 0, 0)
        self.btn_pin = QPushButton()
        self.btn_pin.setFont(_ui_font(10, QFont.DemiBold))
        self.btn_pin.setCursor(Qt.PointingHandCursor)
        self.btn_pin.setFixedHeight(px(34))
        self.btn_pin.setToolTip('달력 고정: 고정 안 함 · 항상 위에 · 바탕화면에 붙이기')
        self.btn_pin.clicked.connect(self._show_pin_menu)
        rh.addWidget(self.btn_pin)

        self._header = _HeaderBar(left, center, right)
        mv.addWidget(self._header)

        # 요일
        dow = QHBoxLayout()
        dow.setContentsMargins(2, px(4), 2, 0)
        dow.setSpacing(px(6))
        for i, name in enumerate(['일', '월', '화', '수', '목', '금', '토']):
            lbl = QLabel(name)
            lbl.setFont(_ui_font(9.5, QFont.DemiBold))
            color = CR.SUN if i == 0 else (CR.SAT if i == 6 else CR.MUTED)
            lbl.setStyleSheet(f"color: {color}; background: transparent; padding-left: 9px;")
            dow.addWidget(lbl, 1)
        mv.addLayout(dow)

        self._grid = QGridLayout()
        self._grid.setSpacing(px(6))
        self._grid.setContentsMargins(0, 0, 0, 0)
        mv.addLayout(self._grid, 1)
        root.addWidget(main, 1)

        # 오른쪽 패널 (맨 위에 창 단추 ─ ✕)
        self._panel = DayDetailPanel(self._open_windows, self)
        self._panel.setFixedWidth(px(300))
        self._panel.task_edited.connect(self._on_task_changed)
        self._panel.task_deleted.connect(self._on_task_changed)
        self.btn_win_min = _window_button('ph.minus', '내리기')
        self.btn_win_min.clicked.connect(self.showMinimized)
        self.btn_win_close = _window_button('ph.x', '닫기', danger=True)
        self.btn_win_close.clicked.connect(self.close)
        self._panel.top_bar.addWidget(self.btn_win_min)
        self._panel.top_bar.addWidget(self.btn_win_close)
        root.addWidget(self._panel)

    def _update_pin_button(self):
        on = self._pin != 'none'
        self.btn_pin.setText(f'📌  {PIN_LABEL[self._pin]}  ▾')
        h = px(34)
        if on:
            qss = (f"QPushButton {{ background: {CR.PIN_BG}; border: 1px solid {CR.PIN_LINE}; border-radius: {h // 2}px;"
                   f" padding: 0 {px(14)}px; color: {CR.PIN_TEXT}; }} QPushButton:hover {{ background: #E0E7FF; }}")
        else:
            qss = (f"QPushButton {{ background: #FFFFFF; border: 1px solid {CR.BTN_LINE}; border-radius: {h // 2}px;"
                   f" padding: 0 {px(14)}px; color: {CR.TEXT2}; }} QPushButton:hover {{ background: {CR.HOVER}; }}")
        self.btn_pin.setStyleSheet(qss)
        self.btn_win_min.setVisible(self._pin != 'desktop')     # 바탕화면 위젯은 내리지 않음
        self._header._place()

    def _show_pin_menu(self):
        from . import menu_style as ms
        menu = ms.make_menu(self)
        ms.add_header(menu, '달력 고정')
        for key, label, desc in PIN_MODES:
            ms.add_action(menu, f'{label}  ·  {desc}', lambda k=key: self._apply_pin(k), checked=self._pin == key)
        menu.addSeparator()
        on = startup_enabled()
        ms.add_action(menu, '서서니 노트를 켤 때 달력도 열기', lambda: set_startup(not on), checked=on)
        menu.exec_(self.btn_pin.mapToGlobal(self.btn_pin.rect().bottomLeft()))

    # ── 고정 (항상 위에 · 바탕화면에 붙이기) ─────────────────────
    def _apply_pin(self, mode, save=True):
        mode = mode if mode in PIN_LABEL else 'none'
        self._pin = mode
        if save:
            from db import set_setting
            set_setting(PIN_KEY, mode)
        visible, geo = self.isVisible(), self.geometry()
        flags = Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinMaxButtonsHint
        if self._caption_hidden:
            flags |= Qt.FramelessWindowHint
        if mode == 'top':
            flags |= Qt.WindowStaysOnTopHint
        elif mode == 'desktop':
            flags |= Qt.WindowStaysOnBottomHint | Qt.Tool     # 작업 표시줄에서도 빠짐 (위젯처럼)
        self._lock_frame = mode == 'desktop'
        if self.windowFlags() != flags:
            self.setWindowFlags(flags)
            if self._caption_hidden:
                from .native_frame import setup_native_frame
                self.winId()
                setup_native_frame(self)
            if visible:
                self.setGeometry(geo)
                self.show()
        self._update_pin_button()
        if mode == 'desktop':
            self._send_to_bottom()

    def _send_to_bottom(self):
        import sys
        if sys.platform != 'win32':
            return
        try:
            import ctypes
            ctypes.windll.user32.SetWindowPos(int(self.winId()), 1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010)  # HWND_BOTTOM
        except Exception:
            pass

    def nativeEvent(self, event_type, message):
        if event_type == b'windows_generic_MSG':
            try:
                import ctypes
                from ctypes import wintypes
                msg = wintypes.MSG.from_address(int(message))
                if msg.message == 0x0046 and self._pin == 'desktop' and msg.lParam:   # WM_WINDOWPOSCHANGING
                    class WINDOWPOS(ctypes.Structure):
                        _fields_ = [('hwnd', wintypes.HWND), ('hwndInsertAfter', wintypes.HWND),
                                    ('x', ctypes.c_int), ('y', ctypes.c_int), ('cx', ctypes.c_int),
                                    ('cy', ctypes.c_int), ('flags', ctypes.c_uint)]
                    wp = WINDOWPOS.from_address(msg.lParam)
                    wp.hwndInsertAfter = 1               # 늘 맨 아래 (눌러도 다른 창 위로 올라오지 않음)
                    wp.flags &= ~0x0004                  # SWP_NOZORDER 끔
                if self._caption_hidden:
                    from .native_frame import caption_native_event
                    r = caption_native_event(self, message)
                    if r is not None:
                        return True, r
            except Exception:
                pass
        return super().nativeEvent(event_type, message)

    def _is_caption_point(self, pos) -> bool:
        """머리줄 줄(창 맨 위 ~ 머리줄 아래) 또는 옆 패널 맨 윗줄의 빈 곳이면 Windows 제목표시줄처럼 (끌기 · 더블클릭 최대화)."""
        hdr_bottom = self._header.mapTo(self, self._header.rect().bottomLeft()).y()
        panel_x = self._panel.mapTo(self, self._panel.rect().topLeft()).x()
        panel_top_bottom = self._panel.mapTo(self, QPoint(0, self._panel.top_bar.geometry().bottom())).y() + px(6)
        in_main = pos.x() < panel_x and pos.y() <= hdr_bottom
        in_panel = pos.x() >= panel_x and pos.y() <= panel_top_bottom
        if not (in_main or in_panel):
            return False
        w = self.childAt(pos)
        while w is not None and w is not self:
            if isinstance(w, QPushButton):
                return False
            w = w.parentWidget()
        return True

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape and self._pin == 'desktop':
            return                                       # 바탕화면 위젯은 Esc로 닫지 않음
        super().keyPressEvent(event)

    # ── 데이터 ───────────────────────────────────────────────
    def _load_tasks(self):
        from db import get_calendar_dates, get_calendar_items
        self._task_data = {d: get_calendar_items(d) for d in get_calendar_dates()}
        self._render_calendar()
        self._panel.load(self._selected, self._task_data.get(self._selected, []))

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self._load_tasks()

    def _month_info(self, y, m):
        prefix = f'{y:04d}-{m:02d}'
        today = date.today().isoformat()
        live = [t for d, ts in self._task_data.items() if d.startswith(prefix)
                for t in ts if t.get('_source') != 'history']
        over = [t for t in live if not t.get('strikethrough') and (t.get('deadline') or '')[:10] < today]
        txt = f'업무 {len(live)}'
        if over:
            txt += f" · <span style='color:{CR.DANGER}; font-weight:600;'>지남 {len(over)}</span>"
        return txt

    def _render_calendar(self):
        from datetime import timedelta
        y, m = self._current_date.year, self._current_date.month
        self._year_label.setText(str(y))
        self._month_title.setText(f'{m}월')
        self._info_label.setText(self._month_info(y, m))
        self._header._place()

        while self._grid.count():
            item = self._grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        first = date(y, m, 1)
        start = first - timedelta(days=(first.weekday() + 1) % 7)      # 일요일부터
        today_str = date.today().isoformat()
        last = (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        weeks = ((last - start).days // 7) + 1                           # 그달이 걸친 주 수 (4~6)
        for i in range(weeks * 7):
            d = start + timedelta(days=i)
            d_str = d.isoformat()
            other = d.month != m
            cell = DayCellWidget(d_str, d.day, is_today=(d_str == today_str), col=i % 7, is_other_month=other)
            cell.set_selected(d_str == self._selected and not other)
            self._add_cell_tasks(cell, d_str)
            self._grid.addWidget(cell, i // 7, i % 7)
        for r in range(6):
            self._grid.setRowStretch(r, 1 if r < weeks else 0)
        for c in range(7):
            self._grid.setColumnStretch(c, 1)

    def _add_cell_tasks(self, cell, date_str, connect_signals=True):
        for task in self._task_data.get(date_str, []):
            cell.add_task(task)
        if connect_signals:
            cell.clicked.connect(self._on_cell_clicked)
            cell.overflow_clicked.connect(self._on_cell_clicked)

    def _on_cell_clicked(self, date_str):
        """날짜를 누르면 오른쪽 패널에 그날 할 일 (다른 달 칸이면 그 달로 넘어감)."""
        self._selected = date_str
        d = date.fromisoformat(date_str)
        if (d.year, d.month) != (self._current_date.year, self._current_date.month):
            self._current_date = d.replace(day=1)
            self._render_calendar()
        else:
            for c in self.findChildren(DayCellWidget):
                c.set_selected(c.date_str == date_str and not c._other)
        self._panel.load(date_str, self._task_data.get(date_str, []))

    def _on_task_changed(self):
        self._load_tasks()
        for win in self._open_windows:
            win._refresh_tasks()

    def _prev_month(self):
        y, m = self._current_date.year, self._current_date.month
        self._current_date = date(y - 1, 12, 1) if m == 1 else date(y, m - 1, 1)
        self._render_calendar()

    def _next_month(self):
        y, m = self._current_date.year, self._current_date.month
        self._current_date = date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)
        self._render_calendar()

    def _go_today(self):
        self._current_date = date.today().replace(day=1)
        self._selected = date.today().isoformat()
        self._load_tasks()
