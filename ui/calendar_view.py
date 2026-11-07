"""ui/calendar_view.py — 월간 달력 / 날짜별 상세 / 업무 편집 대화상자."""
import calendar
from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QMenu, QMessageBox, QDialog, QGridLayout, QPlainTextEdit, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate, QEvent, QPointF, QSettings, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QCursor, QMouseEvent
from db import (add_task, delete_task, add_task_history, delete_task_history,
                get_all_tasks, get_all_task_history)
from .common import mi_icon, pr_font, px, icon_png
from .theme import T, tc, tct

class MonthlyCalendarDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('한눈에 - 업무 캘린더')
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint | Qt.WindowMaximizeButtonHint)
        self.setMinimumSize(900, 600)
        self.resize(1000, 700)

        self._current = date.today()
        self._all_tasks = []
        self._all_history = []
        self._cell_widgets = {}

        self._build_ui()
        self._load_data()
        self._render()

    def _build_ui(self):
        self.setStyleSheet(f"""
            QDialog {{ background: {tc("#f5f0ea")}; font-family: 'Malgun Gothic'; }}
            QPushButton {{ font-family: 'Malgun Gothic'; }}
            QLabel {{ font-family: 'Malgun Gothic'; }}
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(0)

        nav = QHBoxLayout()
        nav.setSpacing(8)

        btn_prev = QPushButton()
        btn_prev.setIcon(mi_icon('\ue040', 16, T.TEXT_ON_NOTE))
        btn_prev.setFixedSize(36, 36)
        btn_prev.setStyleSheet(f"""
            QPushButton {{
                background: {T.SURFACE_WARM};
                border: none;
                border-radius: 6px;
                padding: 0;
            }}
            QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
            QPushButton:pressed {{ background: {T.SURFACE_WARM_PRESSED}; }}
        """)
        btn_prev.clicked.connect(self._prev_month)
        nav.addWidget(btn_prev)

        self._title_label = QLabel()
        self._title_label.setAlignment(Qt.AlignCenter)
        self._title_label.setStyleSheet(f"font-size: 18pt; font-weight: bold; color: {T.INK}; padding: 0 16px;")
        nav.addWidget(self._title_label, 1)

        btn_next = QPushButton()
        btn_next.setIcon(mi_icon('\ue041', 16, T.TEXT_ON_NOTE))
        btn_next.setFixedSize(36, 36)
        btn_next.setStyleSheet(f"""
            QPushButton {{
                background: {T.SURFACE_WARM};
                border: none;
                border-radius: 6px;
                padding: 0;
            }}
            QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
            QPushButton:pressed {{ background: {T.SURFACE_WARM_PRESSED}; }}
        """)
        btn_next.clicked.connect(self._next_month)
        nav.addWidget(btn_next)

        nav.addSpacing(12)

        btn_today = QPushButton('오늘')
        btn_today.setFixedHeight(36)
        btn_today.setStyleSheet(f"""
            QPushButton {{
                background: {T.INK};
                border: none;
                border-radius: 6px;
                padding: 0 18px;
                color: {tct("white")};
                font-size: 10pt;
                font-weight: bold;
            }}
            QPushButton:hover {{ background: {tc("#3a3430")}; }}
            QPushButton:pressed {{ background: {tc("#0a0400")}; }}
        """)
        btn_today.clicked.connect(self._go_today)
        nav.addWidget(btn_today)

        root.addLayout(nav)
        root.addSpacing(12)

        calendar_card = QFrame()
        calendar_card.setStyleSheet(f"""
            QFrame {{
                background: {tc("#fff8f5")};
                border-radius: 16px;
                border: 1px solid {T.BORDER_WARM};
            }}
        """)
        card_layout = QVBoxLayout(calendar_card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)

        day_header = QFrame()
        day_header.setStyleSheet(f"background: {tc('#faf4f0')}; border-radius: 16px 16px 0 0;")
        day_header.setFixedHeight(36)
        day_header_layout = QHBoxLayout(day_header)
        day_header_layout.setContentsMargins(10, 0, 10, 0)
        day_header_layout.setSpacing(0)

        weekday_colors = ['#e8b5a8', '#999', '#999', '#999', '#999', '#999', '#a8b8e8']
        for c, wd in enumerate(['일', '월', '화', '수', '목', '금', '토']):
            lbl = QLabel(wd)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(f"font-weight: bold; font-size: 11pt; color: {tct(weekday_colors[c])}; background: transparent;")
            day_header_layout.addWidget(lbl, 1)
        card_layout.addWidget(day_header)

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background: {T.BORDER_WARM};")
        card_layout.addWidget(sep)

        self._grid_container = QWidget()
        self._grid_container.setStyleSheet("background: transparent;")
        self._grid_layout = QGridLayout(self._grid_container)
        self._grid_layout.setContentsMargins(0, 0, 0, 0)
        self._grid_layout.setSpacing(0)
        for r in range(6):
            self._grid_layout.setRowStretch(r, 1)
        for c in range(7):
            self._grid_layout.setColumnStretch(c, 1)

        self._cell_widgets = {}
        for r in range(6):
            for c in range(7):
                cell = QFrame()
                cell.setStyleSheet("QFrame { background: transparent; }")
                cell.setMinimumHeight(90)
                vl = QVBoxLayout(cell)
                vl.setContentsMargins(6, 4, 6, 4)
                vl.setSpacing(2)
                self._grid_layout.addWidget(cell, r, c)
                self._cell_widgets[(r, c)] = (cell, vl)

        card_layout.addWidget(self._grid_container, 1)
        root.addWidget(calendar_card, 1)

        root.addSpacing(12)

        bot = QHBoxLayout()
        self._summary_label = QLabel()
        self._summary_label.setStyleSheet(f"color: {tct('#999')}; font-size: 10pt;")
        bot.addWidget(self._summary_label)
        bot.addStretch()
        btn_close = QPushButton('닫기')
        btn_close.setFixedHeight(36)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                background: {T.SURFACE_WARM};
                border: none;
                border-radius: 6px;
                padding: 0 24px;
                color: {T.TEXT_ON_NOTE};
                font-size: 10pt;
            }}
            QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
            QPushButton:pressed {{ background: {T.SURFACE_WARM_PRESSED}; }}
        """)
        btn_close.clicked.connect(self.accept)
        bot.addWidget(btn_close)
        root.addLayout(bot)

    def _load_data(self):
        self._all_tasks = get_all_tasks()
        self._all_history = get_all_task_history()

    def _tasks_for_date(self, target_date):
        ds = target_date.isoformat()
        items = []
        for t in self._all_tasks:
            if t['deadline'] and t['deadline'].startswith(ds):
                d = dict(t)
                d['_source'] = 'tasks'
                items.append(d)
        for h in self._all_history:
            if h['deadline'] and h['deadline'].startswith(ds):
                d = dict(h)
                d['_source'] = 'history'
                items.append(d)
        items.sort(key=lambda x: (x.get('strikethrough', 0), x.get('_source') == 'history'))
        return items

    def _format_time(self, deadline_str):
        if not deadline_str or len(deadline_str) < 11:
            return ''
        time_part = deadline_str[11:]
        try:
            h, m = time_part.split(':')
            hour = int(h)
            minute = int(m)
            if hour == 0:
                return f'오전 12:{minute:02d}'
            elif hour < 12:
                return f'오전 {hour}:{minute:02d}'
            elif hour == 12:
                return f'오후 12:{minute:02d}'
            else:
                return f'오후 {hour - 12}:{minute:02d}'
        except:
            return ''

    def _render(self):
        y, m = self._current.year, self._current.month
        self._title_label.setText(f'{y}년 {m}월')

        _, dim = calendar.monthrange(y, m)
        first = date(y, m, 1)
        start_col = (first.weekday() + 1) % 7

        for r in range(6):
            for c in range(7):
                self._clear_cell(r, c)

        today = date.today()
        day = 1
        total = 0
        for row in range(6):
            for col in range(7):
                if (row == 0 and col < start_col) or day > dim:
                    self._render_empty_cell(row, col)
                    continue
                cur = date(y, m, day)
                items = self._tasks_for_date(cur)
                total += len(items)
                self._render_day_cell(row, col, day, cur, items, today)
                day += 1

        self._summary_label.setText(f'총 {total}건  ·  우클릭 → 상세/삭제')

    def _render_day_cell(self, row, col, day, cur, items, today):
        cell, vl = self._cell_widgets[(row, col)]
        is_today = cur == today

        date_header = QHBoxLayout()
        date_header.setSpacing(0)
        date_header.setContentsMargins(4, 2, 4, 2)

        dc = T.SUNDAY if col == 0 else T.SATURDAY if col == 6 else T.TEXT_ON_NOTE

        if is_today:
            lbl_d = QLabel(str(day))
            lbl_d.setAlignment(Qt.AlignCenter)
            lbl_d.setFixedSize(26, 26)
            lbl_d.setStyleSheet(f"font-weight: bold; font-size: 12pt; color: {tct('white')}; background: {T.HIGHLIGHT}; border-radius: 13px;")
            date_header.addWidget(lbl_d)
        else:
            lbl_d = QLabel(str(day))
            lbl_d.setStyleSheet(f"font-weight: bold; font-size: 12pt; color: {dc}; background: transparent;")
            date_header.addWidget(lbl_d)

        date_header.addStretch()
        vl.addLayout(date_header)

        max_v = 4
        event_colors = ['#e85d3a', '#3a8ee8', '#9b59b6', '#27ae60', '#e67e22']
        for i, task in enumerate(items[:max_v]):
            tw = self._make_task_widget(task, tc(event_colors[i % len(event_colors)]))
            if tw:
                vl.addWidget(tw)
        if len(items) > max_v:
            more = QLabel(f'+{len(items) - max_v}건 더보기')
            more.setStyleSheet(f"color: {tct('#aaa')}; font-size: 10pt; background: transparent;")
            more.setCursor(QCursor(Qt.PointingHandCursor))
            more.mousePressEvent = lambda e, it=items: self._show_more_popup(it)
            vl.addWidget(more)
        vl.addStretch()

    def _make_task_widget(self, task, dot_color=tc('#4a90d9')):
        w = QWidget()
        w.setCursor(QCursor(Qt.PointingHandCursor))
        wly = QHBoxLayout(w)
        wly.setContentsMargins(4, 1, 4, 1)
        wly.setSpacing(4)

        is_done = task.get('strikethrough', 0)
        is_hist = task.get('_source') == 'history'

        dot = QLabel()
        dot.setFixedSize(7, 7)
        dot.setStyleSheet(f"background: {dot_color}; border-radius: 3px;")
        wly.addWidget(dot, 0, Qt.AlignVCenter | Qt.AlignLeft)

        time_str = self._format_time(task.get('deadline', ''))
        nm = task['name']
        if time_str:
            display = f'{time_str} {nm}'
        else:
            display = nm
        if len(display) > 20:
            display = display[:20] + '…'

        lbl_nm = QLabel(display)
        if is_hist:
            lbl_nm.setStyleSheet(f"color: {tct('#bbb')}; font-size: 11pt; background: transparent;")
        elif is_done:
            f = lbl_nm.font()
            f.setStrikeOut(True)
            lbl_nm.setFont(f)
            lbl_nm.setStyleSheet(f"color: {tct('#aaa')}; font-size: 11pt; background: transparent;")
        else:
            f = lbl_nm.font()
            f.setBold(True)
            lbl_nm.setFont(f)
            lbl_nm.setStyleSheet(f"color: {tct('#333')}; font-size: 11pt; background: transparent;")
        wly.addWidget(lbl_nm, 1)

        w.mousePressEvent = lambda e, t=task: self._on_cell_click(e, t)
        return w

    def _on_cell_click(self, event, task):
        if event.button() == Qt.LeftButton:
            self._show_detail(task)
        elif event.button() == Qt.RightButton:
            self._show_context_menu(event.globalPos(), task)

    def _show_context_menu(self, pos, task):
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{ background: {tc("white")}; border: 1px solid {tc("#ddd")}; border-radius: 6px; }}
            QMenu::item {{ padding: 6px 22px; font-family: 'Malgun Gothic'; font-size: 10pt; }}
            QMenu::item:selected {{ background: {tc("#f0f4ff")}; }}
        """)
        a_detail = menu.addAction('상세 보기')
        a_detail.triggered.connect(lambda: self._show_detail(task))

        if task['_source'] == 'history':
            a_restore = menu.addAction('↩ 복원')
            a_restore.triggered.connect(lambda: self._restore_task(task))

        menu.addSeparator()
        a_del = menu.addAction('삭제')
        a_del.triggered.connect(lambda: self._delete_task_item(task))
        menu.exec_(pos)

    def _show_detail(self, task):
        dlg = QDialog(self)
        dlg.setWindowTitle('업무 상세')
        dlg.setWindowFlags(dlg.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        dlg.setMinimumWidth(320)

        dlg.setStyleSheet(f"""
            QDialog {{ background: {tc("#fff8f5")}; font-family: 'Malgun Gothic'; }}
            QLabel {{ font-family: 'Malgun Gothic'; }}
            QPushButton {{ font-family: 'Malgun Gothic'; }}
        """)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        src = '현재 업무' if task['_source'] == 'tasks' else '지난 기록'
        st = '완료' if task.get('strikethrough') else '진행중'
        dl = task.get('deadline', '없음')
        nt = task.get('notes', '')

        title_font = QFont('Malgun Gothic', 14, QFont.Bold)
        title = QLabel(task['name'])
        title.setFont(title_font)
        title.setStyleSheet(f"color: {T.INK}; background: transparent;")
        title.setWordWrap(True)
        layout.addWidget(title)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"background: {T.SURFACE_WARM}; border: none;")
        sep.setFixedHeight(1)
        layout.addWidget(sep)

        info_style = f"color: {T.TEXT_ON_NOTE}; font-size: 11pt; background: transparent;"
        label_style = f"color: {tct('#999')}; font-size: 10pt; background: transparent;"

        for label_text, value in [('상태', st), ('출처', src), ('마감일', dl)]:
            row = QHBoxLayout()
            row.setSpacing(8)
            lbl = QLabel(label_text)
            lbl.setStyleSheet(label_style)
            lbl.setFixedWidth(50)
            row.addWidget(lbl)
            val = QLabel(str(value))
            val.setStyleSheet(info_style)
            val.setWordWrap(True)
            row.addWidget(val, 1)
            layout.addLayout(row)

        if nt:
            row = QHBoxLayout()
            row.setSpacing(8)
            lbl = QLabel('메모')
            lbl.setStyleSheet(label_style)
            lbl.setFixedWidth(50)
            lbl.setAlignment(Qt.AlignTop)
            row.addWidget(lbl)
            val = QLabel(nt)
            val.setStyleSheet(info_style)
            val.setWordWrap(True)
            row.addWidget(val, 1)
            layout.addLayout(row)

        layout.addSpacing(8)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_close = QPushButton('닫기')
        btn_close.setFixedHeight(32)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                background: {T.SURFACE_WARM};
                border: none;
                border-radius: 6px;
                padding: 0 24px;
                color: {T.TEXT_ON_NOTE};
                font-size: 10pt;
            }}
            QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
        """)
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        dlg.exec_()

    def _show_more_popup(self, items):
        txt = '\n'.join(f"{'✓' if t.get('strikethrough') else '●'} {t['name']} ({t.get('deadline','')})" for t in items)
        QMessageBox.information(self, f'전체 업무 ({len(items)}건)', txt)

    def _delete_task_item(self, task):
        r = QMessageBox.question(self, '삭제 확인', f'"{task["name"]}" 삭제할까요?', QMessageBox.Yes | QMessageBox.No)
        if r != QMessageBox.Yes:
            return
        if task['_source'] == 'tasks':
            add_task_history(task['window_id'], task['name'], task.get('deadline', ''),
                             strikethrough=task.get('strikethrough', 0),
                             priority=task.get('priority', 0),
                             recurrence=task.get('recurrence', ''),
                             notes=task.get('notes', ''),
                             related_no=task.get('related_no', ''))
            delete_task(task['id'])
        else:
            delete_task_history(task['id'])
        self._load_data()
        self._render()
        p = self.parent()
        if p and hasattr(p, '_refresh_tasks'):
            p._refresh_tasks()

    def _restore_task(self, task):
        add_task(task['window_id'], task['name'], task.get('deadline', ''),
                 strikethrough=task.get('strikethrough', 0),
                 priority=task.get('priority', 0),
                 recurrence=task.get('recurrence', ''),
                 notes=task.get('notes', ''),
                 related_no=task.get('related_no', ''))
        delete_task_history(task['id'])
        self._load_data()
        self._render()
        p = self.parent()
        if p and hasattr(p, '_refresh_tasks'):
            p._refresh_tasks()

    def _render_empty_cell(self, row, col):
        cell, _ = self._cell_widgets[(row, col)]
        cell.setStyleSheet(f"QFrame {{ background: {tc('#f5efea')}; border-radius: 8px; }}")

    def _clear_cell(self, row, col):
        cell, vl = self._cell_widgets[(row, col)]
        while vl.count():
            it = vl.takeAt(0)
            w = it.widget()
            if w:
                w.deleteLater()

    def _prev_month(self):
        y, m = self._current.year, self._current.month - 1
        if m == 0:
            y -= 1; m = 12
        self._current = date(y, m, 1)
        self._render()

    def _next_month(self):
        y, m = self._current.year, self._current.month + 1
        if m == 13:
            y += 1; m = 1
        self._current = date(y, m, 1)
        self._render()

    def _go_today(self):
        self._current = date.today()
        self._render()


# ── Native Calendar Widgets ──

# ── 월간 달력 (Notion·Apple 캘린더 같은 미니멀 디자인) ─────────────────
# 흰 바탕 + 얇은 칸 선, 날짜는 왼쪽 위 작은 숫자(오늘은 빨간 동그라미), 할 일은 옅은 회색 알약.

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

    def __init__(self, task_id, task_text, strikethrough=False, priority=False, parent=None):
        super().__init__(parent)
        self.task_id = task_id
        self.full_text = task_text
        self._strikethrough = strikethrough
        self._priority = priority
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        f = _ui_font(9)
        f.setStrikeOut(strikethrough)
        self.setFont(f)
        self.setFixedHeight(QFontMetrics(f).height() + 8)
        self.setToolTip(task_text)
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
        from PyQt5.QtGui import QPainter, QColor
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect()
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(P.CHIP_HOVER if self._hover else P.CHIP))
        p.drawRoundedRect(r, 5, 5)
        dot = QColor(P.DOT_PRIORITY if self._priority else P.DOT)
        if self._strikethrough:
            dot.setAlpha(110)
        p.setBrush(dot)
        cy = r.height() / 2.0
        p.drawEllipse(QPointF(8.5, cy), 3, 3)
        p.setFont(self.font())
        p.setPen(QColor(P.MUTED if self._strikethrough else P.TEXT))
        fm = QFontMetrics(self.font())
        text_rect = r.adjusted(16, 0, -6, 0)
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
            priority=bool(task.get('priority', 0))
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
        self.setObjectName('DayCell')
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)   # 내용과 상관없이 칸 크기를 똑같이
        self.setMinimumSize(0, 0)
        self.setCursor(Qt.ArrowCursor if is_other_month else Qt.PointingHandCursor)
        self.set_edges(True, True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(3)

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        self.day_label = QLabel(str(day_number))
        self.day_label.setFont(_ui_font(10, QFont.DemiBold if is_today else QFont.Medium))
        self.day_label.setAlignment(Qt.AlignCenter)
        if is_today:
            d = QFontMetrics(self.day_label.font()).height() + 6
            self.day_label.setFixedSize(d, d)
            self.day_label.setStyleSheet(f"QLabel {{ color: #FFFFFF; background: {P.TODAY}; border-radius: {d // 2}px; }}")
        else:
            if is_other_month:
                color = P.FAINT
            elif col == 0:
                color = P.SUN
            elif col == 6:
                color = P.SAT
            else:
                color = P.TEXT2
            self.day_label.setStyleSheet(f"QLabel {{ color: {color}; background: transparent; padding: 1px 3px; }}")
        top.addWidget(self.day_label, 0, Qt.AlignLeft | Qt.AlignTop)
        top.addStretch()
        layout.addLayout(top)

        self.task_layout = QVBoxLayout()
        self.task_layout.setSpacing(3)
        layout.addLayout(self.task_layout)
        layout.addStretch()

    def set_edges(self, right: bool, bottom: bool):
        """칸 사이 얇은 선 (오른쪽·아래만 그려서 선이 겹치지 않게)."""
        hover = '' if self._other else f"QFrame#DayCell:hover {{ background: {P.HOVER}; }}"
        self.setStyleSheet(
            f"QFrame#DayCell {{ background: {P.SURFACE if self._other else P.BG}; border: none;"
            f" border-right: {'1px solid ' + P.LINE if right else 'none'};"
            f" border-bottom: {'1px solid ' + P.LINE if bottom else 'none'}; }}" + hover)

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
        lbl.setFont(_ui_font(8.5, QFont.Medium))
        lbl.setStyleSheet(f"QLabel {{ color: {P.MUTED}; background: transparent; padding: 0 4px; }}"
                          f"QLabel:hover {{ color: {P.TEXT}; }}")
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
            self.task_layout.addWidget(row)
        rest = len(tasks) - len(shown)
        if rest:
            self._hidden_count = rest
            self._overflow_label = self._make_overflow_label(rest)
            self.task_layout.addWidget(self._overflow_label)
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
    edit_requested = pyqtSignal()
    delete_requested = pyqtSignal()

    def __init__(self, task, parent=None):
        super().__init__(parent)
        self.setObjectName('DetailTaskRow')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(
            f"QWidget#DetailTaskRow {{ background: transparent; border-radius: 8px; }}"
            f"QWidget#DetailTaskRow:hover {{ background: {P.HOVER}; }}"
        )
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 6, 8)
        layout.setSpacing(10)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        name_font = _ui_font(10.5, QFont.Medium)
        if task.get('strikethrough'):
            name_font.setStrikeOut(True)
        name = task['name']
        name_lbl = QLabel(name if len(name) <= 40 else name[:40] + '…')
        name_lbl.setFont(name_font)
        name_lbl.setToolTip(name)
        name_lbl.setStyleSheet(f"color: {P.MUTED if task.get('strikethrough') else P.TEXT}; background: transparent;")
        text_col.addWidget(name_lbl)
        # 점: 첫 줄 글자 높이 가운데에
        dot_wrap = QWidget()
        dot_wrap.setFixedSize(8, QFontMetrics(name_font).height())
        dot = QLabel(dot_wrap)
        dot.setFixedSize(8, 8)
        dot.move(0, (dot_wrap.height() - 8) // 2)
        dot.setStyleSheet(f"background: {P.DOT_PRIORITY if task.get('priority') else P.DOT}; border-radius: 4px;")
        layout.insertWidget(0, dot_wrap, 0, Qt.AlignTop)
        notes = task.get('notes', '').strip()
        if notes:
            first = notes.splitlines()[0]
            notes_lbl = QLabel(first if len(first) <= 50 else first[:50] + '…')
            notes_lbl.setFont(_ui_font(9))
            notes_lbl.setStyleSheet(f"color: {P.MUTED}; background: transparent;")
            text_col.addWidget(notes_lbl)
        layout.addLayout(text_col, 1)

        btn_edit = _icon_button('ph.pencil-simple', '편집', 28)
        btn_edit.clicked.connect(self.edit_requested)
        layout.addWidget(btn_edit, 0, Qt.AlignVCenter)
        btn_del = _icon_button('ph.trash', '삭제', 28)
        btn_del.clicked.connect(self.delete_requested)
        layout.addWidget(btn_del, 0, Qt.AlignVCenter)


class DayDetailPanel(QWidget):
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
        self.setStyleSheet(f"QWidget#DayDetailPanel {{ background: {P.BG}; }}")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 14, 14)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(6)
        titles = QVBoxLayout()
        titles.setSpacing(1)
        self._date_label = QLabel()
        self._date_label.setFont(_ui_font(14, QFont.Bold))
        self._date_label.setStyleSheet(f"color: {P.TEXT}; background: transparent;")
        titles.addWidget(self._date_label)
        self._count_label = QLabel()
        self._count_label.setFont(_ui_font(9))
        self._count_label.setStyleSheet(f"color: {P.MUTED}; background: transparent;")
        titles.addWidget(self._count_label)
        header.addLayout(titles, 1)
        btn_add = _primary_button('+ 할 일 추가')
        btn_add.clicked.connect(self._on_add_task)
        header.addWidget(btn_add, 0, Qt.AlignTop)
        btn_close = _icon_button('ph.x', '닫기 (Esc)', 30)
        btn_close.clicked.connect(self.close_requested)
        header.addWidget(btn_close, 0, Qt.AlignTop)
        layout.addLayout(header)

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background: {P.LINE}; border: none;")
        layout.addWidget(sep)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self._task_container = QWidget()
        self._task_container.setStyleSheet("background: transparent;")
        self._task_list_layout = QVBoxLayout(self._task_container)
        self._task_list_layout.setSpacing(2)
        self._task_list_layout.setContentsMargins(0, 0, 4, 0)
        self._task_list_layout.addStretch()
        scroll.setWidget(self._task_container)
        layout.addWidget(scroll, 1)

    def load(self, date_str, tasks):
        self.current_date_str = date_str
        try:
            d = date.fromisoformat(date_str)
            weekday = ['월', '화', '수', '목', '금', '토', '일'][d.weekday()]
            self._date_label.setText(f'{d.month}월 {d.day}일 {weekday}요일')
        except (ValueError, AttributeError):
            self._date_label.setText(date_str)
        self._count_label.setText(f'할 일 {len(tasks)}개' if tasks else '할 일 없음')

        while self._task_list_layout.count() > 1:
            item = self._task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not tasks:
            empty = QLabel('이 날은 비어 있어요.\n[+ 할 일 추가]로 넣을 수 있어요.')
            empty.setAlignment(Qt.AlignCenter)
            empty.setFont(_ui_font(10))
            empty.setStyleSheet(f"color: {P.MUTED}; background: transparent; padding: 30px 0;")
            self._task_list_layout.insertWidget(0, empty)
            return

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
        from db import delete_task as _del
        _del(task['id'])
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
        from db import get_tasks_by_date
        self._panel.load(self._date_str, get_tasks_by_date(self._date_str))

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


class NativeCalendarDialog(QDialog):
    def __init__(self, open_windows, parent=None):
        super().__init__(parent)
        # 주인 없이 열면 작업 표시줄에 따로 뜨는 보통 창 (다른 창 뒤로 갈 수 있음)
        self.setWindowFlags((Qt.Dialog if parent is not None else Qt.Window)
                            | Qt.WindowCloseButtonHint | Qt.WindowMinMaxButtonsHint)
        self.setWindowTitle('달력')
        self._open_windows = open_windows
        self._current_date = date.today()
        self._task_data = {}
        self.setMinimumSize(px(760), px(560))
        self._build_ui()
        self._load_tasks()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(12)
        self.setStyleSheet(f"NativeCalendarDialog {{ background: {P.BG}; }}")

        # ── 머리말: '11월 2026' | 오늘 ‹ › ──
        header = QHBoxLayout()
        header.setSpacing(6)
        self._month_title = QLabel()
        self._month_title.setTextFormat(Qt.RichText)
        self._month_title.setStyleSheet("background: transparent;")
        header.addWidget(self._month_title)
        header.addStretch(1)
        btn_today = _ghost_button('오늘')
        btn_today.setToolTip('이번 달로')
        btn_today.clicked.connect(self._go_today)
        header.addWidget(btn_today)
        header.addSpacing(4)
        btn_prev = _icon_button('ph.caret-left', '지난 달')
        btn_prev.clicked.connect(self._prev_month)
        header.addWidget(btn_prev)
        btn_next = _icon_button('ph.caret-right', '다음 달')
        btn_next.clicked.connect(self._next_month)
        header.addWidget(btn_next)
        main_layout.addLayout(header)

        # ── 요일 + 날짜 칸 (얇은 테두리의 둥근 표) ──
        board = QFrame()
        board.setObjectName('CalBoard')
        board.setStyleSheet(f"QFrame#CalBoard {{ background: {P.BG}; border: 1px solid {P.LINE}; border-radius: 10px; }}")
        grid_vbox = QVBoxLayout(board)
        grid_vbox.setContentsMargins(1, 1, 1, 1)
        grid_vbox.setSpacing(0)
        day_header = QHBoxLayout()
        day_header.setContentsMargins(0, 0, 0, 0)
        day_header.setSpacing(0)
        for i, name in enumerate(['일', '월', '화', '수', '목', '금', '토']):
            lbl = QLabel(name)
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            lbl.setFont(_ui_font(9, QFont.DemiBold))
            color = P.SUN if i == 0 else (P.SAT if i == 6 else P.MUTED)
            lbl.setStyleSheet(f"QLabel {{ color: {color}; background: transparent; padding: 8px 10px;"
                              f" border-bottom: 1px solid {P.LINE}; }}")
            day_header.addWidget(lbl, 1)
        grid_vbox.addLayout(day_header)
        self._grid = QGridLayout()
        self._grid.setSpacing(0)
        self._grid.setContentsMargins(0, 0, 0, 0)
        grid_vbox.addLayout(self._grid, 1)
        main_layout.addWidget(board, 1)

        foot = QLabel('날짜를 누르면 그날 할 일을 보고, 추가·편집할 수 있어요.')
        foot.setFont(_ui_font(8.5))
        foot.setStyleSheet(f"color: {P.FAINT}; background: transparent;")
        main_layout.addWidget(foot)

    def _load_tasks(self):
        from db import get_all_tasks_for_calendar, get_tasks_by_date
        counts = get_all_tasks_for_calendar()
        self._task_data = {}
        for date_str in counts:
            self._task_data[date_str] = get_tasks_by_date(date_str)
        self._render_calendar()
        if getattr(self, '_detail_dlg', None) is not None and self._detail_dlg.isVisible():
            self._detail_dlg.reload()

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and self.isActiveWindow():
            self._load_tasks()

    def _render_calendar(self):
        from datetime import timedelta
        y, m = self._current_date.year, self._current_date.month
        self._month_title.setText(
            f"<span style='font-size:20pt; font-weight:700; color:{P.TEXT};'>{m}월</span>"
            f"<span style='font-size:20pt; font-weight:400; color:{P.MUTED};'>&nbsp;{y}</span>")

        while self._grid.count():
            item = self._grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        first_day = date(y, m, 1)
        start_weekday = (first_day.weekday() + 1) % 7
        if m == 12:
            last_day = date(y + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = date(y, m + 1, 1) - timedelta(days=1)

        today_str = date.today().strftime('%Y-%m-%d')

        if m == 1:
            prev_last = date(y - 1, 12, 1)
        else:
            prev_last = date(y, m, 1) - timedelta(days=1)
        prev_days = prev_last.day
        for i in range(start_weekday):
            day = prev_days - start_weekday + i + 1
            if m == 1:
                d_str = date(y - 1, 12, day).strftime('%Y-%m-%d')
            else:
                d_str = date(y, m - 1, day).strftime('%Y-%m-%d')
            cell = DayCellWidget(d_str, day, is_today=(d_str == today_str), col=i, is_other_month=True)
            cell.setCursor(Qt.ArrowCursor)
            self._add_cell_tasks(cell, d_str, connect_signals=False)
            self._grid.addWidget(cell, 0, i)

        row, col = 0, start_weekday
        for day in range(1, last_day.day + 1):
            d_str = date(y, m, day).strftime('%Y-%m-%d')
            cell = DayCellWidget(d_str, day, is_today=(d_str == today_str), col=col)
            self._add_cell_tasks(cell, d_str, connect_signals=True)
            self._grid.addWidget(cell, row, col)
            col += 1
            if col == 7:
                col = 0
                row += 1

        next_day = 1
        while row < 6:
            nm = m + 1 if m < 12 else 1
            ny = y if m < 12 else y + 1
            d_str = date(ny, nm, next_day).strftime('%Y-%m-%d')
            cell = DayCellWidget(d_str, next_day, is_today=(d_str == today_str), col=col, is_other_month=True)
            cell.setCursor(Qt.ArrowCursor)
            self._add_cell_tasks(cell, d_str, connect_signals=False)
            self._grid.addWidget(cell, row, col)
            next_day += 1
            col += 1
            if col == 7:
                col = 0
                row += 1
            if col == 0 and row == 6:
                break

        for r in range(6):
            self._grid.setRowStretch(r, 1)
        for c in range(7):
            self._grid.setColumnStretch(c, 1)
        # 칸 사이 선: 마지막 열·줄은 바깥 테두리와 겹치지 않게 빼기
        for i in range(self._grid.count()):
            w = self._grid.itemAt(i).widget()
            if isinstance(w, DayCellWidget):
                r, c, _, _ = self._grid.getItemPosition(i)
                w.set_edges(right=c < 6, bottom=r < 5)

    def _add_cell_tasks(self, cell, date_str, connect_signals=True):
        tasks = self._task_data.get(date_str, [])
        for task in tasks:
            cell.add_task(task)
        if connect_signals:
            cell.clicked.connect(self._on_cell_clicked)
            cell.overflow_clicked.connect(self._on_cell_clicked)

    def _on_cell_clicked(self, date_str):
        if getattr(self, '_detail_dlg', None) is not None and self._detail_dlg.isVisible():
            if self._detail_dlg._date_str == date_str:
                self._detail_dlg.raise_()
                self._detail_dlg.activateWindow()
                return
            self._detail_dlg.close()
        tasks = self._task_data.get(date_str, [])
        self._detail_dlg = DayDetailDialog(self._open_windows, date_str, tasks, parent=self)
        self._detail_dlg.task_edited.connect(self._on_task_changed)
        self._detail_dlg.task_deleted.connect(self._on_task_changed)
        self._detail_dlg.show()
        self._detail_dlg.raise_()

    def _on_task_changed(self):
        self._load_tasks()
        for win in self._open_windows:
            win._refresh_tasks()
        if getattr(self, '_detail_dlg', None) is not None and self._detail_dlg.isVisible():
            self._detail_dlg.reload()

    def _prev_month(self):
        y, m = self._current_date.year, self._current_date.month
        if m == 1: self._current_date = date(y - 1, 12, 1)
        else: self._current_date = date(y, m - 1, 1)
        self._load_tasks()

    def _next_month(self):
        y, m = self._current_date.year, self._current_date.month
        if m == 12: self._current_date = date(y + 1, 1, 1)
        else: self._current_date = date(y, m + 1, 1)
        self._load_tasks()

    def _go_today(self):
        self._current_date = date.today()
        self._load_tasks()
