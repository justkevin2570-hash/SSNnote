"""ui/calendar_view.py — 월간 달력 / 날짜별 상세 / 업무 편집 대화상자."""
import calendar
from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QMenu, QMessageBox, QDialog, QGridLayout, QPlainTextEdit, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate, QEvent, QSettings, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QCursor, QMouseEvent
from db import (add_task, delete_task, add_task_history, delete_task_history,
                get_all_tasks, get_all_task_history)
from .common import mi_icon, pr_font
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

class TaskButton(QPushButton):
    def __init__(self, task_id, task_text, strikethrough=False, priority=False, parent=None):
        super().__init__(parent)
        self.task_id = task_id
        self.full_text = task_text
        self._strikethrough = strikethrough
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        base = f"QPushButton {{ text-align: left; padding: 3px 6px; background: {tc('rgba(255,255,255,0.5)')}; border: 1px solid {T.SURFACE_WARM}; border-radius: 3px; font-size: 15px; font-family: Malgun Gothic; color: {tct('#3a3028')}; }}"
        hover = f"QPushButton:hover {{ background: {tc('rgba(212,184,0,0.15)')}; border-color: {T.ACCENT}; }}"
        self.setStyleSheet(base + hover)

    def resizeEvent(self, event):
        metrics = QFontMetrics(self.font())
        available = max(self.width() - 10, 20)
        elided = metrics.elidedText(self.full_text, Qt.ElideRight, available)
        self.setText(elided)
        super().resizeEvent(event)


class TaskIndicatorRow(QWidget):
    """세로 bar + TaskButton 컨테이너"""
    def __init__(self, task, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(3)

        bar = QFrame()
        bar.setFixedSize(3, 16)
        color = T.HIGHLIGHT if task.get('priority') else tc('#a98fd4')
        bar.setStyleSheet(f"QFrame {{ background: {color}; border-radius: 1px; }}")
        layout.addWidget(bar)

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
        self.setFrameShape(QFrame.Box)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(1)

        self.day_label = QLabel(str(day_number))
        self.day_label.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.day_label.setFont(QFont('Malgun Gothic', 12, QFont.Bold))
        if is_today:
            self.day_label.setStyleSheet(f"QLabel {{ color: {tct('#fff')}; background: {T.HIGHLIGHT}; border-radius: 12px; min-width: 24px; min-height: 24px; }}")
            self.setStyleSheet(f"DayCellWidget {{ background: {T.HIGHLIGHT_BG}; border: 1px solid {T.BORDER_WARM}; }}")
        elif is_other_month:
            if col == 0:
                self.day_label.setStyleSheet(f"QLabel {{ color: {tct('#e8a09a')}; }}")
            elif col == 6:
                self.day_label.setStyleSheet(f"QLabel {{ color: {tct('#9ab8e8')}; }}")
            else:
                self.day_label.setStyleSheet(f"QLabel {{ color: {tct('#c0b8b0')}; }}")
            self.setStyleSheet(f"DayCellWidget {{ background: {tc('#fafafa')}; border: 1px solid {T.BORDER_WARM}; }}")
        elif col == 0:
            self.day_label.setStyleSheet(f"QLabel {{ color: {T.HIGHLIGHT}; }}")
            self.setStyleSheet(f"DayCellWidget {{ background: {tc('white')}; border: 1px solid {T.BORDER_WARM}; }}")
        elif col == 6:
            self.day_label.setStyleSheet(f"QLabel {{ color: {T.SATURDAY_SOFT}; }}")
            self.setStyleSheet(f"DayCellWidget {{ background: {tc('white')}; border: 1px solid {T.BORDER_WARM}; }}")
        else:
            self.day_label.setStyleSheet(f"QLabel {{ color: {tct('#3a3028')}; }}")
            self.setStyleSheet(f"DayCellWidget {{ background: {tc('white')}; border: 1px solid {T.BORDER_WARM}; }}")
        layout.addWidget(self.day_label)

        self.task_layout = QVBoxLayout()
        self.task_layout.setSpacing(2)
        layout.addLayout(self.task_layout)
        layout.addStretch()

    def clear_tasks(self):
        while self.task_layout.count():
            item = self.task_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._hidden_count = 0
        self._overflow_label = None

    def _visible_btn_count(self):
        return sum(
            1 for i in range(self.task_layout.count())
            if isinstance(self.task_layout.itemAt(i).widget(), TaskIndicatorRow)
        )

    def _update_overflow_label(self):
        self._hidden_count += 1
        if self._overflow_label is not None:
            self._overflow_label.setText(f'+{self._hidden_count} 더보기')
            return
        lbl = QLabel(f'+{self._hidden_count} 더보기')
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setCursor(Qt.PointingHandCursor)
        lbl.setFont(QFont('Malgun Gothic', 9, QFont.Bold))
        lbl.setStyleSheet(
            f"QLabel {{ color: {tct('#888')}; background: {tc('rgba(212,184,0,0.12)')}; "
            f"border: 1px dashed {T.ACCENT}; border-radius: 3px; padding: 1px 4px; }}"
        )
        lbl.mousePressEvent = lambda e: self.overflow_clicked.emit(self.date_str)
        self._overflow_label = lbl
        self.task_layout.addWidget(lbl)

    def add_task(self, task):
        if self._visible_btn_count() < self.MAX_VISIBLE:
            row = TaskIndicatorRow(task)
            row.button().clicked.connect(lambda checked: self.clicked.emit(self.date_str))
            self.task_layout.addWidget(row)
        else:
            self._update_overflow_label()

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
        self.setStyleSheet(
            f"QWidget#DetailTaskRow {{ background: {tc('white')}; border: 1px solid {T.BORDER_WARM}; border-radius: 6px; }}"
        )
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(6)

        bar = QFrame()
        bar.setFixedSize(4, 18)
        color = T.HIGHLIGHT if task.get('priority') else tc('#a98fd4')
        bar.setStyleSheet(f"QFrame {{ background: {color}; border-radius: 2px; }}")
        layout.addWidget(bar)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        name_font = QFont('Malgun Gothic', 11)
        if task.get('strikethrough'):
            name_font.setStrikeOut(True)
        name = task['name']
        name_lbl = QLabel(name if len(name) <= 30 else name[:30] + '…')
        name_lbl.setFont(name_font)
        name_lbl.setStyleSheet(f"color: {T.INK}; background: transparent;")
        text_col.addWidget(name_lbl)
        notes = task.get('notes', '').strip()
        if notes:
            notes_lbl = QLabel(notes if len(notes) <= 40 else notes[:40] + '…')
            notes_lbl.setFont(QFont('Malgun Gothic', 9))
            notes_lbl.setStyleSheet(
                f"color: {T.TEXT_WARM_MUTED}; background: {T.BG_WARM}; border-radius: 3px; padding: 2px 4px;"
            )
            text_col.addWidget(notes_lbl)
        layout.addLayout(text_col, 1)

        btn_edit = QPushButton('편집')
        btn_edit.setFixedSize(48, 28)
        btn_edit.setFont(QFont('Malgun Gothic', 9))
        btn_edit.setStyleSheet(
            f"QPushButton {{ border: 1px solid {T.ACCENT}; border-radius: 4px; "
            f"background: transparent; color: {T.TEXT_WARM}; }} "
            f"QPushButton:hover {{ background: {tc('rgba(212,184,0,0.2)')}; }}"
        )
        btn_edit.clicked.connect(self.edit_requested)
        layout.addWidget(btn_edit)

        btn_del = QPushButton('삭제')
        btn_del.setFixedSize(48, 28)
        btn_del.setFont(QFont('Malgun Gothic', 9))
        btn_del.setStyleSheet(
            f"QPushButton {{ border: 1px solid {T.HIGHLIGHT}; border-radius: 4px; "
            f"background: transparent; color: {T.HIGHLIGHT}; }} "
            f"QPushButton:hover {{ background: {tc('rgba(232,93,58,0.1)')}; }}"
        )
        btn_del.clicked.connect(self.delete_requested)
        layout.addWidget(btn_del)


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
        self.setStyleSheet(
            f"QWidget#DayDetailPanel {{ background: {T.BG_WARM}; }}"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        self._date_label = QLabel()
        self._date_label.setFont(QFont('Malgun Gothic', 13, QFont.Bold))
        self._date_label.setStyleSheet(f"color: {T.INK}; background: transparent;")
        header.addWidget(self._date_label, 1)
        btn_add = QPushButton('＋ 일정 추가')
        btn_add.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        btn_add.setStyleSheet(
            f"QPushButton {{ background: {T.ACCENT}; color: {tct('white')}; border: none; "
            "border-radius: 6px; padding: 5px 12px; } "
            f"QPushButton:hover {{ background: {T.ACCENT_HOVER}; }}"
        )
        btn_add.clicked.connect(self._on_add_task)
        header.addWidget(btn_add)
        btn_close = QPushButton('✕')
        btn_close.setFixedSize(28, 28)
        btn_close.setStyleSheet(
            f"QPushButton {{ border: none; font-size: 13px; color: {tct('#aaa')}; background: transparent; }} "
            f"QPushButton:hover {{ color: {tct('#555')}; }}"
        )
        btn_close.clicked.connect(self.close_requested)
        header.addWidget(btn_close)
        layout.addLayout(header)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"border: none; border-top: 1px solid {T.SURFACE_WARM}; background: transparent;")
        sep.setFixedHeight(1)
        layout.addWidget(sep)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self._task_container = QWidget()
        self._task_container.setStyleSheet("background: transparent;")
        self._task_list_layout = QVBoxLayout(self._task_container)
        self._task_list_layout.setSpacing(3)
        self._task_list_layout.setContentsMargins(0, 0, 4, 0)
        self._task_list_layout.addStretch()
        scroll.setWidget(self._task_container)
        layout.addWidget(scroll, 1)


    def load(self, date_str, tasks):
        self.current_date_str = date_str
        try:
            d = date.fromisoformat(date_str)
            weekday = ['월', '화', '수', '목', '금', '토', '일'][d.weekday()]
            self._date_label.setText(f'{d.month}월 {d.day}일 ({weekday})')
        except (ValueError, AttributeError):
            self._date_label.setText(date_str)

        while self._task_list_layout.count() > 1:
            item = self._task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not tasks:
            empty = QLabel('등록된 할 일이 없습니다.')
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(f"color: {tct('#aaa')}; font-size: 11pt; background: transparent;")
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
        self.setStyleSheet(
            f"QDialog {{ background: {tc('white')}; }} "
            "QLabel { background: transparent; }"
        )
        self._build_ui(default_date)

    def _build_ui(self, default_date):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText('할 일 이름')
        self._name_edit.setFont(QFont('Malgun Gothic', 11))
        self._name_edit.setStyleSheet(
            f"QLineEdit {{ border: 1px solid {T.ACCENT}; border-radius: 6px; padding: 6px 10px; background: {tc('#fffef5')}; }}"
        )
        if self._task:
            self._name_edit.setText(self._task['name'])
        layout.addWidget(self._name_edit)

        date_row = QHBoxLayout()
        date_lbl = QLabel('마감일:')
        date_lbl.setFont(QFont('Malgun Gothic', 10))
        date_lbl.setStyleSheet(f"color: {T.TEXT_WARM};")
        date_row.addWidget(date_lbl)
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setFont(QFont('Malgun Gothic', 10))
        self._date_edit.setDisplayFormat('yyyy-MM-dd')
        self._date_edit.setStyleSheet(
            f"QDateEdit {{ border: 1px solid {tc('#ccc')}; border-radius: 6px; padding: 4px 8px; background: {tc('white')}; }}"
        )
        if self._task and self._task.get('deadline'):
            self._date_edit.setDate(QDate.fromString(self._task['deadline'][:10], 'yyyy-MM-dd'))
        elif default_date:
            self._date_edit.setDate(QDate.fromString(default_date, 'yyyy-MM-dd'))
        else:
            self._date_edit.setDate(QDate.currentDate())
        date_row.addWidget(self._date_edit, 1)
        layout.addLayout(date_row)

        memo_lbl = QLabel('메모')
        memo_lbl.setFont(QFont('Malgun Gothic', 9))
        memo_lbl.setStyleSheet(f"color: {tct('#888')};")
        layout.addWidget(memo_lbl)
        self._memo_edit = QPlainTextEdit()
        self._memo_edit.setPlaceholderText('메모를 입력하세요...')
        self._memo_edit.setFont(QFont('Malgun Gothic', 10))
        self._memo_edit.setFixedHeight(80)
        self._memo_edit.setStyleSheet(
            f"QPlainTextEdit {{ border: 1px solid {tc('#ccc')}; border-radius: 6px; padding: 6px; background: {tc('#fafafa')}; }}"
        )
        if self._task:
            from db import get_task_notes
            self._memo_edit.setPlainText(get_task_notes(self._task['id']))
        layout.addWidget(self._memo_edit)

        btn_row = QHBoxLayout()
        btn_cancel = QPushButton('취소')
        btn_cancel.setFont(QFont('Malgun Gothic', 10))
        btn_cancel.setStyleSheet(
            f"QPushButton {{ border: 1px solid {tc('#ccc')}; border-radius: 6px; padding: 6px 16px; background: transparent; color: {T.TEXT_WARM}; }}"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton('저장')
        btn_ok.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        btn_ok.setDefault(True)
        btn_ok.setStyleSheet(
            f"QPushButton {{ background: {T.ACCENT}; color: {tct('white')}; border: none; "
            "border-radius: 6px; padding: 6px 20px; } "
            f"QPushButton:hover {{ background: {T.ACCENT_HOVER}; }}"
        )
        btn_ok.clicked.connect(self._on_save)
        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        self.setFixedWidth(360)

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
        self.setMinimumSize(320, 400)
        self.setStyleSheet(f"QDialog {{ background: {T.BG_WARM}; }}")

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
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint | Qt.WindowMinMaxButtonsHint)
        self.setWindowTitle('월별 일정 보기')
        self._open_windows = open_windows
        self._current_date = date.today()
        self._task_data = {}
        self.setMinimumSize(900, 650)
        self._build_ui()
        self._load_tasks()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(6)
        self.setStyleSheet(f"NativeCalendarDialog {{ background: {tc('#f5f0ea')}; }}")

        header = QHBoxLayout()
        header.addStretch(1)
        btn_prev = QPushButton('◀')
        btn_prev.setFixedSize(40, 40)
        btn_prev.setStyleSheet(f"QPushButton {{ background: {T.BG_WARM_ALT}; border: none; border-radius: 50%; font-size: 16px; color: {T.TEXT_WARM}; }} QPushButton:hover {{ background: {T.SURFACE_WARM}; }}")
        btn_prev.clicked.connect(self._prev_month)
        header.addWidget(btn_prev)

        self._month_title = QLabel()
        self._month_title.setAlignment(Qt.AlignCenter)
        self._month_title.setFont(QFont('Malgun Gothic', 18, QFont.Bold))
        self._month_title.setStyleSheet(f"color: {T.INK};")
        self._month_title.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        header.addWidget(self._month_title)

        btn_next = QPushButton('▶')
        btn_next.setFixedSize(40, 40)
        btn_next.setStyleSheet(f"QPushButton {{ background: {T.BG_WARM_ALT}; border: none; border-radius: 50%; font-size: 14px; color: {T.TEXT_WARM}; }} QPushButton:hover {{ background: {T.SURFACE_WARM}; }}")
        btn_next.clicked.connect(self._next_month)
        header.addWidget(btn_next)
        header.addStretch(1)

        btn_today = QPushButton('오늘')
        _today_font = pr_font(16)
        _today_font.setBold(True)
        btn_today.setFont(_today_font)
        btn_today.setStyleSheet(f"QPushButton {{ background: {T.BG_WARM_ALT}; border: none; border-radius: 8px; padding: 6px 14px; color: {T.TEXT_WARM}; }} QPushButton:hover {{ background: {T.SURFACE_WARM}; }}")
        btn_today.clicked.connect(self._go_today)
        header.addWidget(btn_today)
        main_layout.addLayout(header)

        day_header = QHBoxLayout()
        day_header.setSpacing(2)
        for i, name in enumerate(['일', '월', '화', '수', '목', '금', '토']):
            lbl = QLabel(name)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setFont(QFont('Malgun Gothic', 11, QFont.Bold))
            if i == 0: lbl.setStyleSheet(f"color: {T.HIGHLIGHT};")
            elif i == 6: lbl.setStyleSheet(f"color: {T.SATURDAY_SOFT};")
            else: lbl.setStyleSheet(f"color: {tct('#888')};")
            day_header.addWidget(lbl)

        grid_container = QWidget()
        grid_container.setStyleSheet("background: transparent;")
        grid_vbox = QVBoxLayout(grid_container)
        grid_vbox.setContentsMargins(0, 0, 0, 0)
        grid_vbox.setSpacing(4)
        grid_vbox.addLayout(day_header)
        self._grid = QGridLayout()
        self._grid.setSpacing(2)
        grid_vbox.addLayout(self._grid, 1)
        main_layout.addWidget(grid_container, 1)

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
        self._month_title.setText(f'{y}년 {m}월')

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
