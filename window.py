import os
import ctypes
from datetime import date, timedelta
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QDateTimeEdit, QMenu, QAction, QWidgetAction,
    QMessageBox, QDialog, QGridLayout, QCalendarWidget, QToolButton,
    QPlainTextEdit, QTextEdit, QGraphicsColorizeEffect, QTimeEdit,
    QComboBox, QScrollBar
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QDateTime, QPoint, QPointF, QSettings, QUrl
from PyQt5.QtGui import QFont, QColor, QTextCharFormat, QTextBlockFormat, QTextCursor
from db import (update_window, delete_window, get_tasks, add_task, delete_task, update_task,
                add_task_history, get_task_history, delete_task_history,
                get_documents, add_document, update_document, delete_document,
                get_official_documents, delete_official_document,
                get_window_memo, set_window_memo, get_all_windows,
                set_window_merge_group, clear_window_merge_group, get_next_merge_group_id)
from autostart import is_enabled as autostart_is_enabled, set_enabled as autostart_set
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, _normalize_doc_number
import updater
from ui.common import _base_path, _load_pretendard, pr_font, SNAP_THRESHOLD, SNAP_ZONE, TITLE_BAR_HEIGHT
from ui.dates import _next_recurrence_deadline, calc_dday
from ui.widgets import _SelectableLabel, AddDateButton, CustomCalendarWidget, DdayLabel, PlainPasteMemoEdit
from ui.dialogs import CustomRecurrenceDialog
from ui.chrome import EdgeHandle, MergeButton, TitleBar
from ui.task_row import _NoteButton, TaskRow
from ui.document_row import DocumentRow
from ui.calendar_view import MonthlyCalendarDialog, NativeCalendarDialog
from ui.theme import T, PALETTE, tc, tct, note_color, DARK

class MemoWindow(QMainWindow):
    def __init__(self, window_id, on_new=None, open_windows=None, on_toggle_hotkey=None,
                 on_alarm_interval_change=None, get_alarm_interval=None,
                 on_timed_alarm_change=None, get_timed_alarm_enabled=None,
                 on_shortcut_change=None, get_shortcut_enabled=None,
                 force_memo_mode=False):
        super().__init__()
        self.window_id       = window_id
        self.on_new          = on_new or (lambda **kw: None)
        self._open_windows   = open_windows if open_windows is not None else []
        self._force_memo_mode = force_memo_mode
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
        self._bg_color       = tc('#FEFFA7')  # 기본값
        self._capture_hint_shown = False
        self._scale          = 1.0
        self._row_gap        = 0
        self._expanded_note_ids = set()
        self._expanded_attach_ids = set()
        self._selected_task_id = None
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
        self._palette_shortcut = QShortcut(QKeySequence('Ctrl+K'), self)
        self._palette_shortcut.activated.connect(self._open_command_palette)
        self._shortcuts = [self._pin_shortcut, self._shade_shortcut, self._palette_shortcut]
        # 저장된 단축키 상태 적용
        if get_shortcut_enabled is not None and not get_shortcut_enabled():
            for sc in self._shortcuts:
                sc.setEnabled(False)

    def apply_state(self, x, y, width, height, collapsed, color='', scale=1.0, row_gap=0):
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
        if row_gap:
            self._row_gap = row_gap
            self._apply_row_gap()

    def _build_ui(self):
        central = QWidget()
        central.setObjectName('central')
        central.setStyleSheet(f"""
            QWidget#central {{
                background: {T.INPUT_BG};
                border-radius: 0;
                border: 1px solid {tc("rgba(0,0,0,0.4)")};
            }}
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
        self._title_sep.setStyleSheet(f'background: {T.DIVIDER}; border: none;')
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
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(4)

        field_style_line = f"""
            QLineEdit {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {tct("#333")};
            }}
        """
        _date_base = """
            QDateEdit::drop-down { width: 0px; border: none; background: transparent; }
            QDateEdit::up-button { width: 0px; }
            QDateEdit::down-button { width: 0px; }
        """
        self._field_style_date_empty = f"""
            QDateEdit {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {tct("#aaa")};
                qproperty-alignment: AlignLeft;
            }}
        """ + _date_base
        self._field_style_date_filled = f"""
            QDateEdit {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {tct("#333")};
                qproperty-alignment: AlignCenter;
            }}
        """ + _date_base
        field_style_date = self._field_style_date_empty

        # 1행: 업무
        self.task_input_row = QWidget()
        self.task_input_row.setStyleSheet('background: transparent;')
        row1 = QHBoxLayout(self.task_input_row)
        row1.setContentsMargins(16, 0, 8, 0)
        row1.setSpacing(4)
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText('업무명')
        self.input_name.setFont(QFont('Malgun Gothic', 11))
        self.input_name.setStyleSheet(field_style_line)
        self._date_touched = False
        self._time_touched = False
        self.input_name.installEventFilter(self)
        self.input_name.returnPressed.connect(self._add_task)
        self.input_name.textChanged.connect(self._on_name_cleared)
        row1.addWidget(self.input_name, 1)
        content_layout.addWidget(self.task_input_row)

        # 2행: 마감일 (기본 숨김)
        self.date_row_widget = QWidget()
        self.date_row_widget.setStyleSheet('background: transparent;')
        row2 = QHBoxLayout(self.date_row_widget)
        row2.setContentsMargins(16, 0, 8, 0)
        row2.setSpacing(4)
        self.input_date = QDateEdit()
        self.input_date.setDisplayFormat('yyyy-MM-dd')
        self.input_date.setCalendarPopup(False)
        self.input_date.setMinimumDate(QDate(1900, 1, 1))
        self.input_date.setSpecialValueText('마감일')
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
        self.input_time.setStyleSheet(f"""
            QTimeEdit {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {tct("#333")};
                qproperty-alignment: AlignCenter;
            }}
            QTimeEdit:disabled {{
                background: {tc("rgba(255,255,255,0.2)")};
                border: 1px solid {tc("#ddd")};
                color: {tct("#bbb")};
            }}
            QTimeEdit::up-button {{ width: 0px; }}
            QTimeEdit::down-button {{ width: 0px; }}
        """)
        self.input_time.timeChanged.connect(lambda: setattr(self, '_time_touched', True))
        self.input_time.installEventFilter(self)
        for child in self.input_time.findChildren(QWidget):
            child.installEventFilter(self)
        # 입력 행 자체에도 필터 — 빈 공간/비활성(시간) 칸 클릭도 선택 해제 대상
        self.task_input_row.installEventFilter(self)
        self.date_row_widget.installEventFilter(self)
        # 본문 빈 공간(행 밖 목록/구분선/여백) 클릭 → 선택 해제용 필터
        self._top_panel = _top_panel
        _top_panel.installEventFilter(self)
        self._cal_popup = CustomCalendarWidget()
        self._cal_popup.setWindowFlags(Qt.Popup)
        self._cal_popup.setStyleSheet(f"""
            QCalendarWidget {{
                background-color: {tc("white")};
            }}
            QCalendarWidget QAbstractItemView {{
                background-color: {tc("white")};
                color: black;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                selection-background-color: {T.ACCENT};
                selection-color: {tc("white")};
            }}
            QCalendarWidget QWidget {{
                background-color: {tc("white")};
            }}
            QCalendarWidget QToolButton {{
                background-color: {tc("white")};
                color: black;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                padding: 2px 4px;
                min-width: 40px;
            }}
            QCalendarWidget QWidget#qt_calendar_navigationbar {{
                background-color: {tc("white")};
            }}
            QCalendarWidget QSpinBox {{
                background-color: {tc("white")};
                color: black;
            }}
        """)
        fmt_sat = QTextCharFormat()
        fmt_sat.setForeground(QColor(T.SATURDAY))
        self._cal_popup.setWeekdayTextFormat(Qt.Saturday, fmt_sat)
        fmt_sun = QTextCharFormat()
        fmt_sun.setForeground(QColor(T.SUNDAY))
        self._cal_popup.setWeekdayTextFormat(Qt.Sunday, fmt_sun)
        self._cal_popup.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self._cal_popup.clicked.connect(self._on_cal_date_selected)
        btn_history = QPushButton('지난 기록')
        btn_history.setFont(QFont('Malgun Gothic', 10))
        btn_history.setStyleSheet(f"""
            QPushButton {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: {tc("rgba(212,184,0,0.25)")}; }}
        """)
        btn_history.clicked.connect(self._show_history)
        row1.addWidget(btn_history)

        self.input_recurrence = QComboBox()
        self.input_recurrence.addItems(['반복 없음', '매주', '격주', '매월', '매년', '사용자 설정...'])
        for _i in range(self.input_recurrence.count()):
            self.input_recurrence.setItemData(_i, Qt.AlignLeft, Qt.TextAlignmentRole)
        self.input_recurrence.installEventFilter(self)
        self.input_recurrence.setFixedWidth(90)
        self.input_recurrence.setFont(QFont('Malgun Gothic', 11))
        self._recurrence_style_base = (
            f"QComboBox {{{{ border: 1px solid {T.ACCENT}; border-radius: 4px; padding: 0 6px; "
            f"background: {tc('#fffde7')}; color: {{color}}; }}}}"
            f"QComboBox:focus {{{{ border: 2px solid {tc('#4a90d9')}; background: {tc('rgba(255,255,255,0.85)')}; }}}}"
            "QComboBox::drop-down {{ border: none; width: 0px; }}"
            "QComboBox QAbstractItemView {{ font-family: 'Malgun Gothic'; font-size: 11pt; }}"
        )
        self.input_recurrence.setStyleSheet(self._recurrence_style_base.format(color=tc('#aaa')))
        self._custom_recurrence_val = ''
        self.input_recurrence.currentIndexChanged.connect(self._on_recurrence_index_changed)

        _row_h = self.input_name.sizeHint().height()
        self.input_date.setFixedHeight(_row_h)
        self.input_time.setFixedHeight(_row_h)
        self.input_recurrence.setFixedHeight(_row_h)

        row2.addWidget(self.input_date)
        row2.addWidget(self.input_time)
        row2.addWidget(self.input_recurrence)
        row2.addStretch()
        content_layout.addWidget(self.date_row_widget)

        self._sep_date = QFrame()
        self._sep_date.setFrameShape(QFrame.HLine)
        self._sep_date.setFixedHeight(1)
        self._sep_date.setStyleSheet(f'background: {T.DIVIDER}; border: none;')
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
        self.task_scroll.setStyleSheet(f"""
            QScrollArea {{ border: none; background: transparent; }}
            QScrollBar:vertical {{ width: 6px; background: transparent; }}
            QScrollBar::handle:vertical {{ background: {T.ACCENT}; border-radius: 3px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
        content_layout.addWidget(self.task_scroll)

        # 메모 모드 자유 편집기 (기본 숨김)
        self.memo_editor = PlainPasteMemoEdit()
        _memo_font = QFont('굴림체', 12)
        _memo_font.setStretch(100)
        _memo_font.setLetterSpacing(QFont.PercentageSpacing, 100.0)
        self.memo_editor.setFont(_memo_font)
        self.memo_editor.document().blockCountChanged.connect(self._apply_memo_line_spacing)
        self.memo_editor.setStyleSheet(f"""
            QTextEdit {{
                background: transparent;
                border: none;
                color: {tct("#333")};
                padding: 4px 8px;
            }}
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

        _btn_style = f"""
            QPushButton {{
                background: {tc("rgba(255,255,255,0.5)")};
                border: 1px solid {T.ACCENT};
                border-radius: 6px;
                padding: 4px 8px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: {tc("rgba(212,184,0,0.25)")}; }}
        """

        btn_at_a_glance = QPushButton('달력')
        btn_at_a_glance.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        btn_at_a_glance.setStyleSheet(f"""
            QPushButton {{
                background: {T.TEXT_ON_NOTE};
                border: 1px solid {T.ACCENT};
                border-radius: 6px;
                padding: 4px 8px;
                color: {tct("#ffffff")};
            }}
            QPushButton:hover {{ background: {tc("#6b5410")}; }}
        """)
        btn_at_a_glance.clicked.connect(self._show_monthly_calendar)
        bottom_layout.addWidget(btn_at_a_glance)

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

        self._sep_bottom = QFrame()
        self._sep_bottom.setFrameShape(QFrame.HLine)
        self._sep_bottom.setFixedHeight(1)
        self._sep_bottom.setStyleSheet(f'background: {T.DIVIDER}; border: none;')
        _splitter_layout.addWidget(self._sep_bottom)
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
        self._apply_memo_line_spacing()
        if _saved_mode or self._force_memo_mode:
            self._apply_memo_mode(True)
            if self._force_memo_mode and not _saved_mode:
                self._flush_memo_save()
        if self._force_memo_mode:
            self.title_bar.btn_mode.hide()
            self.title_bar.btn_budget.hide()

    def _is_empty_area_click(self, obj, event=None):
        """클릭 대상이 '아무 효과 없는 빈 공간'인지 판정.

        이벤트 필터는 자식 클릭을 먼저 받으므로, 배제 판정은 '클릭 대상 자체'에만
        적용하고 부모 체인은 _top_panel 도달 여부만 확인한다.
        (스크롤 영역 등 컨테이너가 체인에 끼더라도 배제되지 않아야 목록 빈 공간이 잡힌다)
        클릭 좌표(event)가 있으면 실제로 TaskRow 위인지 childAt으로 재확인한다."""
        w = obj if isinstance(obj, QWidget) else None
        if w is None:
            return False
        if isinstance(w, (TaskRow, QPushButton, QToolButton, QScrollArea, QScrollBar,
                          QPlainTextEdit, QTextEdit, QLineEdit, QDateTimeEdit,
                          QComboBox, QCalendarWidget, DdayLabel, AddDateButton)):
            return False
        # 행 내부의 클릭 가능/장식 위젯(반복 아이콘 라벨 등)은 빈 공간 아님
        p = w.parentWidget()
        while p is not None:
            if isinstance(p, TaskRow):
                return False
            p = p.parentWidget()
        # 클릭 대상이 목록 영역(또는 그 컨테이너)이면, 좌표 기반으로 실제 TaskRow 위
        # 클릭인지 확인 (행 본문 빈 영역 클릭도 선택 대상)
        if event is not None:
            # task_list_widget 자신, 그 조상, 또는 자손에서 클릭된 경우
            in_list_area = (w is self.task_list_widget
                            or self.task_list_widget.isAncestorOf(w)
                            or (isinstance(w, QWidget) and w.isAncestorOf(self.task_list_widget)))
            if in_list_area:
                gp = event.globalPos()
                child = self.task_list_widget.childAt(self.task_list_widget.mapFromGlobal(gp))
                cw = child
                while cw is not None:
                    if isinstance(cw, TaskRow):
                        return False
                    cw = cw.parentWidget()
        while w is not None:
            if w is self._top_panel:
                return True
            w = w.parentWidget()
        return False

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
            # 새 업무 입력 영역(업무명·마감일·시간 행) 또는 본문 빈 공간 클릭 → 목록 선택 해제
            def _in_form(w):
                return isinstance(obj, QWidget) and (obj is w or w.isAncestorOf(obj))
            if (_in_form(self.task_input_row) or _in_form(self.date_row_widget)
                    or _in_form(self.input_date) or _in_form(self.input_time)
                    or self._is_empty_area_click(obj, event)):
                self._clear_selection()
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
        color = tc('#aaa') if idx == 0 else tc('#333')
        self.input_recurrence.setStyleSheet(self._recurrence_style_base.format(color=color))
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

    def _apply_memo_line_spacing(self):
        fmt = QTextBlockFormat()
        fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
        cursor = QTextCursor(self.memo_editor.document())
        cursor.select(QTextCursor.Document)
        cursor.mergeBlockFormat(fmt)

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
            _gray.setColor(QColor(tc('#888888')))
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
                         recurrence=task.get('recurrence', ''),
                         notes=task.get('notes', ''),
                         related_no=task.get('related_no', ''))
        delete_task(task['id'])
        self._expanded_note_ids.discard(task['id'])
        if self._selected_task_id == task['id']:
            self._selected_task_id = None
        self._refresh_tasks()

    def _highlight_task(self, task_id):
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            if item and item.widget() and hasattr(item.widget(), 'task'):
                row = item.widget()
                if row.task.get('id') == task_id:
                    row.setStyleSheet(f'background: {tc("rgba(212,184,0,0.35)")}; border-radius: 4px;')
                    QTimer.singleShot(1500, lambda r=row: r._apply_row_style())
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
        sb = self.task_scroll.verticalScrollBar()
        scroll_pos = sb.value()
        while self.task_list_layout.count() > 1:
            item = self.task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for task in get_tasks(self.window_id):
            row = TaskRow(task, self._delete_task, self._refresh_tasks, scale=self._scale,
                          expanded=task['id'] in self._expanded_note_ids,
                          on_toggle=self._on_task_note_toggle,
                          attach_expanded=task['id'] in self._expanded_attach_ids,
                          on_attach_toggle=self._on_task_attach_toggle,
                          on_select=self._select_task,
                          on_navigate=self._move_selection,
                          on_clear=self._clear_selection)
            self.task_list_layout.insertWidget(self.task_list_layout.count() - 1, row)
            if row._expanded:
                row._show_note_editor()
        self._apply_selection()
        QTimer.singleShot(0, lambda: QTimer.singleShot(0, lambda: sb.setValue(scroll_pos)))

    def _on_task_note_toggle(self, task_id, expanded):
        if expanded:
            self._expanded_note_ids.add(task_id)
        else:
            self._expanded_note_ids.discard(task_id)

    def _on_task_attach_toggle(self, task_id, expanded):
        if expanded:
            self._expanded_attach_ids.add(task_id)
        else:
            self._expanded_attach_ids.discard(task_id)

    def _select_task(self, task_id):
        self._selected_task_id = task_id
        self._apply_selection()

    def _clear_selection(self):
        self._selected_task_id = None
        self._apply_selection()

    def _move_selection(self, delta):
        rows = []
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            w = item.widget() if item else None
            if isinstance(w, TaskRow):
                rows.append(w)
        if not rows:
            return
        idx = -1
        if self._selected_task_id is not None:
            for i, r in enumerate(rows):
                if r.task.get('id') == self._selected_task_id:
                    idx = i
                    break
        nidx = idx + delta if idx >= 0 else (0 if delta > 0 else len(rows) - 1)
        nidx = max(0, min(nidx, len(rows) - 1))
        self._select_task(rows[nidx].task['id'])
        rows[nidx].setFocus()
        self.task_scroll.ensureWidgetVisible(rows[nidx], 0, 0)

    def _apply_selection(self):
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            w = item.widget() if item else None
            if isinstance(w, TaskRow):
                w.set_selected(w.task.get('id') == self._selected_task_id)

    def _toggle_selected_note(self):
        if self._selected_task_id is None:
            return
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            w = item.widget() if item else None
            if isinstance(w, TaskRow) and w.task.get('id') == self._selected_task_id:
                w._toggle_note()
                return

    def _start_rename_selected(self):
        if self._selected_task_id is None:
            return
        for i in range(self.task_list_layout.count()):
            item = self.task_list_layout.itemAt(i)
            w = item.widget() if item else None
            if isinstance(w, TaskRow) and w.task.get('id') == self._selected_task_id:
                w._start_rename()
                return

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F2 and self._selected_task_id is not None:
            self._start_rename_selected()
            return
        if e.key() in (Qt.Key_Up, Qt.Key_Down):
            self._move_selection(-1 if e.key() == Qt.Key_Up else 1)
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and self._selected_task_id is not None:
            self._toggle_selected_note()
            return
        super().keyPressEvent(e)

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

    def _open_budget_window(self):
        try:
            import budget
            budget.open_budget_window()
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(self, '오류', f'예산 관리 창 열기 실패:\n{e}')

    def _show_monthly_calendar(self):
        dlg = MonthlyCalendarDialog(self)
        dlg.exec_()
        self._refresh_tasks()

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
        lbl_date.setStyleSheet(f'color: {tct("#888")}; font-size: 9pt;')
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
        btn_del.setStyleSheet(f"""
            QPushButton {{ background: {tc("#fff0f0")}; border: 1px solid {tc("#e08080")};
                          border-radius: 4px; padding: 4px 10px; color: {tct("#a00")}; }}
            QPushButton:hover {{ background: {tc("#ffd8d8")}; }}
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
        btn_close.setStyleSheet(f"""
            QPushButton {{ background: {tc("rgba(255,255,255,0.8)")}; border: 1px solid {tc("#aaa")};
                          border-radius: 4px; padding: 4px 10px; }}
            QPushButton:hover {{ background: {tc("#eee")}; }}
        """)
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        vlay.addLayout(btn_row)

        dlg.exec_()

    def _show_monthly_calendar(self):
        dlg = getattr(self, '_calendar_dlg', None)
        if dlg is not None and dlg.isVisible():
            dlg._load_tasks()
            dlg.raise_()
            dlg.activateWindow()
            return
        dlg = NativeCalendarDialog(self._open_windows, self)
        settings = QSettings('SSNnote', 'SSNnote')
        geo = settings.value('calendar_dlg/geometry')
        if geo:
            dlg.restoreGeometry(geo)
        self._calendar_dlg = dlg
        dlg.finished.connect(lambda *_: settings.setValue('calendar_dlg/geometry', dlg.saveGeometry()))
        dlg.finished.connect(lambda *_: setattr(self, '_calendar_dlg', None))
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _start_capture(self):
        # 모든 메모 창 숨기고 화면 캡처 후 오버레이 표시
        for win in self._open_windows:
            win.hide()
        QApplication.processEvents()
        QTimer.singleShot(250, self._show_capture_overlay)

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

    # ── 빠른 찾기 (Ctrl+K) / 오늘 한눈에 보기 ───────────────────────
    def _palette_commands(self):
        from ui.command_palette import Command
        from school_profile import privacy_enabled
        target = self._task_list_window()
        on_top = bool(self.windowFlags() & Qt.WindowStaysOnTopHint)

        def new_task():
            target.show()
            target.raise_()
            target.activateWindow()
            target.input_name.setFocus()

        new_memo = lambda: self.on_new(offset_from=self, memo_mode=True,  # noqa: E731
                                       on_toggle_hotkey=self._on_toggle_hotkey,
                                       on_shortcut_change=self._on_shortcut_change,
                                       get_shortcut_enabled=self._get_shortcut_enabled)
        # 순서·이름은 「…」 메뉴와 맞춘다 (자주 쓰는 기능 → 창 → 이 메모장 → 설정 → 도움말)
        return [
            Command('📋 오늘 한눈에 보기', self._open_dashboard, '', 'dashboard today 대시보드 마감'),
            Command('📸 AI 공문 캡처', self._start_ai_capture, 'Ctrl+Shift+X', 'capture ocr 캡쳐 스크린샷'),
            Command('✏️ 새 업무 입력', new_task, '', 'task add 추가 할일'),
            Command('🆕 새 메모장', new_memo, '', 'new memo 메모'),
            Command('✍️ 공문 작성', self._open_document_editor, '', 'document 기안 초안 회신'),
            Command('₩ 예산 관리', self._open_budget_window, '', 'budget money 돈 집행'),
            Command('📅 월간 달력', self._show_monthly_calendar, '', 'calendar 일정 캘린더'),
            Command('📜 지난 기록', self._show_history, '', 'history 완료 삭제'),
            Command('📌 항상 위에 ' + ('고정 끄기' if on_top else '고정'), self.toggle_always_on_top,
                    'Ctrl+Shift+F', 'pin top 핀'),
            Command('↕️ 메모장 접기/펼치기', self.toggle_shade, 'Ctrl+Shift+R', 'shade fold collapse'),
            Command('🌙 다크 모드 켜기/끄기', self._toggle_theme, '', 'dark light theme 테마 어둡게 밝게 설정'),
            Command('🛡️ 개인정보 보호 모드 ' + ('끄기' if privacy_enabled() else '켜기'),
                    self._toggle_privacy_mode, '', 'privacy 보안 마스킹 설정'),
            Command('🏫 학교 정보 설정', self._open_profile_dialog, '', 'profile school 프로필 담당자 설정'),
            Command('🔑 AI 연결 (API 키)', self._open_ai_settings, '', 'ai api key opencode 오픈코드 gemini chatgpt 설정'),
            Command('🤖 Claude 연결', self._open_claude_connect, '', 'claude mcp ai 연결 클로드 설정'),
            Command('💡 사용 설명서', self.show_help, '', 'help 도움말'),
            Command('⬇️ 업데이트 확인', lambda: updater.check_for_update_manual(self), '', 'update 버전 도움말'),
        ]

    def _open_command_palette(self):
        from ui.command_palette import CommandPalette
        from db import search_tasks_all
        old = getattr(self, '_palette', None)
        try:
            if old is not None and old.isVisible():
                old.close()
                return
        except RuntimeError:
            pass  # 이미 닫혀 정리된 팔레트
        self._palette = CommandPalette(self._palette_commands(), tasks=search_tasks_all(''),
                                       on_task=self._jump_to_task)
        self._palette.popup(self)

    def _jump_to_task(self, task):
        """업무가 있는 메모 창을 앞으로 가져와 그 업무를 잠깐 강조한다."""
        for win in self._open_windows:
            if win.window_id == task.get('window_id'):
                win.show()
                win.raise_()
                win.activateWindow()
                win._highlight_task(task.get('id'))
                return

    def _open_dashboard(self):
        from ui.dashboard_dialog import DashboardDialog
        old = getattr(self, '_dashboard', None)
        try:
            if old is not None:
                old.close()
        except RuntimeError:
            pass
        dlg = DashboardDialog()
        dlg.setAttribute(Qt.WA_DeleteOnClose)
        dlg.jump_requested.connect(self._jump_to_task)
        dlg.open_budget.connect(self._open_budget_window)
        self._dashboard = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    # ── AI 공문 캡처 (캡처 → AI 분석 → 업무 등록 / 회신 공문) ─────────
    def _start_ai_capture(self):
        for win in self._open_windows:
            win.hide()
        QApplication.processEvents()
        QTimer.singleShot(250, self._show_ai_capture_overlay)

    def _show_ai_capture_overlay(self):
        screenshot = grab_fullscreen()
        self._overlay = ScreenCaptureOverlay(screenshot)
        self._overlay.region_captured.connect(self._on_ai_capture_complete)
        self._overlay.cancelled.connect(self._restore_windows)
        self._overlay.show()
        self._overlay.activateWindow()
        self._overlay.setFocus()

    def _on_ai_capture_complete(self, pixmap):
        self._overlay.region_captured.disconnect()  # 중복 호출 방지
        self._restore_windows()
        from ui.capture_dialog import CaptureResultDialog
        dlg = CaptureResultDialog(pixmap, on_register=self._register_captured_task,
                                  on_reply=self._reply_from_capture, parent=None)
        self._capture_dlg = dlg  # GC 방지
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _task_list_window(self):
        """업무를 넣을 창: 자신이 업무 목록 모드면 자신, 아니면 열린 업무 목록 창."""
        if not self._memo_mode:
            return self
        for win in self._open_windows:
            if not getattr(win, '_memo_mode', True):
                return win
        return self

    def _register_captured_task(self, name, deadline, notes, related_no):
        target = self._task_list_window()
        add_task(target.window_id, name, deadline, notes=notes, related_no=related_no)
        target._refresh_tasks()
        target.show()
        target.raise_()

    def _reply_from_capture(self, analysis):
        self._open_document_editor()
        editor = getattr(self, '_doc_editor', None)
        if editor is not None:
            editor.start_reply(analysis)
            editor.raise_()
            editor.activateWindow()

    def _open_profile_dialog(self):
        from ui.profile_dialog import ProfileDialog
        ProfileDialog(self).exec_()

    def _open_ai_settings(self):
        """AI 공문 캡처·공문 초안에 쓸 AI(OpenCode, Claude API 등) 설정 — 공문 작성 창의 AI 설정을 연다."""
        self._open_document_editor()
        editor = getattr(self, '_doc_editor', None)
        if editor is not None:
            editor.raise_()
            QTimer.singleShot(250, editor._show_ai_settings_menu)

    def _open_claude_connect(self):
        """Claude Code / Claude 데스크톱 앱에 서서니 노트 연결 (MCP)."""
        from ui.claude_connect_dialog import ClaudeConnectDialog
        old = getattr(self, '_claude_dlg', None)
        try:
            if old is not None and old.isVisible():
                old.raise_()
                old.activateWindow()
                return
        except RuntimeError:
            pass
        self._claude_dlg = ClaudeConnectDialog()
        self._claude_dlg.show()

    def _toggle_theme(self):
        """라이트 ↔ 다크. 화면들이 켤 때 색을 정하므로 다음 실행부터 적용."""
        from ui.theme import MODE, set_theme
        from db import get_setting
        saved = get_setting('theme', MODE)
        new = 'light' if saved == 'dark' else 'dark'
        set_theme(new)
        label = '다크' if new == 'dark' else '라이트'
        if new == MODE:
            msg = f'{label} 모드로 되돌렸습니다. 지금 화면 그대로 유지됩니다.'
        else:
            msg = (f'{label} 모드로 바꿨습니다.\n\n'
                   '트레이 아이콘 → 종료 후 서서니 노트를 다시 켜면 적용됩니다.\n'
                   '(예산 관리·공문 작성 창은 아직 라이트로 표시됩니다)')
        QMessageBox.information(self, '화면 테마', msg)

    def _toggle_privacy_mode(self):
        from school_profile import privacy_enabled, set_privacy_enabled
        set_privacy_enabled(not privacy_enabled())

    def _make_menu_style(self):
        return f"""
            QMenu {{
                background: {T.INPUT_BG};
                border: 1px solid {T.ACCENT};
                border-radius: 6px;
                padding: 4px 0;
                font-family: 'Malgun Gothic';
                font-size: 11pt;
                color: {tct("#333")};
                min-width: 176px;
            }}
            QMenu::item {{ padding: 7px 16px 7px 4px; }}
            QMenu::item:selected {{ background: {tc("rgba(212,184,0,0.3)")}; }}
            QMenu::item:disabled {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; padding: 6px 16px 2px 10px; }}
            QMenu::item:disabled:selected {{ background: transparent; }}
        """

    @staticmethod
    def _menu_check_icon(checked, size=16):
        """메뉴용 체크 표시 — 켜짐은 ✓, 꺼짐은 같은 크기의 빈 칸 (글자 줄이 맞도록)."""
        from PyQt5.QtGui import QPixmap, QPainter, QPen, QIcon
        px = QPixmap(size, size)
        px.fill(Qt.transparent)
        if checked:
            p = QPainter(px)
            p.setRenderHint(QPainter.Antialiasing)
            pen = QPen(QColor(T.TEXT_ON_NOTE), 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
            p.setPen(pen)
            p.drawPolyline(QPointF(size * 0.18, size * 0.52), QPointF(size * 0.42, size * 0.76),
                           QPointF(size * 0.84, size * 0.26))
            p.end()
        return QIcon(px)

    def _menu_add(self, menu, text, on_trigger, checked=None, shortcut=''):
        """메뉴 항목 추가. checked가 True/False면 체크 항목, shortcut은 오른쪽 끝에 표시."""
        act = QAction(text + (f'\t{shortcut}' if shortcut else ''), menu)
        if checked is not None:
            act.setIcon(self._menu_check_icon(checked))
        act.triggered.connect(lambda *_: on_trigger())
        menu.addAction(act)
        return act

    @staticmethod
    def _menu_header(menu, text):
        """누를 수 없는 작은 머리글 (스타일시트를 쓴 메뉴에서는 addSection 글자가 안 보여서 대신 사용)."""
        act = QAction(text, menu)
        act.setEnabled(False)
        menu.addAction(act)
        return act

    def _submenu(self, parent, title, min_width=None):
        sub = QMenu(title, parent)
        style = self._make_menu_style()
        if min_width:
            style = style.replace('min-width: 176px', f'min-width: {min_width}px')
        sub.setStyleSheet(style)
        sub.setToolTipsVisible(True)
        parent.addMenu(sub)
        return sub

    def show_menu(self, btn):
        """「…」 메뉴: 자주 쓰는 기능 → 이 메모장 → 설정 → 도움말 → (삭제)."""
        menu = QMenu(self)
        menu.setStyleSheet(self._make_menu_style())
        add = self._menu_add

        # ── 자주 쓰는 기능 ──
        add(menu, '📋 오늘 한눈에 보기', self._open_dashboard)
        add(menu, '📸 AI 공문 캡처', self._start_ai_capture, shortcut='Ctrl+Shift+X')
        add(menu, '🔍 빠른 찾기', self._open_command_palette, shortcut='Ctrl+K')
        add(menu, '🆕 새 메모장', lambda: self.on_new(offset_from=self, memo_mode=True,
                                                    on_toggle_hotkey=self._on_toggle_hotkey,
                                                    on_shortcut_change=self._on_shortcut_change,
                                                    get_shortcut_enabled=self._get_shortcut_enabled))
        menu.addSeparator()

        # ── 이 메모장 (이 창에만 적용) ──
        memo = self._submenu(menu, '🎨 이 메모장')
        color_widget = QWidget()
        color_widget.setStyleSheet('background: transparent;')
        grid = QGridLayout(color_widget)
        grid.setContentsMargins(10, 8, 10, 8)
        grid.setSpacing(6)
        for i, color in enumerate(PALETTE):
            border = f'2px solid {tc("#333")}' if color == '#FFFFFF' else f'1px solid {tc("rgba(0,0,0,0.15)")}'
            cb = QPushButton()
            cb.setFixedSize(26, 26)
            cb.setToolTip('배경색')
            cb.setStyleSheet(f"""
                QPushButton {{
                    background: {note_color(color)};
                    border-radius: 13px;
                    border: {border};
                }}
                QPushButton:hover {{ border: 2px solid {tc("rgba(0,0,0,0.5)")}; }}
            """)
            cb.clicked.connect(lambda _, c=color: (self._apply_color(c), self.save_state(), menu.close()))
            grid.addWidget(cb, i // 6, i % 6)
        self._menu_header(memo, '배경색')
        wa = QWidgetAction(memo)
        wa.setDefaultWidget(color_widget)
        memo.addAction(wa)
        size_menu = self._submenu(memo, '글자 크기', 110)
        for label, val in [('기본', 1.0), ('크게', 1.1), ('더 크게', 1.2)]:
            add(size_menu, label, lambda v=val: self._set_scale(v, menu), checked=abs(self._scale - val) < 0.01)
        gap_menu = self._submenu(memo, '줄 간격', 110)
        for label, val in [('0px', 0), ('5px', 5), ('10px', 10), ('15px', 15)]:
            add(gap_menu, label, lambda v=val: self._set_row_gap(v, menu), checked=self._row_gap == val)

        # ── 설정 (앱 전체에 적용) ──
        settings = self._submenu(menu, '⚙️ 설정', 230)
        alarm = self._submenu(settings, '마감 알림', 230)
        self._menu_header(alarm, 'D-day 5일 이내 알림')
        current_min = self._get_alarm_interval() if self._get_alarm_interval else 180
        alarm_options = [('사용 안 함', 0), ('30분마다 (잡도리 모드)', 30),
                         ('1시간마다 (긴장 모드)', 60), ('3시간마다 (기본)', 180)]
        for label, minutes in alarm_options:
            add(alarm, label, lambda m=minutes: self._set_alarm_interval(m, menu), checked=current_min == minutes)
        is_custom = current_min not in [m for _, m in alarm_options]
        add(alarm, '사용자 설정…' + (f' ({max(1, round(current_min / 60))}시간마다)' if is_custom else ''),
            lambda: self._set_alarm_interval_custom(menu), checked=is_custom)
        alarm.addSeparator()
        timed_enabled = self._get_timed_alarm_enabled() if self._get_timed_alarm_enabled else True
        act_timed = add(alarm, '당일 마감 3시간 전부터 알림', lambda: self._set_timed_alarm_enabled(not timed_enabled, menu),
                        checked=timed_enabled)
        act_timed.setToolTip('3시간, 2시간, 1시간, 30분 전에 팝업 알림이 뜹니다.')

        sc_enabled = self._get_shortcut_enabled() if self._get_shortcut_enabled else True
        add(settings, '단축키 사용', lambda: self._toggle_shortcut(not sc_enabled), checked=sc_enabled)
        autostart_on = autostart_is_enabled()
        add(settings, '시작 시 자동실행', lambda: autostart_set(not autostart_on), checked=autostart_on)
        settings.addSeparator()
        from db import get_setting as _gs
        from ui.theme import MODE as _mode
        from school_profile import privacy_enabled
        add(settings, '다크 모드 (다시 켜면 적용)', self._toggle_theme, checked=_gs('theme', _mode) == 'dark')
        add(settings, '개인정보 보호 모드', self._toggle_privacy_mode, checked=privacy_enabled())
        add(settings, '학교 정보 설정…', self._open_profile_dialog)
        add(settings, 'AI 연결 (API 키)…', self._open_ai_settings)
        add(settings, 'Claude 연결…', self._open_claude_connect)

        # ── 도움말 ──
        help_menu = self._submenu(menu, '💡 도움말', 150)
        add(help_menu, '사용 설명서', self.show_help)
        add(help_menu, '업데이트 확인', lambda: updater.check_for_update_manual(self))

        # ── 메모장 삭제 (첫 메모장 제외) ──
        _all_ids = [w['id'] for w in get_all_windows()]
        if not _all_ids or self.window_id != min(_all_ids):
            menu.addSeparator()
            add(menu, '🗑️ 메모장 삭제', self.delete_memo)

        self._last_menu = menu  # 테스트·디버그용 참조
        pos = btn.mapToGlobal(btn.rect().bottomLeft())
        menu.exec_(pos)

    _SEP_COLOR_MAP = {
        '#CCFF90': '#A3CC73',
        '#A8F0E8': '#86C0BA',
        '#AECBFA': '#8BA2C8',
        '#D7AEFB': '#AC8BC8',
        '#FDCFE8': '#CAA6BA',
        '#E6C9A8': '#B8A186',
        '#AAAAAA': '#888888',
        '#DDDDDD': '#B1B1B1',
        '#FFFFFF': '#CCCCCC',
    }

    def _apply_color(self, hex_color):
        self._bg_color = hex_color            # 저장값(라이트 색) 그대로 유지
        shown = note_color(hex_color)          # 다크 테마면 같은 색의 깊은 톤
        title_color = QColor(shown).darker(115).name()
        self.centralWidget().setStyleSheet(f"""
            QWidget#central {{
                background: {shown};
                border-radius: 0;
                border: 1px solid {tc("rgba(0,0,0,0.4)")};
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
            QPushButton:hover {{ background: {tc("rgba(0,0,0,0.12)")}; }}
            QPushButton#btn_mode {{
                background: transparent;
                border: 1px solid transparent;
                border-radius: 6px;
                padding: 4px 8px;
                color: {T.TEXT_ON_NOTE};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                font-weight: bold;
            }}
            QPushButton#btn_mode:hover {{ background: {tc("rgba(0,0,0,0.12)")}; border-color: {tc("rgba(0,0,0,0.12)")}; }}
            QPushButton#btn_mode:pressed {{ background: {tc("rgba(0,0,0,0.20)")}; border-color: {tc("rgba(0,0,0,0.20)")}; }}
        """)
        if DARK:
            sep_color = QColor(shown).lighter(160).name()
        else:
            sep_color = self._SEP_COLOR_MAP.get(hex_color.upper(), T.DIVIDER)
        sep_style = f'background: {sep_color}; border: none;'
        self._title_sep.setStyleSheet(sep_style)
        self._sep_date.setStyleSheet(sep_style)
        self._sep_bottom.setStyleSheet(sep_style)

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
        dlg.setStyleSheet(f"font-family: 'Malgun Gothic'; font-size: 11pt; background: {tc('#fffef0')};")

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        if not records:
            lbl = QLabel('기록이 없습니다.')
            lbl.setStyleSheet(f'color: {tct("#999")};')
            layout.addWidget(lbl)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
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
                    and not getattr(layout.itemAt(i).widget(), '_is_memo_panel', False)
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
                         recurrence=r.get('recurrence', ''),
                         notes=r.get('notes', ''),
                         related_no=r.get('related_no', ''))
                delete_task_history(r['id'])
                self._refresh_tasks()
                mw = getattr(row_widget, '_memo_widget', None)
                if mw is not None:
                    mw.hide()
                row_widget.hide()
                refresh_folder(btn_toggle, container, year, month)

            def delete_history(r, row_widget, btn_toggle, container, year, month):
                delete_task_history(r['id'])
                mw = getattr(row_widget, '_memo_widget', None)
                if mw is not None:
                    mw.hide()
                row_widget.hide()
                refresh_folder(btn_toggle, container, year, month)

            def make_row(r, btn_toggle, container, year, month):
                name     = r['name']
                deadline = r['deadline']
                cleared  = r['cleared_at'][:10]
                note     = r.get('notes', '') or ''
                related  = r.get('related_no', '') or ''
                has_memo = bool(note.strip() or related.strip())

                overdue = False
                if deadline:
                    dday = calc_dday(deadline)
                    overdue = dday.startswith('D+')
                    line = f"{name}  ({deadline} / {dday})"
                else:
                    line = name

                row_widget = QWidget()
                row_widget.setStyleSheet(f'background: {tc("rgba(255,255,255,0.6)")}; border-radius: 4px;')
                rh = QHBoxLayout(row_widget)
                rh.setContentsMargins(20, 4, 8, 4)
                rh.setSpacing(6)

                lbl_task = QLineEdit(line)
                lbl_task.setReadOnly(True)
                lbl_task.setFrame(False)
                lbl_task.setStyleSheet(f"""
                    QLineEdit {{
                        color: {tct("#333")};
                        background: transparent;
                        border: none;
                        padding: 0;
                        font-family: 'Malgun Gothic';
                        font-size: 11pt;
                    }}
                """)
                lbl_task.setCursorPosition(0)
                lbl_cleared = _SelectableLabel(f'삭제: {cleared}')
                lbl_cleared.setStyleSheet(f'color: {tct("#999")}; font-size: 9pt; background: transparent;')
                lbl_cleared.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

                rh.addWidget(lbl_task, 1)
                btn_note = None
                if has_memo:
                    btn_note = _NoteButton(scale=1.0, glyph_offset=0)
                    btn_note.set_note_color(True)
                    btn_note.setToolTip('업무 메모 (펼치기/접기)')
                    rh.addWidget(btn_note)
                rh.addSpacing(11)
                rh.addWidget(lbl_cleared)

                if not overdue:
                    btn_restore = QPushButton('복원')
                    btn_restore.setFixedHeight(22)
                    btn_restore.setStyleSheet(f"""
                        QPushButton {{
                            background: {tc("#d4f0d4")}; border: 1px solid {tc("#7cc87c")};
                            border-radius: 4px; color: {tct("#2a6e2a")};
                            font-size: 9pt; padding: 0 8px;
                        }}
                        QPushButton:hover {{ background: {tc("#b6e8b6")}; }}
                    """)
                    btn_restore.clicked.connect(lambda _, rec=r, rw=row_widget: restore_task(rec, rw, btn_toggle, container, year, month))
                    rh.addWidget(btn_restore)

                btn_del = QPushButton('삭제')
                btn_del.setFixedHeight(22)
                btn_del.setStyleSheet(f"""
                    QPushButton {{
                        background: {tc("#f0d4d4")}; border: 1px solid {tc("#c87c7c")};
                        border-radius: 4px; color: {tct("#6e2a2a")};
                        font-size: 9pt; padding: 0 8px;
                    }}
                    QPushButton:hover {{ background: {tc("#e8b6b6")}; }}
                """)
                btn_del.clicked.connect(lambda _, rec=r, rw=row_widget: delete_history(rec, rw, btn_toggle, container, year, month))
                rh.addWidget(btn_del)

                container.layout().addWidget(row_widget)

                if has_memo:
                    memo_widget = QWidget()
                    memo_widget._is_memo_panel = True
                    memo_widget.setStyleSheet(f'background: {tc("rgba(255,255,255,0.7)")}; border-radius: 4px;')
                    mv = QVBoxLayout(memo_widget)
                    mv.setContentsMargins(28, 4, 8, 4)
                    mv.setSpacing(3)
                    if related.strip():
                        lbl_ref = _SelectableLabel(f'공문: {related}')
                        lbl_ref.setStyleSheet(f'color: {tct("#7a6a00")}; font-size: 9pt; background: transparent; font-weight: bold;')
                        lbl_ref.setWordWrap(True)
                        mv.addWidget(lbl_ref)
                    if note.strip():
                        lbl_note = _SelectableLabel(note)
                        lbl_note.setStyleSheet(f'color: {tct("#444")}; font-size: 10pt; background: transparent;')
                        lbl_note.setWordWrap(True)
                        mv.addWidget(lbl_note)
                    memo_widget.hide()

                    def toggle_note(_, w=memo_widget, b=btn_note):
                        if w.isHidden():
                            w.show()
                            b.set_expanded(True)
                        else:
                            w.hide()
                            b.set_expanded(False)

                    btn_note.clicked.connect(toggle_note)
                    row_widget._memo_widget = memo_widget
                    container.layout().addWidget(memo_widget)

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
                btn_toggle.setStyleSheet(f"""
                    QPushButton {{
                        background: {tc("rgba(200,190,150,0.25)")};
                        border: 1px solid {T.BORDER_SOFT};
                        border-radius: 4px;
                        color: {tct("#5a4a00")};
                        font-size: 10pt;
                        font-weight: bold;
                        text-align: left;
                        padding: 5px 10px;
                    }}
                    QPushButton:hover {{ background: {tc("rgba(200,190,150,0.5)")}; }}
                """)

                month_container = QWidget()
                month_container.setObjectName('monthContainer')
                month_container.setStyleSheet(f"""
                    QWidget#monthContainer {{
                        background: transparent;
                        border: 1px solid {T.BORDER_SOFT};
                        border-radius: 4px;
                    }}
                """)
                mc_vbox = QVBoxLayout(month_container)
                mc_vbox.setContentsMargins(12, 4, 8, 4)
                mc_vbox.setSpacing(4)
                month_container.hide()

                for r in grp:
                    make_row(r, btn_toggle, month_container, year, month)

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
                btn_cur.setStyleSheet(f"""
                    QPushButton {{
                        background: {tc("rgba(200,190,150,0.25)")};
                        border: 1px solid {T.BORDER_SOFT};
                        border-radius: 4px;
                        color: {tct("#5a4a00")};
                        font-size: 10pt;
                        font-weight: bold;
                        text-align: left;
                        padding: 5px 10px;
                    }}
                    QPushButton:hover {{ background: {tc("rgba(200,190,150,0.5)")}; }}
                """)

                cur_container = QWidget()
                cur_container.setObjectName('monthContainer')
                cur_container.setStyleSheet(f"""
                    QWidget#monthContainer {{
                        background: transparent;
                        border: 1px solid {T.BORDER_SOFT};
                        border-radius: 4px;
                    }}
                """)
                cur_vbox = QVBoxLayout(cur_container)
                cur_vbox.setContentsMargins(12, 4, 8, 4)
                cur_vbox.setSpacing(4)

                for r in current_records:
                    make_row(r, btn_cur, cur_container, cur_year, cur_month)

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
        dlg.setWindowTitle('사용 설명서')
        dlg.setMinimumWidth(460)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(12, 12, 12, 12)
        icon_path = os.path.join(_base_path(), 'assets', '수정 아이콘.png')
        icon_url = QUrl.fromLocalFile(icon_path).toString()
        _load_pretendard()
        lbl = QLabel(
            '<p style="line-height: 160%; font-family: Pretendard; font-size: 12pt;">'
            '<b>〈업무〉</b><br>'
            '1. 업무명 칸에서 <b>Tab</b> 키를 누르면 다음 설정 칸으로 이동<br>'
            '2. 업무 목록에서 업무를 더블 클릭하거나 선택한 후 <b>F2</b> 키를 누르면 업무명 변경 가능<br>'
            '3. 업무 목록을 선택하고 위아래 방향키로 이동 가능<br>'
            '4. 업무 목록을 선택하고 <b>Enter</b> 키를 누르면 메모장 열림<br>'
            '<br>'
            '<b>〈기타〉</b><br>'
            "1. '메뉴 - 업데이트 확인' 누르면 업데이트 가능. 이때, 학교 컴퓨터에서는 2번 해야 가능함.<br>"
            '&nbsp;&nbsp;(방화벽 때문에 처음 한 번은 막히는 듯?)<br>'
            f'2. <img src="{icon_url}" width="20" height="20" style="vertical-align: middle;"> 버튼을 누르면 메모 모드로 변환. 한 번 더 누르면 원위치.<br>'
            '3. <b>₩</b> 버튼을 누르면 예산 관리 창이 열립니다.<br>'
            '4. 나머지는 직접 눌러보세요. 단순합니다.<br>'
            '<br>'
            '<b>〈단축키〉</b><br>'
            '1. <b>ctrl + s</b> : 앱 포커싱 (다른 작업 하다가 이거 누르면 서서니 노트 앱으로 집중)<br>'
            '2. <b>ctrl + r</b> : 제목표시줄만 남기고 접기/펼치기<br>'
            '3. <b>ctrl + f</b> : 항상 화면 위에 고정(Fix)'
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
            btn.setStyleSheet(f'background: {tc("rgba(0,0,0,0.18)")}; border-radius: 4px;')
            btn.setGraphicsEffect(None)
        else:
            btn.setText('📌')
            btn.setToolTip('항상 위 꺼짐')
            btn.setStyleSheet('')
            _gray = QGraphicsColorizeEffect()
            _gray.setColor(QColor(tc('#888888')))
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

    def _snap_resize(self, nx, ny, nw, nh, edge):
        """리사이즈 드래그 중 snap 계산. 스냅된 (nx, ny, nw, nh) 반환."""
        moves_left  = 'left'   in edge
        moves_right = 'right'  in edge
        moves_top   = 'top'    in edge
        moves_bot   = 'bottom' in edge

        right_anchor  = nx + nw
        bottom_anchor = ny + nh

        if moves_left or moves_right:
            my_x = nx if moves_left else nx + nw
            best_d, snap_x = SNAP_THRESHOLD, None
            for other in self._open_windows:
                if other is self:
                    continue
                for candidate in (other.x(), other.x() + other.width()):
                    d = abs(candidate - my_x)
                    if d < best_d:
                        best_d, snap_x = d, candidate
            if snap_x is not None:
                if moves_left:
                    new_nx = snap_x
                    new_nw = right_anchor - new_nx
                    if new_nw >= EdgeHandle.MIN_W:
                        nx, nw = new_nx, new_nw
                else:
                    new_nw = snap_x - nx
                    if new_nw >= EdgeHandle.MIN_W:
                        nw = new_nw

        if moves_top or moves_bot:
            my_y = ny if moves_top else ny + nh
            best_d, snap_y = SNAP_THRESHOLD, None
            for other in self._open_windows:
                if other is self:
                    continue
                for candidate in (other.y(), other.y() + other.height()):
                    d = abs(candidate - my_y)
                    if d < best_d:
                        best_d, snap_y = d, candidate
            if snap_y is not None:
                if moves_top:
                    new_ny = snap_y
                    new_nh = bottom_anchor - new_ny
                    if new_nh >= EdgeHandle.MIN_H:
                        ny, nh = new_ny, new_nh
                else:
                    new_nh = snap_y - ny
                    if new_nh >= EdgeHandle.MIN_H:
                        nh = new_nh

        return nx, ny, nw, nh

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

    def _apply_row_gap(self):
        """업무 목록 칸(TaskRow) 사이 간격 적용."""
        self.task_list_layout.setSpacing(self._row_gap)

    def _set_row_gap(self, gap, menu=None):
        self._row_gap = gap
        self._apply_row_gap()
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
            row_gap=self._row_gap,
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
