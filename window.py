import ctypes
from datetime import date, timedelta
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QDateTimeEdit, QWidgetAction,
    QMessageBox, QDialog, QGridLayout, QCalendarWidget, QToolButton,
    QPlainTextEdit, QTextEdit, QTimeEdit,
    QComboBox, QScrollBar
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QDateTime, QPoint, QSettings, QUrl, QSize
from PyQt5.QtGui import QFont, QColor, QTextBlockFormat, QTextCursor, QFontMetrics
from db import (update_window, delete_window, get_tasks, add_task, delete_task, update_task,
                add_task_history, delete_task_history,
                get_documents, add_document, update_document, delete_document,
                get_window_memo, set_window_memo, get_all_windows, get_window_font, set_window_font,
                set_window_merge_group, clear_window_merge_group, get_next_merge_group_id)
from autostart import is_enabled as autostart_is_enabled, set_enabled as autostart_set
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, _normalize_doc_number
import updater
from ui.common import (pr_font, note_font, note_family, note_font_available, set_current_note_font,
                       refresh_installed_fonts, NOTE_FONTS, DEFAULT_NOTE_FONT)
from ui.common import SNAP_THRESHOLD, SNAP_ZONE, TITLE_BAR_HEIGHT, fit_to_screen, px
from ui.dates import _next_recurrence_deadline
from ui.widgets import AddDateButton, CustomCalendarWidget, DdayLabel, PlainPasteMemoEdit
from ui.dialogs import CustomRecurrenceDialog
from ui.chrome import EdgeHandle, MergeButton, TitleBar
from ui.task_row import TaskRow
from ui.document_row import DocumentRow
from ui.calendar_view import MonthlyCalendarDialog, NativeCalendarDialog
from ui import menu_style
from ui import ai_chat_window
from ui.theme import T, PALETTE, tc, tct, note_color, qcolor, DARK, accent_for_note, set_current_accent, rgba
import qtawesome as qta

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
        self._accent         = T.ACCENT       # 메모 색에 맞춘 포인트 색 (_apply_color에서 정함)
        self._capture_hint_shown = False
        self._scale          = 1.0
        self._row_gap        = 0
        self._font_key       = get_window_font(window_id) or DEFAULT_NOTE_FONT   # 창마다 고른 글꼴
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
        # 메모장(메모 전용 창)은 단축키 없이 마우스로만 쓴다 — 핀·접기는 제목 표시줄 버튼과 더블클릭으로 충분
        self._shortcuts = []
        if not self._force_memo_mode:
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
        # 큰 모니터에서 쓰던 위치·크기라도 지금 화면 안에 들어오게 (노트북·모니터 분리 대비)
        x, y, width, height = fit_to_screen(x, y, max(width, EdgeHandle.MIN_W), height, max_frac=0.95,
                                            min_w=EdgeHandle.MIN_W)
        self.expanded_height = height
        self.collapsed       = collapsed
        self.move(x, y)
        self.resize(width, TITLE_BAR_HEIGHT if collapsed else height)
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

        # 1행: 업무
        self.task_input_row = QWidget()
        self.task_input_row.setStyleSheet('background: transparent;')
        row1 = QHBoxLayout(self.task_input_row)
        row1.setContentsMargins(16, 0, 8, 0)
        row1.setSpacing(4)
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText('＋ 업무 추가')
        self.input_name.setToolTip('업무 이름을 쓰고 Enter — Tab을 누르면 마감일로')
        self.input_name.setFont(note_font(11, self._font_key))
        self._date_touched = False
        self._time_touched = False
        self.input_name.installEventFilter(self)
        self.input_name.returnPressed.connect(self._add_task)
        self.input_name.textChanged.connect(self._on_name_cleared)
        row1.addWidget(self.input_name, 1)
        content_layout.addWidget(self.task_input_row)

        # 2행: 마감일·시각·반복 — 업무 추가 칸을 누르거나 값이 있을 때만 보임 (_sync_input_details)
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
        self.input_date.setFont(QFont('Malgun Gothic', 11))
        self.input_date.dateChanged.connect(self._on_date_changed)
        self.input_date.installEventFilter(self)
        for child in self.input_date.findChildren(QWidget):
            child.installEventFilter(self)
        self.input_time = QTimeEdit(QTime(9, 0))
        self.input_time.setDisplayFormat('HH:mm')
        self.input_time.setFont(QFont('Malgun Gothic', 11))
        self.input_time.setEnabled(False)
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
        self._cal_popup.setMinimumSize(px(300), px(260))
        self._cal_popup.clicked.connect(self._on_cal_date_selected)

        self.input_recurrence = QComboBox()
        self.input_recurrence.addItems(['반복 없음', '매주', '격주', '매월', '매년', '사용자 설정...'])
        for _i in range(self.input_recurrence.count()):
            self.input_recurrence.setItemData(_i, Qt.AlignLeft, Qt.TextAlignmentRole)
        self.input_recurrence.installEventFilter(self)
        self.input_recurrence.setFont(QFont('Malgun Gothic', 11))
        self._custom_recurrence_val = ''
        self.input_recurrence.currentIndexChanged.connect(self._on_recurrence_index_changed)

        self._size_inputs()
        self.date_row_widget.hide()
        QApplication.instance().focusChanged.connect(self._on_app_focus_changed)

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
        content_layout.addWidget(self.task_scroll)

        # 메모 모드 자유 편집기 (기본 숨김)
        self.memo_editor = PlainPasteMemoEdit()
        self.memo_editor.setFont(note_font(12, self._font_key))   # 글자 크기를 바꿔도 고른 글꼴이 유지됨
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
        bottom_layout.setContentsMargins(8, 7, 8, 8)
        bottom_layout.setSpacing(6)

        # 세 버튼 모두 같은 모양 (아이콘 + 글자, 은은한 바탕) — 진한 색은 쓰지 않는다
        _btn_style = f"""
            QPushButton {{
                background: {tc("rgba(255,255,255,0.45)")};
                border: 1px solid {tc("rgba(0,0,0,0.08)")};
                border-radius: 8px;
                padding: 6px 10px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: {tc("rgba(255,255,255,0.75)")}; border-color: {tc("rgba(0,0,0,0.14)")}; }}
            QPushButton:pressed {{ background: {tc("rgba(0,0,0,0.06)")}; }}
        """
        _btn_font = QFont('Pretendard GOV', 10)
        _btn_font.setFamilies(['Pretendard GOV', 'Pretendard', 'Malgun Gothic'])
        _btn_font.setWeight(QFont.DemiBold)
        _btn_icon = qcolor('#000000', 165)

        def _bottom_button(text, icon, slot, tip):
            b = QPushButton(text)
            b.setFont(_btn_font)
            b.setStyleSheet(_btn_style)
            b.setIcon(qta.icon(icon, color=_btn_icon))
            b.setIconSize(QSize(px(15), px(15)))
            b.setCursor(Qt.PointingHandCursor)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            bottom_layout.addWidget(b)
            return b

        _bottom_button('달력', 'ph.calendar-blank-bold', self._show_monthly_calendar, '월간 달력 (모든 메모장의 마감 업무)')
        _bottom_button('공문 작성', 'ph.pen-nib-bold', self._open_document_editor,
                       '공문 작성 창 (AI 초안·공문 캡처)')
        _bottom_button('지난 기록', 'ph.clock-counter-clockwise-bold', self._show_history,
                       '지난 업무(삭제한 업무 되살리기) · 작성 공문(저장한 공문)')

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

        # 급함 표시: 30초마다 남은 시간·단계·기한 지남을 고침 (ui/motion.py)
        self._urgency_timer = QTimer(self)
        self._urgency_timer.setInterval(30 * 1000)
        self._urgency_timer.timeout.connect(self._tick_urgency)
        self._urgency_timer.start()

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
            self.title_bar.btn_ai.hide()

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
        self._sync_input_details()
        if not text:
            self.input_date.setDate(QDate(1900, 1, 1))
            self.input_time.setEnabled(False)
            self.input_time.setTime(QTime(9, 0))
            self._date_touched = False
            self._time_touched = False

    def _on_recurrence_index_changed(self, idx):
        self.input_recurrence.setStyleSheet(self._recur_style(idx != 0))
        self._sync_input_details()
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
        self._sync_input_details()
        self.task_scroll.setVisible(not on)
        self.memo_editor.setVisible(on)
        btn = self.title_bar.btn_mode
        self.title_bar.set_active(btn, on)
        if hasattr(self, 'task_list_layout'):
            self._update_urgent_badge()
        self.title_bar.btn_ai.setVisible(not on and not self._force_memo_mode)
        if on:
            self.memo_editor.setFocus()
            btn.setToolTip('업무 모드로 전환')
            ai_chat_window.host_hidden(self)         # 메모 모드에서는 AI 대화 창을 숨김
        else:
            self.input_name.setFocus()
            btn.setToolTip('메모 모드로 전환')
            if self.isVisible():
                ai_chat_window.host_shown(self)

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
                    row.setStyleSheet(f'background: {rgba(self._accent, 0.32)}; border-radius: 6px;')
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
                item.widget().hide()            # 지워질 때까지 겹쳐 그려지지 않게
                item.widget().deleteLater()
        set_current_accent(self._accent)        # 업무 줄이 이 메모 창 색의 포인트 색을 쓰도록
        set_current_note_font(self._font_key)   # 업무 줄이 이 창에서 고른 글꼴을 쓰도록
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
        def restore_scroll():
            try:
                sb.setValue(scroll_pos)
            except RuntimeError:            # 그 사이 창이 닫힘
                pass
        QTimer.singleShot(0, lambda: QTimer.singleShot(0, restore_scroll))
        self._update_urgent_badge()
        slide = getattr(self, '_pending_slide', None)
        self._pending_slide = None
        if slide:
            QTimer.singleShot(0, lambda: self._play_slide(*slide))

    # ── 급함 표시 (ui/motion.py) ─────────────────────────────────
    def _task_rows(self):
        return [self.task_list_layout.itemAt(i).widget() for i in range(self.task_list_layout.count())
                if isinstance(self.task_list_layout.itemAt(i).widget(), TaskRow)]

    def _tick_urgency(self):
        """30초마다: 남은 시간·숨쉬기 단계·기한 지남을 줄을 새로 만들지 않고 고친다."""
        from datetime import datetime
        now = datetime.now()
        for row in self._task_rows():
            try:
                row.apply_urgency(now)
            except RuntimeError:
                pass
        self._update_urgent_badge()

    def _update_urgent_badge(self):
        """제목 줄 배지: 지난 마감·오늘 마감 개수 (메모 모드 창에서는 숨김)."""
        from ui import motion
        badge = getattr(self.title_bar, 'badge', None)
        if badge is None:
            return
        if getattr(self, '_memo_mode', False):
            badge.set_counts(0, 0, False)
            return
        cfg = motion.load()
        over = today = 0
        for row in self._task_rows():
            u = getattr(row, '_urg', None) or {}
            if u.get('overdue'):
                over += 1
            elif u.get('days_left') == 0:
                today += 1
        badge.set_counts(over, today, cfg['enabled'])

    def _play_slide(self, task_id, pixmap, old_pos):
        """★를 누른 업무가 옛 자리에서 새 자리로 미끄러져 가는 모습 (0.32초)."""
        from PyQt5.QtCore import QPropertyAnimation, QEasingCurve
        from PyQt5.QtWidgets import QGraphicsOpacityEffect
        row = next((r for r in self._task_rows() if r.task.get('id') == task_id), None)
        if row is None or pixmap is None:
            return
        new_pos = row.pos()
        if abs(new_pos.y() - old_pos.y()) < 4:
            return
        ghost = QLabel(self.task_list_widget)
        ghost.setPixmap(pixmap)
        ghost.setAttribute(Qt.WA_TransparentForMouseEvents)
        ghost.setStyleSheet(f'background: {rgba(self._accent, 0.10)}; border-radius: 6px;')
        ghost.resize(pixmap.size() / max(1.0, pixmap.devicePixelRatio()))
        ghost.move(old_pos)
        ghost.show()
        ghost.raise_()
        hide = QGraphicsOpacityEffect(row)
        hide.setOpacity(0.0)
        row.setGraphicsEffect(hide)
        anim = QPropertyAnimation(ghost, b'pos', self)
        anim.setDuration(320)
        anim.setStartValue(old_pos)
        anim.setEndValue(new_pos)
        anim.setEasingCurve(QEasingCurve.OutCubic)

        def done():
            ghost.hide()
            ghost.deleteLater()
            try:
                row.setGraphicsEffect(None)
            except RuntimeError:
                pass
        anim.finished.connect(done)
        self._slide_anim = anim
        anim.start()

    def _set_motion(self, menu=None, **values):
        """급함 표시 설정 저장 (enabled / countdown_hours) → 열린 창 모두 바로 반영."""
        from ui import motion
        motion.save(**values)
        for win in (self._open_windows or [self]):
            try:
                win._tick_urgency()
            except RuntimeError:
                pass
        if menu:
            menu.close()

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

    def _ensure_doc_editor(self):
        """열려 있는 공문 작성 창(없으면 새로)을 돌려준다. 메모 창이 여러 개여도 창은 하나."""
        from document_editor import DocumentEditorWindow
        ed = DocumentEditorWindow._instance
        if ed is None or not (ed.isVisible() or ed._capturing):
            ed = DocumentEditorWindow(parent=None)
        self._doc_editor = ed
        return ed

    def _open_document_editor(self):
        try:
            ed = self._ensure_doc_editor()
            ed.show()
            ed.raise_()
            ed.activateWindow()
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

    def _show_monthly_calendar(self):
        dlg = getattr(self, '_calendar_dlg', None)
        if dlg is not None and dlg.isVisible():
            dlg._load_tasks()
            dlg.raise_()
            dlg.activateWindow()
            return
        # 메모 창(항상 위)에 딸린 창이면 Windows가 달력도 항상 위에 둔다 → 주인 없는 보통 창으로 연다
        dlg = NativeCalendarDialog(self._open_windows, None)
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
        # 순서·이름은 「…」 메뉴와 맞춘다 (자주 쓰는 기능 → 창 → 배경색·글자 크기 → 설정 → 도움말)
        return [
            Command('📋 오늘 한눈에 보기', self._open_dashboard, '', 'dashboard today 대시보드 마감'),
            Command('📸 공문 캡처', self._start_ai_capture, 'Ctrl+Shift+X', 'capture ocr 캡쳐 스크린샷 ai 공문 분석 요약'),
            Command('✏️ 새 업무 입력', new_task, '', 'task add 추가 할일'),
            Command('💬 AI 대화 열기/숨기기', lambda: target._toggle_ai_chat(), '', 'ai chat 채팅 대화 질문 클로드'),
            Command('🆕 새 메모장', new_memo, '', 'new memo 메모'),
            Command('✍️ 공문 작성', self._open_document_editor, '', 'document 기안 초안 회신'),
            Command('₩ 예산 관리', self._open_budget_window, '', 'budget money 돈 집행'),
            Command('📅 월간 달력', self._show_monthly_calendar, '', 'calendar 일정 캘린더'),
            Command('📜 지난 업무', lambda: self._show_history(0), '', 'history 지난 기록 삭제 되살리기 복원'),
            Command('🗂 작성 공문', lambda: self._show_history(1), '', 'document 공문 저장 기록 지난 기록'),
            Command('📌 항상 위에 ' + ('고정 끄기' if on_top else '고정'), self.toggle_always_on_top,
                    'Ctrl+Shift+F', 'pin top 핀'),
            Command('↕️ 메모장 접기/펼치기', self.toggle_shade, 'Ctrl+Shift+R', 'shade fold collapse'),
            Command('🌙 다크 모드 켜기/끄기', self._toggle_theme, '', 'dark light theme 테마 어둡게 밝게 설정'),
            Command('🛡️ 개인정보 보호 모드 ' + ('끄기' if privacy_enabled() else '켜기'),
                    self._toggle_privacy_mode, '', 'privacy 보안 마스킹 설정'),
            Command('🔑 AI 연결 (API 키)', self._open_ai_settings, '', 'ai api key opencode 오픈코드 claude 클로드 chatgpt 챗지피티 openai 설정'),
            Command('🤖 Claude 연결', self._open_claude_connect, '', 'claude mcp ai 연결 클로드 설정'),
            Command('💡 사용 설명서', self.show_help, '', 'help 도움말'),
            Command('⬇️ 업데이트 확인', lambda: updater.check_for_update_manual(self), '', 'update 버전 도움말'),
        ]

    def _open_command_palette(self):
        if self._force_memo_mode:
            return
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

    # ── 공문 캡처 (공문 작성 창의 「받은 공문」 칸에서 분석 → 제목·관련 번호·요약) ──
    def _start_ai_capture(self):
        try:
            ed = self._ensure_doc_editor()
        except Exception as e:
            QMessageBox.critical(self, '오류', f'공문 작성 창 열기 실패:\n{e}')
            return
        ed._start_capture(extra_hide=list(self._open_windows))

    def _task_list_window(self):
        """업무를 넣을 창: 자신이 업무 목록 모드면 자신, 아니면 열린 업무 목록 창."""
        if not self._memo_mode:
            return self
        for win in self._open_windows:
            if not getattr(win, '_memo_mode', True):
                return win
        return self

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

    # ── 「…」 메뉴 (ui/menu_style.py: 흰 바탕·선 아이콘·옅은 단축키) ─────────
    def _menu_add(self, menu, text, on_trigger, checked=None, shortcut='', icon=None):
        """메뉴 항목 추가. checked가 True/False면 체크 항목, shortcut은 오른쪽 끝에 옅게."""
        return menu_style.add_action(menu, text, on_trigger, icon=icon, shortcut=shortcut, checked=checked)

    @staticmethod
    def _menu_header(menu, text):
        return menu_style.add_header(menu, text)

    def _submenu(self, parent, title, min_width=None, icon=None):
        sub = menu_style.add_submenu(parent, title, icon=icon)
        if min_width:
            sub.setMinimumWidth(min_width + 2 * menu_style.SHADOW)
        return sub

    def _color_swatches(self, menu):
        """배경색 고르기: 동그란 색 견본, 지금 색은 바깥 고리로 표시."""
        color_widget = QWidget()
        color_widget.setStyleSheet('background: transparent;')
        grid = QGridLayout(color_widget)
        grid.setContentsMargins(12, 4, 12, 8)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        current = (self._bg_color or '').upper()
        for i, color in enumerate(PALETTE):
            cb = QPushButton()
            cb.setFixedSize(24, 24)
            cb.setCursor(Qt.PointingHandCursor)
            cb.setToolTip('이 색으로')
            ring = (f'2px solid {menu_style.TEXT.name()}' if color.upper() == current
                    else f'1px solid {tc("rgba(0,0,0,0.12)")}')
            cb.setStyleSheet(f"""
                QPushButton {{ background: {note_color(color)}; border-radius: 12px; border: {ring}; }}
                QPushButton:hover {{ border: 2px solid {tc("rgba(0,0,0,0.35)")}; }}
            """)
            cb.clicked.connect(lambda _, c=color: (self._apply_color(c), self.save_state(), menu.close()))
            grid.addWidget(cb, i // 6, i % 6)
        return color_widget

    def _style_submenu(self, menu, with_row_gap):
        """배경색·글자 크기 (이 메모장에만 적용). 줄 간격은 업무 목록 사이 간격이라 메모장에는 없다."""
        add = self._menu_add
        memo = self._submenu(menu, '배경색·글자 크기', icon='ph.palette')
        self._menu_header(memo, '배경색')
        wa = QWidgetAction(memo)
        wa.setDefaultWidget(self._color_swatches(menu))
        memo.addAction(wa)
        memo.addSeparator()
        refresh_installed_fonts()
        font_menu = self._submenu(memo, '글꼴', 200, icon='ph.text-t')
        for key, label in NOTE_FONTS:
            ok = note_font_available(key)
            a = add(font_menu, label if ok else f'{label} (설치 필요)',
                    (lambda k=key: self._set_font(k, menu)) if ok else (lambda k=key: self._font_missing(k, menu)),
                    checked=self._font_key == key)
            a.setFont(note_font(11, key))               # 글꼴 이름을 그 글꼴로 미리 보기
            if not ok:
                a.setToolTip('이 PC에 설치되어 있지 않아요 — 누르면 받는 곳과 설치 방법을 알려 드려요')
        size_menu = self._submenu(memo, '글자 크기', 110, icon='ph.text-aa')
        for label, val in [('기본', 1.0), ('크게', 1.1), ('더 크게', 1.2)]:
            add(size_menu, label, lambda v=val: self._set_scale(v, menu), checked=abs(self._scale - val) < 0.01)
        if with_row_gap:
            gap_menu = self._submenu(memo, '줄 간격', 110, icon='ph.rows')
            for label, val in [('0px', 0), ('5px', 5), ('10px', 10), ('15px', 15)]:
                add(gap_menu, label, lambda v=val: self._set_row_gap(v, menu), checked=self._row_gap == val)

    def _show_memo_menu(self, btn):
        """메모장(메모 전용 창)의 「…」 메뉴: 배경색·글자 크기 · 다크 모드 · 메모장 삭제 — 이것뿐."""
        menu = menu_style.make_menu(self)
        self._style_submenu(menu, with_row_gap=False)
        from db import get_setting as _gs
        from ui.theme import MODE as _mode
        self._menu_add(menu, '다크 모드 (다시 켜면 적용)', self._toggle_theme, checked=_gs('theme', _mode) == 'dark')
        menu.addSeparator()
        self._menu_add(menu, '메모장 삭제', self.delete_memo, icon='ph.trash')
        self._last_menu = menu  # 테스트·디버그용 참조
        menu.exec_(btn.mapToGlobal(btn.rect().bottomLeft()))

    def show_menu(self, btn):
        """「…」 메뉴: 자주 쓰는 기능 → 배경색·글자 크기 → 설정 → 도움말 → (삭제)."""
        if self._force_memo_mode:
            return self._show_memo_menu(btn)
        menu = menu_style.make_menu(self)
        add = self._menu_add

        # ── 자주 쓰는 기능 ──
        add(menu, '오늘 한눈에 보기', self._open_dashboard, icon='ph.squares-four')
        add(menu, '공문 캡처', self._start_ai_capture, shortcut='Ctrl+Shift+X', icon='ph.crop')
        add(menu, '빠른 찾기', self._open_command_palette, shortcut='Ctrl+K', icon='ph.magnifying-glass')
        if self._ai_chat_usable():
            add(menu, 'AI 대화 숨기기' if ai_chat_window.is_attached_to(self) else 'AI 대화 열기',
                self._toggle_ai_chat, icon='ph.chat-circle-dots')
        add(menu, '새 메모장', lambda: self.on_new(offset_from=self, memo_mode=True,
                                                on_toggle_hotkey=self._on_toggle_hotkey,
                                                on_shortcut_change=self._on_shortcut_change,
                                                get_shortcut_enabled=self._get_shortcut_enabled),
            icon='ph.note-blank')
        menu.addSeparator()

        # ── 배경색·글자 크기 (이 메모장에만 적용) ──
        self._style_submenu(menu, with_row_gap=True)

        # ── 설정 (앱 전체에 적용) ──
        settings = self._submenu(menu, '설정', 230, icon='ph.gear-six')
        alarm = self._submenu(settings, '마감 알림', 230, icon='ph.bell')
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

        from ui import motion as _motion
        _mcfg = _motion.load()
        urg = self._submenu(settings, '급함 표시 (움직임)', 230, icon='ph.lightning')
        self._menu_header(urg, _motion.describe(_mcfg))
        add(urg, '움직임 켜기', lambda: self._set_motion(menu, enabled=not _mcfg['enabled']), checked=_mcfg['enabled'])
        urg.addSeparator()
        self._menu_header(urg, '남은 시간 표시 (오늘 마감 시각이 있을 때)')
        for _h in _motion.HOUR_CHOICES:
            add(urg, '끄기' if _h == 0 else f'{_h}시간 전부터', lambda h=_h: self._set_motion(menu, countdown_hours=h),
                checked=_mcfg['countdown_hours'] == _h)

        sc_enabled = self._get_shortcut_enabled() if self._get_shortcut_enabled else True
        add(settings, '단축키 사용', lambda: self._toggle_shortcut(not sc_enabled), checked=sc_enabled)
        autostart_on = autostart_is_enabled()
        add(settings, '시작 시 자동실행', lambda: autostart_set(not autostart_on), checked=autostart_on)
        auto_up = updater.auto_enabled()
        add(settings, '자동 업데이트', lambda: updater.set_auto_enabled(not auto_up), checked=auto_up)
        settings.addSeparator()
        from db import get_setting as _gs
        from ui.theme import MODE as _mode
        from school_profile import privacy_enabled
        add(settings, '다크 모드 (다시 켜면 적용)', self._toggle_theme, checked=_gs('theme', _mode) == 'dark')
        add(settings, '개인정보 보호 모드', self._toggle_privacy_mode, checked=privacy_enabled())
        add(settings, 'AI 연결 (API 키)…', self._open_ai_settings, icon='ph.key')
        add(settings, 'Claude 연결…', self._open_claude_connect, icon='ph.link')
        settings.addSeparator()
        add(settings, '창 배치 기본값으로…', self._reset_layouts, icon='ph.arrows-counter-clockwise')

        # ── 도움말 ──
        help_menu = self._submenu(menu, '도움말', 150, icon='ph.question')
        add(help_menu, '사용 설명서', self.show_help, icon='ph.book-open')
        add(help_menu, '업데이트 확인', lambda: updater.check_for_update_manual(self), icon='ph.arrow-clockwise')

        # ── 메모장 삭제 (첫 메모장 제외) ──
        _all_ids = [w['id'] for w in get_all_windows()]
        if not _all_ids or self.window_id != min(_all_ids):
            menu.addSeparator()
            add(menu, '메모장 삭제', self.delete_memo, icon='ph.trash')

        self._last_menu = menu  # 테스트·디버그용 참조
        pos = btn.mapToGlobal(btn.rect().bottomLeft())
        menu.exec_(pos)


    # ── 입력칸 (메모 색에 맞춘 포인트 색) ───────────────────────────
    def _chip(self, kind, color):
        """테두리 없는 둥근 칸: 평소엔 옅은 바탕, 누르면 흰 바탕 + 포인트 색 테두리."""
        ac = self._accent
        return f"""
            {kind} {{
                background: {tc("rgba(0,0,0,0.055)")};
                border: 1px solid transparent;
                border-radius: 7px;
                padding: 2px 9px;
                color: {color};
            }}
            {kind}:hover {{ background: {tc("rgba(0,0,0,0.08)")}; }}
            {kind}:focus {{ background: {tc("rgba(255,255,255,0.75)")}; border: 1px solid {ac}; }}
        """

    def _recur_style(self, chosen: bool) -> str:
        return (self._chip('QComboBox', tct('#333') if chosen else tc('rgba(0,0,0,0.42)'))
                + "QComboBox::drop-down { border: none; width: 0px; }"
                  "QComboBox QAbstractItemView { font-family: 'Malgun Gothic'; font-size: 11pt; }")

    def _style_inputs(self):
        if not hasattr(self, 'input_name'):
            return
        ac = self._accent
        self.input_name.setStyleSheet(self._chip('QLineEdit', tct('#222')).replace('padding: 2px 9px', 'padding: 3px 11px'))
        _date_base = """
            QDateEdit::drop-down { width: 0px; border: none; background: transparent; }
            QDateEdit::up-button { width: 0px; }
            QDateEdit::down-button { width: 0px; }
        """
        self._field_style_date_empty = (self._chip('QDateEdit', tc('rgba(0,0,0,0.42)'))
                                        + 'QDateEdit { qproperty-alignment: AlignLeft; }' + _date_base)
        self._field_style_date_filled = (self._chip('QDateEdit', tct('#222'))
                                         + 'QDateEdit { qproperty-alignment: AlignCenter; }' + _date_base)
        self.input_date.setStyleSheet(self._field_style_date_filled if self._date_touched
                                      else self._field_style_date_empty)
        self.input_time.setStyleSheet(self._chip('QTimeEdit', tct('#222')) + f"""
            QTimeEdit {{ qproperty-alignment: AlignCenter; }}
            QTimeEdit:disabled {{ background: transparent; color: {tc("rgba(0,0,0,0.22)")}; }}
            QTimeEdit::up-button {{ width: 0px; }}
            QTimeEdit::down-button {{ width: 0px; }}
        """)
        self.input_recurrence.setStyleSheet(self._recur_style(self.input_recurrence.currentIndex() != 0))
        self.task_scroll.setStyleSheet(f"""
            QScrollArea {{ border: none; background: transparent; }}
            QScrollBar:vertical {{ width: 6px; background: transparent; }}
            QScrollBar::handle:vertical {{ background: {rgba(ac, 0.45)}; border-radius: 3px; }}
            QScrollBar::handle:vertical:hover {{ background: {rgba(ac, 0.7)}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)

    def _size_inputs(self):
        """칸 크기를 글자 크기에서 정한다 (고정 숫자 대신) — 배율·글꼴·글자 크기 설정이 달라도 안 잘리게."""
        row_h = self.input_name.sizeHint().height() + 6
        self.input_name.setFixedHeight(row_h)
        for w in (self.input_date, self.input_time, self.input_recurrence):
            w.setFixedHeight(row_h - 4)
        pad = 2 * 9 + 2 + 8          # 좌우 안쪽 여백(9px씩) + 테두리 + 여유
        fm = QFontMetrics(self.input_date.font())
        self.input_date.setFixedWidth(max(fm.horizontalAdvance('0000-00-00'), fm.horizontalAdvance('마감일')) + pad)
        self.input_time.setFixedWidth(QFontMetrics(self.input_time.font()).horizontalAdvance('00:00') + pad)
        fr = QFontMetrics(self.input_recurrence.font())
        self.input_recurrence.setFixedWidth(max(fr.horizontalAdvance(self.input_recurrence.itemText(i))
                                                for i in range(self.input_recurrence.count())) + pad)

    def _on_app_focus_changed(self, *_):
        try:
            self._sync_input_details()
        except RuntimeError:        # 닫혀서 정리된 창
            pass

    def _sync_input_details(self):
        """마감일·시각·반복 줄은 업무 추가 칸을 쓰고 있거나 값을 골랐을 때만 보인다."""
        if not hasattr(self, 'date_row_widget'):
            return
        if getattr(self, '_memo_mode', False):
            self.date_row_widget.hide()
            return
        fw = QApplication.focusWidget()
        in_form = fw is not None and any(fw is w or w.isAncestorOf(fw)
                                         for w in (self.task_input_row, self.date_row_widget))
        show = (in_form or bool(self.input_name.text().strip()) or self._date_touched
                or self.input_recurrence.currentIndex() != 0
                or (hasattr(self, '_cal_popup') and self._cal_popup.isVisible()))
        if show != self.date_row_widget.isVisible():
            self.date_row_widget.setVisible(show)

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
        self.title_bar.apply_style(title_color)
        self._accent = accent_for_note(shown)   # 이 메모 색에 맞춘 포인트 색 (입력칸·선택한 줄·스크롤바)
        self._style_inputs()
        if hasattr(self, 'task_list_layout') and self.task_list_layout.count() > 1:
            self._refresh_tasks()               # 업무 줄도 새 색으로
        if DARK:
            sep_color = QColor(shown).lighter(160).name()
        else:
            sep_color = QColor(shown).darker(118).name()
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

    def _show_history(self, tab=None):
        """지난 기록 창 (지난 업무 · 작성 공문). tab: 0 지난 업무, 1 작성 공문, None이면 마지막에 본 칸."""
        from ui.history_dialog import HistoryDialog

        def restore(r):
            add_task(self.window_id, r['name'], r['deadline'],
                     strikethrough=r.get('strikethrough', 0), priority=r.get('priority', 0),
                     recurrence=r.get('recurrence', ''), notes=r.get('notes', ''),
                     related_no=r.get('related_no', ''))
            delete_task_history(r['id'])
            self._refresh_tasks()
        HistoryDialog(self, self.window_id, restore, tab=tab).exec_()

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


    MANUAL_URL = 'https://github.com/justkevin2570-hash/SSNnote/blob/master/docs/manual.md'

    def show_help(self):
        """사용 설명서: 웹 페이지(docs/manual.md)를 기본 브라우저로 연다. 못 열면 주소를 보여준다."""
        from PyQt5.QtGui import QDesktopServices
        if QDesktopServices.openUrl(QUrl(self.MANUAL_URL)):
            return
        box = QMessageBox(self)
        box.setWindowTitle('사용 설명서')
        box.setTextFormat(Qt.RichText)
        box.setTextInteractionFlags(Qt.TextBrowserInteraction)
        box.setText('브라우저를 열지 못했습니다. 아래 주소를 브라우저에 붙여 넣어 주세요.<br><br>'
                    f'<a href="{self.MANUAL_URL}">{self.MANUAL_URL}</a>')
        copy_btn = box.addButton('주소 복사', QMessageBox.ActionRole)
        box.addButton('닫기', QMessageBox.RejectRole)
        box.exec_()
        if box.clickedButton() is copy_btn:
            QApplication.clipboard().setText(self.MANUAL_URL)

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
            ai_chat_window.host_shown(self)
        else:
            self.expanded_height = self.height()
            self.collapsed = True           # resizeEvent보다 먼저 설정
            self.content.hide()
            self.setFixedHeight(TITLE_BAR_HEIGHT)
            ai_chat_window.host_hidden(self)         # 제목만 남게 접으면 AI 대화 창도 숨김
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
            self.title_bar.set_active(btn, True, 'ph.push-pin-fill')
            btn.setToolTip('항상 위 켜짐')
        else:
            self.title_bar.set_active(btn, False, 'ph.push-pin-bold')
            btn.setToolTip('항상 위 꺼짐')
        if ai_chat_window.is_attached_to(self):
            ai_chat_window.instance().attach(self, remember=False)   # 항상 위 설정을 AI 창에도

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
        ai_chat_window.follow(self)
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
        ai_chat_window.follow(self)

    def _apply_scale_to_inputs(self):
        s = self._scale
        self.input_name.setFont(note_font(int(11 * s), self._font_key))
        self.input_date.setFont(QFont('Malgun Gothic', int(11 * s)))
        self.input_time.setFont(QFont('Malgun Gothic', int(11 * s)))
        self.input_recurrence.setFont(QFont('Malgun Gothic', int(10 * s)))
        self._size_inputs()
        self.title_bar.set_icon_scale(s)
        self.memo_editor.setFont(note_font(int(12 * s), self._font_key))

    def _font_missing(self, key, menu=None):
        """설치되지 않은 글꼴을 골랐을 때: [받기](공식 사이트) · [설치 확인](설치했으면 바로 바꿈) · [닫기]."""
        import html as _html
        from ui.common import NOTE_FONT_DOWNLOAD, NOTE_FONT_FILES, REATIC_SITE
        from PyQt5.QtGui import QDesktopServices
        if menu:
            menu.close()
        label = dict(NOTE_FONTS).get(key, key)
        url = NOTE_FONT_DOWNLOAD.get(key)
        if url:
            files = ''.join(f'<br>&nbsp;&nbsp;&nbsp;· <b>{f}.ttf</b>' for f in NOTE_FONT_FILES.get(key, ()))
            body = (f'<b>{_html.escape(label)}</b>가 이 PC에 설치되어 있지 않아요.<br><br>'
                    '1. <b>[받기]</b> — 공식 사이트에서 리틱체.zip을 받아요.<br>'
                    f'2. 압축을 풀고 TTF 폴더 안의 아래 파일을 더블클릭 → [설치]{files}<br>'
                    '3. <b>[설치 확인]</b> — 설치됐으면 바로 이 글꼴로 바뀌어요.<br><br>'
                    f'받기가 안 되면 <a href="{REATIC_SITE}">리틱인더스트리 공식 사이트</a>에서 받으세요.<br>'
                    '<span style="color:#64748B">리틱체는 글꼴 파일 재배포가 금지되어 있어 앱에 넣지 않았어요.</span>')
        else:
            body = (f'<b>{_html.escape(label)}</b>을(를) 이 PC에서 찾지 못했어요.<br>'
                    'Windows 설정 → 개인 설정 → 글꼴에서 설치되어 있는지 확인해 주세요.')
        dlg = QDialog(self)
        dlg.setWindowTitle('글꼴 설치 필요')
        v = QVBoxLayout(dlg)
        v.setContentsMargins(20, 18, 20, 14)
        v.setSpacing(12)
        text = QLabel(body)
        text.setTextFormat(Qt.RichText)
        text.setWordWrap(True)
        text.setOpenExternalLinks(True)
        text.setMinimumWidth(px(430))
        text.setFont(QFont('Malgun Gothic', 10))
        v.addWidget(text)
        status = QLabel('')
        status.setWordWrap(True)
        status.setStyleSheet('color: #B45309;')
        status.hide()
        v.addWidget(status)
        row = QHBoxLayout()
        row.addStretch()
        result = {'ok': False}

        def get():
            QDesktopServices.openUrl(QUrl(url))
            status.setText('받은 파일을 설치한 뒤 [설치 확인]을 눌러 주세요.')
            status.show()

        def check():
            refresh_installed_fonts()
            if note_font_available(key):
                result['ok'] = True
                dlg.accept()
            else:
                status.setText('⚠ 아직 찾지 못했어요. 설치를 마쳤는지 확인해 주세요.')
                status.show()
        buttons = ([('받기', get)] if url else []) + [('설치 확인', check), ('닫기', dlg.reject)]
        dlg._buttons = {}
        for i, (t, fn) in enumerate(buttons):
            b = QPushButton(t)
            b.setMinimumWidth(px(84))
            b.clicked.connect(fn)
            b.setDefault(i == 0)
            row.addWidget(b)
            dlg._buttons[t] = b
        v.addLayout(row)
        self._font_dialog = dlg                         # 테스트에서 버튼을 누를 수 있게
        dlg.exec_()
        if result['ok']:
            self._set_font(key)
        return result['ok']

    def _reset_layouts(self):
        """예산·공문 작성·AI 대화 창의 위치·크기·보던 칸을 처음 상태로 (업무·예산·공문 내용은 그대로)."""
        ans = QMessageBox.question(
            self, '창 배치 기본값으로',
            '예산 관리 · 공문 작성 · AI 대화 창의 위치와 크기, 보던 칸을 처음 상태로 되돌릴까요?\n'
            '업무 · 예산 · 공문 내용은 그대로예요.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return
        import os
        import budget
        from document_editor import DocumentEditorWindow
        if getattr(budget, '_instance', None) is not None:
            budget._instance.reset_layout()
        else:
            try:
                os.remove(budget._STATE_FILE)
            except OSError:
                pass
        ed = DocumentEditorWindow._instance
        if ed is not None:
            ed.reset_layout()
        else:
            try:
                os.remove(DocumentEditorWindow._STATE_FILE)
            except OSError:
                pass
        ai_chat_window.reset_layout()

    def _set_font(self, key, menu=None):
        """이 창의 글꼴 바꾸기 (업무 이름·메모·D-day·업무 추가 칸, 메모장 본문)."""
        self._font_key = key if key in dict(NOTE_FONTS) else DEFAULT_NOTE_FONT
        set_window_font(self.window_id, '' if self._font_key == DEFAULT_NOTE_FONT else self._font_key)
        self._apply_scale_to_inputs()
        self._refresh_tasks()
        if menu:
            menu.close()

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
            if not self._force_memo_mode:
                QTimer.singleShot(600, self._auto_open_ai_chat)
        elif self._ai_chat_usable():
            ai_chat_window.host_shown(self)

    def hideEvent(self, e):
        super().hideEvent(e)
        ai_chat_window.host_hidden(self)

    # ── AI 대화 창 (ui/ai_chat_window.py — 앱 전체에 하나, 업무창 옆에 붙음) ──
    def _ai_chat_usable(self):
        return not (self._force_memo_mode or self._memo_mode or self.collapsed)

    def _toggle_ai_chat(self):
        if self._ai_chat_usable():
            ai_chat_window.toggle(self)

    def _auto_open_ai_chat(self):
        try:
            if self.isVisible() and self._ai_chat_usable():
                ai_chat_window.auto_open(self)
        except RuntimeError:
            pass

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
        ai_chat_window.host_closed(self)
        if not getattr(self, '_deleted', False):
            self.save_state()
            if self._memo_save_timer and self._memo_save_timer.isActive():
                self._memo_save_timer.stop()
                self._flush_memo_save()
        e.accept()
