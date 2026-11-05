"""ui/task_row.py — 업무 목록의 한 줄(TaskRow)과 그 안의 보조 위젯들 (메모 버튼, 인라인 메모 편집기 등)."""
import os
import sys
import html
import json
import subprocess
import ctypes
from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QMenu, QAction, QMessageBox, QDialog, QCalendarWidget, QPlainTextEdit, QSizePolicy, QGraphicsOpacityEffect, QTimeEdit,
    QCheckBox
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QPoint, QPointF, QSize, pyqtSignal, QMimeData
from PyQt5.QtGui import QFont, QFontMetrics, QColor, QPainter, QTextCharFormat, QPalette, QTextOption, QTextLayout, QIcon, QTextCursor
from db import (update_task,
                set_task_priority, set_task_recurrence, get_task_notes, set_task_notes,
                get_task_related_no, set_task_related_no,
                get_task_attachments, set_task_attachments)
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, _normalize_doc_number
from .common import _base_path, _svg_icon, MAT, mat_font, MI, mi_font, pr_font
from .dates import calc_dday
from .widgets import AddDateButton, CustomCalendarWidget, DdayLabel, RotatedMenuButton
from .dialogs import CustomRecurrenceDialog
from .theme import T

class _AutoHeightEdit(QPlainTextEdit):
    selectRequested = pyqtSignal()
    editFinished = pyqtSignal()
    heightChanged = pyqtSignal()
    """콘텐츠 높이가 바뀌었을 때 emit — 행(TaskRow)이 높이를 따라가도록."""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.document().setDocumentMargin(0)
        opt = QTextOption()
        opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        self.document().setDefaultTextOption(opt)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._fixed_height = 28
        self._preferred_width = 40
        self._v_pad = 0
        self._last_y = 0.0
        self._editing = False
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(500)
        self._resize_timer.timeout.connect(self._apply_size)
        # 줄바꿈 경계(높이 변화)는 즉시 반영해 접힌 텍스트 방지,
        # 폭/전체 크기 변경은 500ms 디바운스 (IME 조합 중 레이아웃 보호)
        self.document().contentsChanged.connect(self._on_contents_changed)
        self._apply_size()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Escape):
            self.clearFocus()
            self.editFinished.emit()
        else:
            super().keyPressEvent(e)

    def viewportEvent(self, e):
        if e.type() == QEvent.MouseButtonPress and e.button() == Qt.LeftButton:
            if not self.hasFocus():
                # 비편집 상태 클릭 → 행 선택 (기존 동작 유지)
                self.selectRequested.emit()
                return True
            # 편집 중(포커스 보유) 클릭 → 소비하지 않고 기본 처리로 → 커서가 클릭 위치로 이동
        if e.type() == QEvent.MouseButtonDblClick and e.button() == Qt.LeftButton:
            self.setFocus()
        return super().viewportEvent(e)

    def _begin_edit(self):
        self.setFocus()
        self._editing = True
        self.setLineWrapMode(QPlainTextEdit.NoWrap)   # 편집 중 한 줄 — 커서가 항상 오른쪽 끝을 따라감
        self.setMinimumWidth(160)                     # 좁은 박스(48px)로 텍스트가 잘리는 문제 방지
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.End)          # 가장 오른쪽 텍스트의 오른쪽에 커서
        self.setTextCursor(cursor)
        self.ensureCursorVisible()
        self._apply_height()                          # 한 줄 높이로 즉시 정리

    def mouseDoubleClickEvent(self, e):
        self._begin_edit()

    def _on_contents_changed(self):
        self._apply_height()        # 줄바꿈 경계에서 즉시 높이 확보 (접힌 텍스트 방지)
        self._resize_timer.start()  # 폭/전체 크기는 500ms 디바운스

    def focusOutEvent(self, e):
        if self._resize_timer.isActive():
            self._resize_timer.stop()
        self._editing = False
        self.setLineWrapMode(QPlainTextEdit.WidgetWidth)  # 표시 모드 줄바꿈 복귀
        self.setMinimumWidth(0)     # 편집 중 최소 폭 해제
        self._apply_size()
        super().focusOutEvent(e)

    def sizeHint(self):
        return QSize(self._preferred_width, self._fixed_height)

    def _apply_height(self):
        old_h = self._fixed_height
        self._update_height()
        # 같은 값이면 재배치가 일어나지 않으므로 항상 적용해도 안전.
        # 1줄 업무명처럼 높이가 28→28로 안 바뀌는 경우에도 세로 가운데 보정(v_pad)이 적용돼야 정렬이 맞는다.
        if self._last_y > 0:  # 폭 미지정(레이아웃 전)이면 적용 보류 — resizeEvent에서 다시 계산
            self.setMinimumHeight(self._fixed_height)
            self.setMaximumHeight(self._fixed_height)
            self.setViewportMargins(0, self._v_pad, 0, self._v_pad)
        if self._fixed_height != old_h:
            self.heightChanged.emit()
            self.updateGeometry()

    def _apply_size(self):
        old_w = self._preferred_width
        self._apply_height()
        self._update_width()
        # 폭 변경 시에만 레이아웃 통지
        if self._preferred_width != old_w:
            self.updateGeometry()

    def _update_width(self):
        fm = QFontMetrics(self.font())
        w = fm.horizontalAdvance(self.toPlainText() or ' ') + 2
        self._preferred_width = max(w, 40)

    def changeEvent(self, e):
        super().changeEvent(e)
        if e.type() == QEvent.FontChange:
            self._apply_size()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._apply_size()

    def _update_height(self):
        width = self.viewport().width()
        if width <= 0:
            self._last_y = 0.0
            return
        text = self.toPlainText() or ' '
        layout = QTextLayout(text, self.font())
        opt = QTextOption()
        opt.setWrapMode(
            QTextOption.NoWrap if self._editing
            else QTextOption.WrapAtWordBoundaryOrAnywhere
        )
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
        self._last_y = y
        self._v_pad = max(0, (fixed_h - int(y)) // 2)

class _NoteButton(QPushButton):
    """업무 메모 펼치기/접기 토글 버튼. 접힘 ▼ / 펼침 ▲."""
    _CODEPOINT_COLLAPSED = '\u25bc'  # ▼
    _CODEPOINT_EXPANDED  = '\u25b2'  # ▲

    def __init__(self, scale=1.0, glyph_offset=4):
        super().__init__()
        self._color_normal = QColor('#555555')
        self._color_hover  = QColor(T.ACCENT)
        self._hovered      = False
        self._has_note     = False
        self._overdue      = False
        self._expanded     = False
        self._glyph_offset = glyph_offset
        self.setFlat(True)
        sz = int(22 * scale)
        self.setFixedSize(sz, sz)
        self.setFont(pr_font(int(10 * scale)))
        self.setCursor(Qt.PointingHandCursor)

    def set_note_color(self, has_note: bool, overdue: bool = None):
        self._has_note = has_note
        if overdue is not None:
            self._overdue = overdue
        if self._overdue:
            self._color_normal = QColor(102, 102, 102, 80)
        else:
            self._color_normal = QColor(85, 85, 85, 255) if has_note else QColor(85, 85, 85, 80)
        self.update()

    def set_expanded(self, expanded: bool):
        self._expanded = expanded
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
        p.setFont(self.font())
        p.setPen(self._color_hover if self._hovered else self._color_normal)
        rect = self.rect()
        rect.translate(0, self._glyph_offset)
        p.drawText(rect, Qt.AlignCenter,
                   self._CODEPOINT_EXPANDED if self._expanded else self._CODEPOINT_COLLAPSED)
        p.end()


class _ElideButton(QPushButton):
    """폭을 넘는 텍스트를 오른쪽 '…'로 잘라내는 버튼 (붙임 파일명 표시용)."""

    def __init__(self, text=''):
        super().__init__(text)
        self._hovered = False
        self.setFlat(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)

    def enterEvent(self, e):
        self._hovered = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hovered = False
        self.update()
        super().leaveEvent(e)

    def minimumSizeHint(self):
        sz = super().minimumSizeHint()
        return QSize(0, sz.height())

    def paintEvent(self, e):
        p = QPainter(self)
        p.setFont(self.font())
        p.setPen(QColor(T.ACCENT) if self._hovered else QColor('#333333'))
        fm = self.fontMetrics()
        elided = fm.elidedText(self.text(), Qt.ElideRight, max(0, self.width() - 4))
        p.drawText(self.rect().adjusted(2, 0, -2, 0), Qt.AlignLeft | Qt.AlignVCenter, elided)
        p.end()


class _RefLineEdit(QLineEdit):
    """Pretendard 입력 텍스트 + Pretendard placeholder를 직접 그리는 관련번호 입력줄.
    (QLineEdit::placeholder 스타일시트는 Windows 네이티브 스타일에서 무시되므로 직접 페인트)"""

    def __init__(self, ph_text='', ph_font=None, parent=None):
        super().__init__(parent)
        self._ph_text = ph_text
        self._ph_font = ph_font
        self.setPlaceholderText('')

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._ph_text and not self.text():
            p = QPainter(self)
            p.setRenderHint(QPainter.TextAntialiasing)
            p.setFont(self._ph_font)
            p.setPen(self.palette().color(QPalette.PlaceholderText))
            p.drawText(self.rect().adjusted(7, 0, -7, 0),
                       Qt.AlignVCenter | Qt.AlignLeft, self._ph_text)
            p.end()


class _InlineNoteEdit(QWidget):
    """업무 행 아래 펼쳐지는 2단 메모 패널. 상단: 관련번호 1줄 + 복사 버튼, 하단: 메모 에디터."""
    HEIGHT = 160       # 기본(최소) 높이
    MAX_HEIGHT = 280    # 상한: 도달 시 에디터 내부 세로 스크롤
    closeRequested = pyqtSignal()
    relatedNoChanged = pyqtSignal()

    def __init__(self, parent=None, scale=1.0):
        super().__init__(parent)
        # 높이는 TaskRow가 내용 기반으로 조절(setGeometry). min/max만 지정.
        self.setMinimumHeight(self.HEIGHT)
        self.setMaximumHeight(self.MAX_HEIGHT)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        top = QHBoxLayout()
        top.setSpacing(6)
        self.ref_edit = _RefLineEdit('공문 관련번호 적어두기', pr_font(int(11 * scale)))
        _ref_font = pr_font(int(12 * scale))
        _ref_font.setStretch(100)
        _ref_font.setLetterSpacing(QFont.PercentageSpacing, 100.0)
        self.ref_edit.setFont(_ref_font)
        self.ref_edit.setStyleSheet(f"""
            QLineEdit {{
                background: rgba(255,255,255,0.6);
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
            }}
            QLineEdit:focus {{ border: 1px solid {T.ACCENT_HOVER}; }}
        """)
        self.ref_edit.textChanged.connect(self.relatedNoChanged)
        self.ref_edit.installEventFilter(self)
        top.addWidget(self.ref_edit, 1)

        self.btn_copy = QPushButton('복사')
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setFixedWidth(56)
        _btn_font = pr_font(int(12 * scale))
        self.btn_copy.setFont(_btn_font)
        self.btn_copy.setFixedHeight(self.ref_edit.sizeHint().height())
        self.btn_copy.setStyleSheet(f"""
            QPushButton {{
                background: rgba(255,255,255,0.6);
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: rgba(212,184,0,0.25); }}
            QPushButton:pressed {{ background: rgba(212,184,0,0.42); padding: 2px 0 0 2px; }}
        """)
        self.btn_copy.clicked.connect(self._copy_related_no)
        top.addWidget(self.btn_copy)

        self.btn_shot = QPushButton('스크린샷')
        self.btn_shot.setCursor(Qt.PointingHandCursor)
        # 고정 80px은 Pretendard 4글자(68px@12pt) + 패딩/보더(14px)보다 좁아 잘림.
        # 실제 폰트 메트릭으로 폭 계산 → 글자 크기(1.0~1.2)·DPI 배율과 무관하게 항상 맞음.
        self.btn_shot.setFont(_btn_font)
        _fm = self.btn_shot.fontMetrics()
        self.btn_shot.setFixedWidth(_fm.horizontalAdvance('스크린샷') + 18)  # 14(QSS 패딩/보더) + 4(여유)
        self.btn_shot.setFixedHeight(self.ref_edit.sizeHint().height())  # 세로 = 복사와 동일
        self.btn_shot.setStyleSheet(f"""
            QPushButton {{
                background: rgba(255,255,255,0.6);
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: rgba(212,184,0,0.25); }}
            QPushButton:pressed {{ background: rgba(212,184,0,0.42); padding: 2px 0 0 2px; }}
        """)
        self.btn_shot.clicked.connect(self._start_inline_capture)
        top.addWidget(self.btn_shot)
        outer.addLayout(top)

        self.editor = QPlainTextEdit()
        self.editor.setPlaceholderText('esc 또는 ctrl+enter 을 누르면 접어집니다.')
        self.editor.setFont(pr_font(int(11 * scale)))
        self.editor.document().setDocumentMargin(4)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setStyleSheet(f"""
            QPlainTextEdit {{
                background: rgba(255,255,255,0.6);
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                color: #333;
                padding: 2px 4px;
            }}
            QPlainTextEdit:focus {{ border: 1px solid {T.ACCENT_HOVER}; }}
        """)
        self.editor.installEventFilter(self)
        outer.addWidget(self.editor, 1)

    def _doc_content_height(self):
        """현재 폭 기준 문서 콘텐츠 높이: 줄 수 × 줄 간격 + 문서 여백.
        (document().size()는 블록 수를 돌려줘서 부정확 — lineCount()는 래핑까지 반영)"""
        doc = self.editor.document()
        lines = doc.lineCount()
        if lines <= 0:
            return None
        return lines * self.editor.fontMetrics().lineSpacing() + 8  # 문서여백 4×2

    def desired_height(self):
        """현재 내용(관련번호줄 + 메모)에 필요한 패널 높이. HEIGHT~MAX_HEIGHT 범위.

        MAX_HEIGHT 초과 시 에디터 내부 세로 스크롤바(기본 AsNeeded)가 생긴다.
        레이아웃 미완료(첫 페인트 전)면 None — 호출 측에서 건너뛴다.
        """
        top_h = self.ref_edit.sizeHint().height()
        doc_h = self._doc_content_height()
        if doc_h is None:
            return None
        editor_h = doc_h + 8  # 여유(테두리 2 + QSS 패딩 4 + 진동 방지 2)
        needed = 8 + top_h + 6 + editor_h + 8
        return min(max(self.HEIGHT, needed), self.MAX_HEIGHT)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                self.closeRequested.emit()
                return True
            if event.modifiers() & Qt.ControlModifier and event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.closeRequested.emit()
                return True
        return super().eventFilter(watched, event)

    def _copy_related_no(self):
        text = self.ref_edit.text().strip()
        if not text:
            return
        mime = QMimeData()
        mime.setHtml(
            "<span style='font-family:굴림체; font-size:12pt; letter-spacing:0;'>"
            + html.escape(text) + "</span>"
        )
        mime.setText(text)
        QApplication.clipboard().setMimeData(mime)

    # ── 캡처 → 관련번호 자동 입력 ──────────────────────────────
    def _start_inline_capture(self):
        self._capture_win = self.window()
        if self._capture_win is not None:
            self._capture_win.hide()
        QApplication.processEvents()
        QTimer.singleShot(250, self._show_inline_capture_overlay)

    def _show_inline_capture_overlay(self):
        screenshot = grab_fullscreen()
        self._overlay = ScreenCaptureOverlay(screenshot)
        self._overlay.region_captured.connect(self._on_inline_capture_complete)
        self._overlay.cancelled.connect(self._restore_inline_window)
        self._overlay.show()
        self._overlay.activateWindow()
        self._overlay.setFocus()

    def _restore_inline_window(self):
        if getattr(self, '_capture_win', None) is not None:
            self._capture_win.show()
            self._capture_win.raise_()
            self._capture_win.activateWindow()
            try:
                hwnd = int(self._capture_win.winId())
                ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                pass
            self._capture_win = None

    def _on_inline_capture_complete(self, pixmap):
        if hasattr(self, '_overlay'):
            self._overlay.region_captured.disconnect()  # 중복 호출 방지
        self._restore_inline_window()

        try:
            import winrt  # noqa: F401
        except ImportError:
            QMessageBox.warning(
                self, 'OCR 불가',
                'Windows OCR(winrt) 모듈이 없습니다.\n\n'
                '① 터미널에서 pip install winsdk 를 실행하거나\n'
                '② 공문번호를 직접 입력해 주세요.'
            )
            return

        def _after_ocr(text):
            if text:
                normalized = _normalize_doc_number(text)
                if normalized:
                    self.ref_edit.setText(normalized)
            else:
                QMessageBox.information(self, '알림', 'OCR 텍스트를 인식하지 못했습니다.')

        def _on_error(msg):
            QMessageBox.warning(self, 'OCR 오류', msg)

        self._ocr_worker = run_ocr(pixmap, _after_ocr, _on_error)

    def focus_editor(self):
        self.editor.setFocus()

    def setPlainText(self, text):
        self.editor.setPlainText(text)

    def toPlainText(self):
        return self.editor.toPlainText()

    @property
    def textChanged(self):
        return self.editor.textChanged


class TaskRow(QWidget):
    def __init__(self, task, on_delete, on_update, scale=1.0, expanded=False, on_toggle=None,
                 on_select=None, on_navigate=None, on_clear=None,
                 attach_expanded=False, on_attach_toggle=None):
        super().__init__()
        # 행이 남는 공간을 먹고 세로로 늘어나지 않게 (여유 공간은 addStretch가 흡수)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.task      = task
        self.on_update = on_update
        self._scale    = scale
        self._expanded = expanded
        self._on_toggle = on_toggle
        self._on_select = on_select
        self._on_navigate = on_navigate
        self._on_clear = on_clear
        self._attach_expanded_init = attach_expanded
        self._on_attach_toggle = on_attach_toggle
        self._selected = False
        self._hovered = False
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setObjectName('TaskRow')
        self.setFocusPolicy(Qt.ClickFocus)
        self._apply_row_style()

        self._outer = QVBoxLayout(self)
        self._outer.setContentsMargins(9, 0, 8, 0)
        self._outer.setSpacing(0)

        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        self._outer.addLayout(layout)

        # 별 버튼 (우선순위)
        _is_starred = bool(task.get('priority', 0))
        self._star_sz = int(22 * scale)
        _star_pt = int(15 * scale)
        self._star_svg = os.path.join(_base_path(), 'assets', 'ic_fluent_star_emphasis_20_filled.svg')
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
                dday_color = '#666'
            elif dday == 'D-day' or '시간' in dday:
                dday_color = T.DANGER
                urgent = True
            elif dday == '날짜오류':
                dday_color = '#aaa'
            else:
                days_left = int(dday[2:])
                if days_left <= 5:
                    dday_color = T.DANGER
                    urgent = True
                else:
                    dday_color = '#333'
            if overdue:
                suffix = f'기한 지남({dday})'
            else:
                suffix = f'({dday})'
        else:
            dday_color = '#333'
            suffix = ''

        if overdue:
            _eff = QGraphicsOpacityEffect()
            _eff.setOpacity(0.0)
            self.btn_star.setGraphicsEffect(_eff)
            self.btn_star.setEnabled(False)

        self._strikethrough = bool(task.get('strikethrough', 0))
        name_color = '#666' if overdue else '#111'

        self.name_edit = _AutoHeightEdit(task['name'])
        self.name_edit.setFocusPolicy(Qt.NoFocus)
        # 파일 드롭이 name_edit(텍스트 위젯)에 가로채지고 부모 TaskRow로 버블업되도록 끔
        self.name_edit.setAcceptDrops(False)
        _base_pt = 12 * scale
        name_font = pr_font(12)
        name_font.setPointSizeF(_base_pt)  # 기한 지나도(overdue) 크기 유지 — 축소하지 않음
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
                border-bottom: 1px solid {T.ACCENT};
            }}
        """)
        if self._strikethrough:
            font = self.name_edit.font()
            font.setStrikeOut(True)
            self.name_edit.setFont(font)
        self.name_edit.installEventFilter(self)
        self.name_edit.heightChanged.connect(self._sync_row_height)
        self.name_edit.selectRequested.connect(self._request_select)
        self.name_edit.editFinished.connect(self.setFocus)

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
            dday_lbl.clicked.connect(self._show_date_picker)
            dday_font = pr_font(12)
            dday_font.setBold((urgent or _is_starred) and not overdue)
            dday_font.setPointSizeF(_base_pt)  # 기한 지나도(overdue) 크기 유지 — 축소하지 않음
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

        _has_note = bool(task.get('notes', '') or task.get('related_no', ''))
        self._overdue = overdue
        self.btn_note = _NoteButton(scale=scale)
        self.btn_note.set_note_color(_has_note, overdue)
        self.btn_note.setToolTip('업무 메모 (펼치기/접기)')
        self.btn_note.clicked.connect(self._toggle_note)
        self.btn_note.set_expanded(expanded)

        # ── 붙임파일 버튼 (클립 아이콘) ─────────────────────────
        self.btn_attach = QPushButton()
        self.btn_attach.setCursor(Qt.PointingHandCursor)
        self.btn_attach.setToolTip('붙임파일 목록 (펼치기/접기)')
        self.btn_attach.setFlat(True)
        self.btn_attach.setFixedSize(round(19 * scale), round(19 * scale))
        self.btn_attach.setFont(pr_font(round(11 * scale)))
        self.btn_attach.setStyleSheet(
            'QPushButton { border: none; color: #555555; background: transparent; padding: 0; }'
            f'QPushButton:pressed {{ color: {T.ACCENT}; }}'
        )
        _att_icon = os.path.join(_base_path(), 'assets', 'clip.png')
        self._attach_icon_ok = os.path.exists(_att_icon)
        if self._attach_icon_ok:
            self.btn_attach.setIcon(QIcon(_att_icon))
            self.btn_attach.setIconSize(QSize(round(15 * scale), round(15 * scale)))
        else:
            self.btn_attach.setText('📎')
        self.btn_attach.clicked.connect(self._toggle_attach)
        self.btn_attach.hide()

        # ── 붙임파일 목록 상태 ──────────────────────────────
        self._attachments = self._load_attachments()
        self._attach_expanded = self._attach_expanded_init and bool(self._attachments)
        self._attach_container = None
        self._drag_hover = False
        self.setAcceptDrops(True)
        self.installEventFilter(self)
        self._refresh_attach_btn()
        if self._attach_expanded:
            self._rebuild_attach_list()
            self._sync_row_height()

        name_note_layout = QHBoxLayout()
        name_note_layout.setContentsMargins(0, 0, 0, 0)
        name_note_layout.setSpacing(0)
        name_note_layout.addWidget(self.name_edit, 0, Qt.AlignVCenter)
        name_note_layout.addWidget(self.btn_note, 0, Qt.AlignTop)
        name_note_layout.addWidget(self.btn_attach, 0, Qt.AlignVCenter)
        layout.addLayout(name_note_layout)
        layout.addStretch(1)
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
        layout.addWidget(btn_menu, 0, Qt.AlignVCenter)

        # 메모 에디터 (▼ 클릭 시 행 높이 불변, 행 아래 예약 공간 위에 오버레이로 펼침)
        self.note_editor = _InlineNoteEdit(scale=self._scale)
        self.note_editor.closeRequested.connect(self._collapse_note)
        self._note_save_timer = QTimer(self)
        self._note_save_timer.setSingleShot(True)
        self._note_save_timer.timeout.connect(self._flush_note)
        self.note_editor.textChanged.connect(lambda: self._note_save_timer.start(500))
        self._note_height_timer = QTimer(self)
        self._note_height_timer.setSingleShot(True)
        self._note_height_timer.setInterval(60)  # IME 조합 중 깜빡임 완화
        self._note_height_timer.timeout.connect(self._update_note_editor_height)
        self.note_editor.textChanged.connect(lambda: self._note_height_timer.start())
        self._ref_save_timer = QTimer(self)
        self._ref_save_timer.setSingleShot(True)
        self._ref_save_timer.timeout.connect(self._flush_ref)
        self.note_editor.relatedNoChanged.connect(lambda: self._ref_save_timer.start(500))
        self.note_editor.hide()
        self._note_placeholder = None
        self.destroyed.connect(self.note_editor.deleteLater)
        if expanded:
            self._expanded = True
            self.btn_note.set_expanded(True)
            self.note_editor.setPlainText(get_task_notes(self.task['id']))
            self.note_editor.ref_edit.setText(get_task_related_no(self.task['id']))

    def _set_row_highlight(self, on: bool):
        self._hovered = on
        self._apply_row_style()

    def _apply_row_style(self):
        if getattr(self, '_drag_hover', False):
            style = 'QWidget#TaskRow { background: rgba(212,184,0,0.16); border-radius: 4px; }'
        elif self._selected:
            style = 'QWidget#TaskRow { background: rgba(212,184,0,0.25); border-radius: 4px; }'
        elif self._hovered:
            style = 'QWidget#TaskRow { background: rgba(0,0,0,18); }'
        else:
            style = 'QWidget#TaskRow { background: transparent; }'
        self.setStyleSheet(style)

    def set_selected(self, selected: bool):
        self._selected = selected
        self._apply_row_style()

    def _request_select(self):
        self.setFocus()
        if self._on_select:
            self._on_select(self.task['id'])

    def _sync_row_height(self):
        """업무명 줄바꿈 높이 변화를 행 높이에 즉시 반영.
        (sizeHint 캐시 갱신이 행까지 늦게 전파되는 문제를 setFixedHeight로 우회)"""
        need = max(self.name_edit._fixed_height, 31)  # 31 = 우측 버튼군(메뉴) 높이
        need += self._attach_area_height()
        if self.height() != need:
            self.setFixedHeight(need)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._request_select()
            e.accept()  # 부모(eventFilter)로 버블업 차단 → 선택 직후 클리어 방지
        super().mousePressEvent(e)

    def _start_rename(self):
        self.name_edit._begin_edit()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F2:
            self._start_rename()
            return
        if e.key() == Qt.Key_Escape:
            if self._on_clear:
                self._on_clear()
            return
        if e.key() in (Qt.Key_Up, Qt.Key_Down):
            if self._on_navigate:
                self._on_navigate(-1 if e.key() == Qt.Key_Up else 1)
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            self._toggle_note()
            return
        super().keyPressEvent(e)

    def eventFilter(self, watched, event):
        if watched is self.name_edit and event.type() == QEvent.FocusOut:
            self._save()
        elif event.type() == QEvent.DragEnter and self._drag_has_files(event):
            self._set_drag_hover(True)
            event.acceptProposedAction()
            return True
        elif event.type() == QEvent.DragLeave:
            self._set_drag_hover(False)
        elif event.type() == QEvent.Drop and watched is self.name_edit and self._drag_has_files(event):
            self._set_drag_hover(False)
            self._add_attachment_paths(self._drop_file_paths(event))
            event.acceptProposedAction()
            return True
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
        menu.setStyleSheet(f"""
            QMenu {{
                background: {T.INPUT_BG};
                border: 1px solid {T.ACCENT};
                border-radius: 6px;
                padding: 4px 0;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                color: #333;
                min-width: 120px;
            }}
            QMenu::item {{ padding: 6px 16px 6px 12px; }}
            QMenu::item:selected {{ background: rgba(212,184,0,0.3); }}
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

    def _toggle_note(self):
        if self._expanded:
            self._collapse_note()
        else:
            self._expand_note()

    def _show_note_editor(self):
        """행 높이를 바꾸지 않고, 행 아래 예약 공간을 만들어 메모를 오버레이로 표시."""
        list_widget = self.parentWidget()
        if list_widget is None:
            return
        if self.note_editor.parent() is not list_widget:
            self.note_editor.setParent(list_widget)
        layout = list_widget.layout()
        idx = layout.indexOf(self)
        if idx < 0:
            return
        if self._note_placeholder is None:
            ph = QWidget()
            ph.setFixedHeight(_InlineNoteEdit.HEIGHT)
            ph.setStyleSheet('background: transparent;')
            layout.insertWidget(idx + 1, ph)
            self._note_placeholder = ph
        # 이동/리사이즈/스크롤 시에도 현재 placeholder 높이(자동 성장값)를 유지.
        # HEIGHT로 덮어쓰면 확장된 패널이 원래대로 되돌아감.
        ph_h = self._note_placeholder.height()
        self.note_editor.setGeometry(self.x() + 9, self.y() + self.height(),
                                     self.width() - 17, ph_h)
        self.note_editor.show()
        self.note_editor.raise_()

    def _update_note_editor_height(self):
        """내용 변화에 따라 placeholder/패널 높이 갱신. 아래 행들이 밀려남."""
        if not self._expanded or self._note_placeholder is None:
            return
        h = self.note_editor.desired_height()
        if h is None:
            return  # 레이아웃 미완료(첫 페인트 전) — 다음 갱신에서 처리
        if self._note_placeholder.height() == h:
            return
        self._note_placeholder.setFixedHeight(h)
        self.note_editor.setGeometry(self.x() + 9, self.y() + self.height(),
                                     self.width() - 17, h)
        self.note_editor.update()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self._expanded and self.note_editor.isVisible():
            self._show_note_editor()
            # 폭 변경 → 재래핑은 페인트 후에 완료되므로 디바운스 타이머로 재계산
            self._note_height_timer.start()

    def moveEvent(self, e):
        super().moveEvent(e)
        if self._expanded and self.note_editor.isVisible():
            self._show_note_editor()

    def _expand_note(self):
        self._expanded = True
        # ▼ 토글은 메모와 붙임 목록을 함께 펼침 (행 높이부터 늘린 후 오버레이 배치)
        if self._attachments and not self._attach_expanded:
            self._attach_expanded = True
            self._rebuild_attach_list()
            self._sync_row_height()
            if self._on_attach_toggle:
                self._on_attach_toggle(self.task['id'], True)
        self.note_editor.setPlainText(get_task_notes(self.task['id']))
        self.note_editor.ref_edit.setText(get_task_related_no(self.task['id']))
        self._show_note_editor()
        self._update_note_editor_height()  # 저장된 긴 메모 펼치자마자 반영
        self.btn_note.set_expanded(True)
        self.note_editor.focus_editor()
        if self._on_toggle:
            self._on_toggle(self.task['id'], True)

    def _collapse_note(self):
        self._expanded = False
        # ▼ 토글은 메모와 붙임 목록을 함께 접음 (클립 토글로 이미 접힌 경우는 무시)
        if self._attach_expanded:
            self._attach_expanded = False
            self._rebuild_attach_list()
            self._sync_row_height()
            if self._on_attach_toggle:
                self._on_attach_toggle(self.task['id'], False)
        self._flush_note()
        self._flush_ref()
        self.note_editor.hide()
        self.btn_note.set_expanded(False)
        if self._note_placeholder is not None:
            list_widget = self.parentWidget()
            if list_widget is not None:
                list_widget.layout().removeWidget(self._note_placeholder)
            self._note_placeholder.deleteLater()
            self._note_placeholder = None
        self.setFocus()
        if self._on_toggle:
            self._on_toggle(self.task['id'], False)

    def _flush_note(self):
        text = self.note_editor.toPlainText()
        set_task_notes(self.task['id'], text)
        self._update_note_color()

    def _flush_ref(self):
        ref = self.note_editor.ref_edit.text().strip()
        set_task_related_no(self.task['id'], ref)
        self._update_note_color()

    # ── 붙임파일 ──────────────────────────────────────────────
    def _load_attachments(self) -> list:
        raw = self.task.get('attachments')
        if not raw:
            raw = get_task_attachments(self.task['id'])
        try:
            lst = json.loads(raw) if raw else []
            return [p for p in lst if isinstance(p, str) and p.strip()] if isinstance(lst, list) else []
        except Exception:
            return []

    def _save_attachments(self):
        try:
            set_task_attachments(self.task['id'], json.dumps(self._attachments, ensure_ascii=False))
        except Exception:
            pass

    def _refresh_attach_btn(self):
        n = len(self._attachments)
        if n == 0:
            self.btn_attach.hide()
        else:
            if getattr(self, '_attach_icon_ok', False):
                # 아이콘 모드: 개수는 숫자 텍스트로, 버튼 폭을 숫자만큼 넓힘
                self.btn_attach.setText(str(n) if n > 1 else '')
                if n > 1:
                    fm = self.btn_attach.fontMetrics()
                    w = fm.horizontalAdvance(str(n)) + round(19 * self._scale)
                    self.btn_attach.setFixedSize(max(round(19 * self._scale), w),
                                                 round(19 * self._scale))
            else:
                self.btn_attach.setText('📎' if n == 1 else f'📎{n}')
                self.btn_attach.setFixedSize(round(19 * self._scale), round(19 * self._scale))
            accent = f'QPushButton {{ border: none; color: {T.ACCENT}; background: transparent; padding: 0; }}' \
                     f'QPushButton:pressed {{ color: {T.ACCENT}; }}' if self._attach_expanded else \
                     'QPushButton { border: none; color: #555555; background: transparent; padding: 0; }' \
                     f'QPushButton:pressed {{ color: {T.ACCENT}; }}'
            self.btn_attach.setStyleSheet(accent)
            self.btn_attach.show()

    def _attach_area_height(self) -> int:
        if not getattr(self, '_attach_expanded', False) or not self._attachments:
            return 0
        # 컨테이너 minimumSizeHint는 위젯 추가 직후 캐시 미갱신으로 최소값을
        # 반환함 → 자식 위젯 sizeHint를 직접 합산하는 방식이 정확 (잘림 방지)
        if self._attach_container is not None:
            lay = self._attach_container.layout()
            if lay is not None:
                m = lay.contentsMargins()
                total = m.top() + m.bottom()
                cnt = lay.count()
                for i in range(cnt):
                    w = lay.itemAt(i).widget()
                    if w is not None:
                        h = w.minimumSizeHint().height() or w.sizeHint().height()
                        total += h
                total += lay.spacing() * max(cnt - 1, 0)
                if total > 0:
                    return total
        n = len(self._attachments)
        # 폴백: 상하 패딩(4) + 헤더(16) + 행(26 × N) + 간격(2 × N-1) — 모두 스케일 반영
        s = self._scale
        return round(4 * s) + round(16 * s) + n * round(26 * s) + (n - 1) * round(2 * s)

    def _rebuild_attach_list(self):
        """붙임 목록 컨테이너를 현재 목록에 맞게 재구성."""
        if not self._attachments:
            self._attach_expanded = False
        if self._attach_container is None:
            self._attach_container = QWidget()
            self._attach_layout = QVBoxLayout(self._attach_container)
            self._attach_layout.setContentsMargins(6, 2, 6, 2)
            self._attach_layout.setSpacing(2)
            self._outer.addWidget(self._attach_container)
        while self._attach_layout.count():
            item = self._attach_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._attach_container.setVisible(self._attach_expanded)
        self._attach_container.setStyleSheet('background: transparent;')
        if not self._attach_expanded:
            self._refresh_attach_btn()
            return
        header = QLabel(f'붙임파일 {len(self._attachments)}개')
        header.setFont(pr_font(round(10 * self._scale)))
        header.setStyleSheet('color: #666; background: transparent;')
        header.setFixedHeight(round(16 * self._scale))
        self._attach_layout.addWidget(header)
        for p in self._attachments:
            self._attach_layout.addWidget(self._make_attach_row(p))
        self._refresh_attach_btn()

    def _make_attach_row(self, path):
        s = self._scale
        row = QWidget()
        row.setFixedHeight(round(26 * s))
        row.setStyleSheet('QWidget { background: rgba(255,255,255,0.4); border-radius: 4px; }')
        h = QHBoxLayout(row)
        h.setContentsMargins(4, 1, 2, 1)
        h.setSpacing(2)
        fname = os.path.basename(path)
        btn = _ElideButton(fname)
        btn.setFixedHeight(round(24 * s))
        btn.setFont(pr_font(round(11 * s)))
        btn.setToolTip(path)
        btn.setStyleSheet(
            'QPushButton { border: none; background: transparent; }'
        )
        btn.clicked.connect(lambda: self._open_attachment(path))
        h.addWidget(btn, 1)
        x = QPushButton('✕')
        x.setFixedSize(round(18 * s), round(18 * s))
        x.setCursor(Qt.PointingHandCursor)
        x.setStyleSheet(
            'QPushButton { border: none; background: transparent; color: #999; font-size: 11px; }'
            f'QPushButton:hover {{ color: {T.DANGER}; }}'
        )
        x.clicked.connect(lambda: self._remove_attachment(path))
        h.addWidget(x)
        return row

    def _toggle_attach(self):
        if not self._attachments:
            return
        self._attach_expanded = not self._attach_expanded
        self._rebuild_attach_list()
        self._sync_row_height()
        self.setFocus()
        if self._on_attach_toggle:
            self._on_attach_toggle(self.task['id'], self._attach_expanded)

    def _add_attachment_paths(self, paths):
        added = False
        for p in paths:
            norm = os.path.normpath(p)
            if norm not in self._attachments:
                self._attachments.append(norm)
                added = True
        if added:
            self._save_attachments()
            self._attach_expanded = True
            self._refresh_attach_btn()
            self._rebuild_attach_list()
            self._sync_row_height()
            if self._on_attach_toggle:
                self._on_attach_toggle(self.task['id'], True)

    def _remove_attachment(self, path):
        if path in self._attachments:
            self._attachments.remove(path)
        self._save_attachments()
        self._refresh_attach_btn()
        self._rebuild_attach_list()
        self._sync_row_height()
        if not self._attachments and self._on_attach_toggle:
            self._on_attach_toggle(self.task['id'], False)

    def _open_attachment(self, path):
        if not os.path.exists(path):
            QMessageBox.information(
                self, '붙임파일',
                f'파일을 찾을 수 없습니다.\n\n{path}\n\n파일이 이동되거나 삭제되었을 수 있습니다.'
            )
            return
        try:
            if sys.platform == 'win32' and hasattr(os, 'startfile'):
                os.startfile(path)
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception:
            QMessageBox.warning(self, '붙임파일', f'파일을 열 수 없습니다.\n{path}')

    # ── 드래그앤드롭 ──────────────────────────────────────────
    def _drag_has_files(self, event) -> bool:
        mime = event.mimeData()
        if mime is None or not mime.hasUrls():
            return False
        return any(u.isLocalFile() for u in mime.urls())

    def _drop_file_paths(self, event) -> list:
        paths = []
        for u in event.mimeData().urls():
            if u.isLocalFile():
                p = u.toLocalFile()
                if p and os.path.isfile(p):
                    paths.append(os.path.normpath(p))
        return paths

    def _set_drag_hover(self, on: bool):
        if self._drag_hover != on:
            self._drag_hover = on
            self._apply_row_style()

    def dragEnterEvent(self, e):
        if self._drag_has_files(e):
            self._set_drag_hover(True)
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if self._drag_has_files(e):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragLeaveEvent(self, e):
        self._set_drag_hover(False)
        super().dragLeaveEvent(e)

    def dropEvent(self, e):
        self._set_drag_hover(False)
        if self._drag_has_files(e):
            self._add_attachment_paths(self._drop_file_paths(e))
            e.acceptProposedAction()
        else:
            e.ignore()

    def _update_note_color(self):
        has = bool(self.note_editor.toPlainText().strip()
                   or self.note_editor.ref_edit.text().strip())
        self.btn_note.set_note_color(has, self._overdue)

    def _show_date_picker(self):
        if not hasattr(self, '_date_picker_popup'):
            self._date_picker_popup = QWidget()
            self._date_picker_popup.setWindowFlags(Qt.Popup)
            self._date_picker_popup.setStyleSheet(f'background: white; border: 1px solid {T.ACCENT}; border-radius: 6px;')
            
            layout = QVBoxLayout(self._date_picker_popup)
            layout.setContentsMargins(8, 8, 8, 8)
            layout.setSpacing(6)
            
            self._cal_popup = CustomCalendarWidget()
            self._cal_popup.setStyleSheet(f"""
                QCalendarWidget {{ background-color: white; }}
                QCalendarWidget QAbstractItemView {{
                    background-color: white; color: black;
                    font-family: 'Malgun Gothic'; font-size: 10pt;
                    selection-background-color: {T.ACCENT}; selection-color: white;
                }}
                QCalendarWidget QWidget {{ background-color: white; }}
                QCalendarWidget QToolButton {{
                    background-color: white; color: black;
                    font-family: 'Malgun Gothic'; font-size: 10pt;
                    padding: 2px 4px; min-width: 40px;
                }}
                QCalendarWidget QWidget#qt_calendar_navigationbar {{ background-color: white; }}
                QCalendarWidget QSpinBox {{ background-color: white; color: black; }}
            """)
            fmt_sat = QTextCharFormat()
            fmt_sat.setForeground(QColor(T.SATURDAY))
            self._cal_popup.setWeekdayTextFormat(Qt.Saturday, fmt_sat)
            fmt_sun = QTextCharFormat()
            fmt_sun.setForeground(QColor(T.SUNDAY))
            self._cal_popup.setWeekdayTextFormat(Qt.Sunday, fmt_sun)
            self._cal_popup.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
            layout.addWidget(self._cal_popup)
            
            time_row = QHBoxLayout()
            time_row.setSpacing(6)
            
            self._time_checkbox = QCheckBox('시간 설정')
            self._time_checkbox.setFont(pr_font(10))
            self._time_checkbox.setStyleSheet('color: #333; background: transparent;')
            self._time_checkbox.stateChanged.connect(self._on_time_checkbox_changed)
            time_row.addWidget(self._time_checkbox)
            
            self._picker_time_edit = QTimeEdit(QTime(9, 0))
            self._picker_time_edit.setDisplayFormat('HH:mm')
            self._picker_time_edit.setFixedWidth(68)
            self._picker_time_edit.setFont(pr_font(11))
            self._picker_time_edit.setEnabled(False)
            self._picker_time_edit.setStyleSheet(f"""
                QTimeEdit {{
                    background: rgba(255,255,255,0.5);
                    border: 1px solid {T.ACCENT};
                    border-radius: 4px;
                    padding: 2px 6px;
                    color: #333;
                    qproperty-alignment: AlignCenter;
                }}
                QTimeEdit:disabled {{
                    background: rgba(255,255,255,0.2);
                    border: 1px solid #ddd;
                    color: #bbb;
                }}
                QTimeEdit::up-button {{ width: 0px; }}
                QTimeEdit::down-button {{ width: 0px; }}
            """)
            time_row.addWidget(self._picker_time_edit)
            time_row.addStretch()
            
            self._confirm_btn = QPushButton('확인')
            self._confirm_btn.setFont(pr_font(10))
            self._confirm_btn.setFixedSize(60, 28)
            self._confirm_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {T.ACCENT};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    font-weight: bold;
                }}
                QPushButton:hover {{ background: #c0a800; }}
                QPushButton:pressed {{ background: #b09800; }}
            """)
            self._confirm_btn.clicked.connect(self._on_confirm_date_time)
            time_row.addWidget(self._confirm_btn)
            
            layout.addLayout(time_row)

        anchor = self.sender() if isinstance(self.sender(), QWidget) else getattr(self, '_add_date_btn', None)
        if anchor is None:
            return

        if self.task.get('deadline'):
            try:
                d = date.fromisoformat(self.task['deadline'][:10])
                self._cal_popup.setSelectedDate(QDate(d.year, d.month, d.day))
                self._cal_popup.setCurrentPage(d.year, d.month)
                
                if len(self.task['deadline']) > 10:
                    time_str = self.task['deadline'][11:]
                    h, m = map(int, time_str.split(':'))
                    self._time_checkbox.setChecked(True)
                    self._picker_time_edit.setTime(QTime(h, m))
                else:
                    self._time_checkbox.setChecked(False)
                    self._picker_time_edit.setEnabled(False)
                    self._picker_time_edit.setTime(QTime(9, 0))
            except ValueError:
                today = QDate.currentDate()
                self._cal_popup.setCurrentPage(today.year(), today.month())
                self._cal_popup.setSelectedDate(today)
                self._time_checkbox.setChecked(False)
                self._picker_time_edit.setEnabled(False)
                self._picker_time_edit.setTime(QTime(9, 0))
        else:
            today = QDate.currentDate()
            self._cal_popup.setCurrentPage(today.year(), today.month())
            self._cal_popup.setSelectedDate(today)
            self._time_checkbox.setChecked(False)
            self._picker_time_edit.setEnabled(False)
            self._picker_time_edit.setTime(QTime(9, 0))

        popup_size = self._date_picker_popup.sizeHint()
        pos = anchor.mapToGlobal(QPoint(0, anchor.height()))

        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        screen_rect = screen.availableGeometry()
        if pos.x() + popup_size.width() > screen_rect.right():
            pos.setX(screen_rect.right() - popup_size.width())
        if pos.y() + popup_size.height() > screen_rect.bottom():
            pos = anchor.mapToGlobal(QPoint(0, -popup_size.height()))

        self._date_picker_popup.move(pos)
        self._date_picker_popup.show()

    def _on_time_checkbox_changed(self, state):
        self._picker_time_edit.setEnabled(state == Qt.Checked)

    def _on_confirm_date_time(self):
        self._date_picker_popup.hide()
        qdate = self._cal_popup.selectedDate()
        deadline_str = qdate.toString('yyyy-MM-dd')
        
        if self._time_checkbox.isChecked():
            time_str = self._picker_time_edit.time().toString('HH:mm')
            deadline_str += f' {time_str}'
        
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
