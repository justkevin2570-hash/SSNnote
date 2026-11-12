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
    QCheckBox, QFrame
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QPoint, QPointF, QRectF, QSize, pyqtSignal, QMimeData
from PyQt5.QtGui import QFont, QFontMetrics, QColor, QPainter, QTextCharFormat, QPalette, QTextOption, QTextLayout, QIcon, QTextCursor
from db import (update_task,
                set_task_priority, set_task_recurrence, get_task_notes, set_task_notes,
                get_task_related_no, set_task_related_no,
                get_task_attachments, set_task_attachments)
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, _normalize_doc_number
from .common import _base_path, _svg_icon, MAT, mat_font, MI, mi_font, pr_font, note_font, note_scale, set_note_bold, themed_pixmap, px
from .dates import calc_dday
from .widgets import AddDateButton, CustomCalendarWidget, DdayLabel, RotatedMenuButton
from .dialogs import CustomRecurrenceDialog
from .theme import T, tc, tct, qcolor, current_accent, rgba

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
            line.setLeadingIncluded(True)        # 줄 사이 여백이 있는 글꼴(리틱체 등)은 실제 줄 간격만큼
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
    """업무 메모 펼치기/접기 토글 버튼.

    메모가 있으면 메모장 아이콘, 펼쳐져 있으면 위쪽 꺾쇠. 메모가 없으면 평소엔 보이지 않고
    업무 줄에 마우스를 올렸을 때만 옅은 연필 아이콘(메모 쓰기)으로 보인다.
    """

    def __init__(self, scale=1.0, glyph_offset=4):
        super().__init__()
        self._accent       = current_accent()
        self._hovered      = False
        self._has_note     = False
        self._overdue      = False
        self._expanded     = False
        self._row_hover    = True     # 업무 줄에서는 TaskRow가 켜고 끈다 (그 밖의 곳에서는 늘 보임)
        self._glyph_offset = glyph_offset
        self.setFlat(True)
        sz = int(24 * scale)
        self._icon_px = max(14, int(18 * scale))
        self.setFixedSize(sz, sz)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet('QPushButton { border: none; background: transparent; padding: 0; }')

    def set_note_color(self, has_note: bool, overdue: bool = None):
        self._has_note = has_note
        if overdue is not None:
            self._overdue = overdue
        self.update()

    def set_expanded(self, expanded: bool):
        self._expanded = expanded
        self.update()

    def set_row_hover(self, on: bool):
        self._row_hover = on
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
        if self._expanded:
            name, alpha = 'ph.caret-up-bold', 200
        elif self._has_note:
            name, alpha = 'ph.notepad-bold', (140 if self._overdue else 215)
        elif self._row_hover or self._hovered:
            name, alpha = 'ph.note-pencil', 110
        else:
            return
        if self._hovered:
            color = QColor(self._accent)
        else:
            color = qcolor('#000000', alpha)
        try:
            import qtawesome as qta
            pm = qta.icon(name, color=color).pixmap(self._icon_px, self._icon_px)
        except Exception:
            return
        p = QPainter(self)
        x = (self.width() - self._icon_px) // 2
        y = (self.height() - self._icon_px) // 2
        p.drawPixmap(x, y, pm)
        p.end()


class _ElideButton(QPushButton):
    """폭을 넘는 텍스트를 오른쪽 '…'로 잘라내는 버튼 (붙임 파일명 표시용)."""

    def __init__(self, text=''):
        super().__init__(text)
        self._accent = current_accent()
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
        p.setPen(QColor(self._accent) if self._hovered else QColor(tc('#333333')))
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
        _ac = current_accent()
        _ach = QColor(_ac).darker(125).name()
        # 높이는 TaskRow가 내용 기반으로 조절(setGeometry). min/max만 지정.
        self.setMinimumHeight(self.HEIGHT)
        self.setMaximumHeight(self.MAX_HEIGHT)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        top = QHBoxLayout()
        top.setSpacing(6)
        self.ref_edit = _RefLineEdit('공문 관련번호 적어두기', note_font(int(11 * scale)))
        _ref_font = note_font(int(12 * scale))
        _ref_font.setStretch(100)
        _ref_font.setLetterSpacing(QFont.PercentageSpacing, 100.0)
        self.ref_edit.setFont(_ref_font)
        self.ref_edit.setStyleSheet(f"""
            QLineEdit {{
                background: {tc("rgba(255,255,255,0.6)")};
                border: 1px solid {_ac};
                border-radius: 4px;
                padding: 2px 6px;
                color: {tct("#333")};
            }}
            QLineEdit:focus {{ border: 1px solid {_ach}; }}
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
                background: {tc("rgba(255,255,255,0.6)")};
                border: 1px solid {_ac};
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: {rgba(_ac, 0.25)}; }}
            QPushButton:pressed {{ background: {rgba(_ac, 0.42)}; padding: 2px 0 0 2px; }}
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
                background: {tc("rgba(255,255,255,0.6)")};
                border: 1px solid {_ac};
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
            }}
            QPushButton:hover {{ background: {rgba(_ac, 0.25)}; }}
            QPushButton:pressed {{ background: {rgba(_ac, 0.42)}; padding: 2px 0 0 2px; }}
        """)
        self.btn_shot.clicked.connect(self._start_inline_capture)
        top.addWidget(self.btn_shot)
        outer.addLayout(top)

        self.editor = QPlainTextEdit()
        self.editor.setPlaceholderText('esc 또는 ctrl+enter 을 누르면 접어집니다.')
        self.editor.setFont(note_font(int(11 * scale)))
        self.editor.document().setDocumentMargin(4)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.editor.setStyleSheet(f"""
            QPlainTextEdit {{
                background: {tc("rgba(255,255,255,0.6)")};
                border: 1px solid {_ac};
                border-radius: 4px;
                color: {tct("#333")};
                padding: 2px 4px;
            }}
            QPlainTextEdit:focus {{ border: 1px solid {_ach}; }}
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
        self._accent = current_accent()     # 이 메모 창 색에 맞춘 포인트 색
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
            self.btn_star.setIcon(_svg_icon(self._star_svg, tc('#D4A800'), self._star_sz))
            self.btn_star.setIconSize(QSize(self._star_sz, self._star_sz))
        else:
            self.btn_star.setText(MI.STAR_BORDER)
            self.btn_star.setFont(mi_font(_star_pt))
            self.btn_star.setStyleSheet(f"""
                QPushButton {{ color: {tc("rgba(0,0,0,0.25)")}; border: none; padding: 0; }}
                QPushButton:hover {{ color: {tc("rgba(0,0,0,0.5)")}; }}
                QPushButton:pressed {{ padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }}
            """)
        self.btn_star.clicked.connect(self._toggle_priority)
        # 업무 줄의 아이콘·표시는 모두 '업무 이름 첫 줄'의 세로 가운데에 맞춘다
        # (이름이 두 줄로 넘어가도 첫 줄 옆에 나란히). 첫 줄 높이 = 이름 칸 한 줄 높이.
        _nf = note_font(12)
        _nf.setPointSizeF(12 * scale * note_scale())
        self._line1_h = max(28, QFontMetrics(_nf).height() + 4)
        layout.addWidget(self._line_box(self.btn_star), 0, Qt.AlignTop)

        deadline = task['deadline']
        overdue = False
        tone = 'later'
        if deadline:
            dday = calc_dday(deadline)
            if dday.startswith('D+'):
                overdue = True
                tone = 'overdue'
            elif dday == 'D-day' or '시간' in dday:
                tone = 'today'
            elif dday == '날짜오류':
                tone = 'error'
            else:
                days_left = int(dday[2:])
                if days_left <= 3:
                    tone = 'soon'
                elif days_left <= 7:
                    tone = 'week'
            suffix = f'{dday} 지남' if overdue else dday
        else:
            suffix = ''

        if overdue:
            _eff = QGraphicsOpacityEffect()
            _eff.setOpacity(0.0)
            self.btn_star.setGraphicsEffect(_eff)
            self.btn_star.setEnabled(False)

        self._strikethrough = bool(task.get('strikethrough', 0))
        name_color = tc('#666') if overdue else tc('#111')

        self.name_edit = _AutoHeightEdit(task['name'])
        self.name_edit.setFocusPolicy(Qt.NoFocus)
        # 파일 드롭이 name_edit(텍스트 위젯)에 가로채지고 부모 TaskRow로 버블업되도록 끔
        self.name_edit.setAcceptDrops(False)
        _base_pt = 12 * scale * note_scale()      # 글꼴별 크기 보정 (리틱체는 작게)
        name_font = note_font(12)
        name_font.setPointSizeF(_base_pt)  # 기한 지나도(overdue) 크기 유지 — 축소하지 않음
        if _is_starred and not overdue:     # 굵은 글씨는 중요 표시한 업무만 (급한 정도는 D-day 색으로)
            set_note_bold(name_font)
        self.name_edit.setFont(name_font)
        self.name_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background: transparent;
                border: none;
                color: {name_color};
                padding: 0px;
            }}
            QPlainTextEdit:focus {{
                background: {tc("rgba(255,255,255,0.6)")};
                border-bottom: 1px solid {self._accent};
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
            dday_lbl = DdayLabel(suffix, date_text, tone=tone)
            dday_lbl.clicked.connect(self._show_date_picker)
            dday_font = note_font(10)
            dday_font.setPointSizeF(10.5 * scale * note_scale())
            dday_font.setWeight(QFont.DemiBold if tone in ('overdue', 'today', 'soon') else QFont.Normal)
            dday_lbl.setFont(dday_font)
            dday_lbl.setToolTip('마감일 바꾸기')
            self._dday_text0 = suffix
            self._tone0 = tone
        else:
            dday_lbl = None
            self._add_date_btn = AddDateButton()
            self._add_date_btn.clicked.connect(self._show_date_picker)

        btn_menu = RotatedMenuButton()
        btn_menu.set_row_active(False)      # 줄에 마우스를 올리거나 선택했을 때만 보임
        self.btn_menu = btn_menu
        _f = btn_menu.font()
        _f.setPixelSize(19)
        btn_menu.setFont(_f)
        btn_menu.setStyleSheet('background: transparent; border: none; padding: 4px 6px; border-radius: 4px;')
        btn_menu.clicked.connect(lambda: self._open_task_menu(btn_menu, task, on_delete))
        btn_menu.hovered.connect(self._set_row_highlight)

        self._dday_lbl = dday_lbl
        _has_note = bool(task.get('notes', '') or task.get('related_no', ''))
        self._overdue = overdue
        self.btn_note = _NoteButton(scale=scale)
        self.btn_note.set_note_color(_has_note, overdue)
        self.btn_note.setToolTip('업무 메모 (펼치기/접기)')
        self.btn_note.clicked.connect(self._toggle_note)
        self.btn_note.set_expanded(expanded)
        self.btn_note.set_row_hover(False)

        # ── 붙임파일 버튼 (클립 아이콘) ─────────────────────────
        self.btn_attach = QPushButton()
        self.btn_attach.setCursor(Qt.PointingHandCursor)
        self.btn_attach.setToolTip('붙임파일 목록 (펼치기/접기)')
        self.btn_attach.setFlat(True)
        self.btn_attach.setFixedSize(round(19 * scale), round(19 * scale))
        self.btn_attach.setFont(pr_font(round(11 * scale)))
        self.btn_attach.setStyleSheet(
            f'QPushButton {{ border: none; color: {tct("#555555")}; background: transparent; padding: 0; }}'
            f'QPushButton:pressed {{ color: {self._accent}; }}'
        )
        _att_icon = os.path.join(_base_path(), 'assets', 'clip.png')
        self._attach_icon_ok = os.path.exists(_att_icon)
        if self._attach_icon_ok:
            self.btn_attach.setIcon(QIcon(themed_pixmap(_att_icon)))
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
        name_note_layout.addWidget(self.name_edit, 0, Qt.AlignTop)
        name_note_layout.addWidget(self._line_box(self.btn_note), 0, Qt.AlignTop)
        name_note_layout.addWidget(self._line_box(self.btn_attach), 0, Qt.AlignTop)
        layout.addLayout(name_note_layout)
        layout.addStretch(1)
        if task.get('recurrence', ''):
            lbl_recur = QLabel(MAT.REPEAT)
            _recur_font = mat_font(int(13 * scale))
            lbl_recur.setFont(_recur_font)
            lbl_recur.setStyleSheet(f'color: {tc("rgba(0,0,0,0.45)")}; background: transparent;')
            lbl_recur.setToolTip('반복 업무')
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
            layout.addWidget(self._line_box(icon_dday_widget), 0, Qt.AlignTop)
        else:
            if dday_lbl:
                layout.addWidget(self._line_box(dday_lbl), 0, Qt.AlignTop)
            elif hasattr(self, '_add_date_btn'):
                layout.addWidget(self._line_box(self._add_date_btn), 0, Qt.AlignTop)
        btn_menu.setFixedHeight(min(btn_menu.sizeHint().height(), self._line1_h))
        layout.addWidget(self._line_box(btn_menu), 0, Qt.AlignTop)

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
        self._urg = {}
        self.apply_urgency()

    # ── 급함 표시 (ui/motion.py) ─────────────────────────────────
    def apply_urgency(self, now=None):
        """마감까지 남은 정도에 맞춰 D-day 숨쉬기·남은 시간·박동·별 반짝임·세로선을 정한다.
        메모 창이 30초마다 다시 부른다 (줄을 새로 만들지 않고 그 자리에서 바꿈)."""
        from . import motion
        cfg = motion.load()
        u = motion.urgency(self.task, now, cfg)
        self._urg = u
        lbl = self._dday_lbl
        if lbl is not None:
            if u['countdown'] and u['remaining'] is not None:
                lbl.set_tone('today')
                lbl.set_texts(motion.remaining_text(u['remaining']))
                lbl.setToolTip('마감까지 남은 시간 — 눌러서 마감일 바꾸기')
            elif u['overdue'] and not self._overdue:          # 오늘 마감 시각이 지남
                lbl.set_tone('overdue')
                lbl.set_texts('마감 지남')
            else:
                lbl.set_tone(getattr(self, '_tone0', 'later'))
                lbl.set_texts(getattr(self, '_dday_text0', lbl.text()))
            lbl.set_motion(0.0, False)          # D-day 표시는 움직이지 않음 (움직임은 왼쪽 세로선 하나로)
        self._apply_star_motion(u, cfg)
        self._apply_line_motion(u, cfg)
        self.update()

    def _line_style(self, u, cfg):
        """왼쪽 세로선: (색, 숨쉬기 주기) 또는 None. ★ 금색 · 그 밖 주황 (기한 지남은 선 없음 — 색 표시로)."""
        if not cfg['enabled'] or u.get('overdue') or not u.get('breathe'):
            return None
        return ('#C98A00' if self.task.get('priority') else '#EA580C'), u['breathe']

    def _apply_line_motion(self, u, cfg):
        from . import motion
        self._line = self._line_style(u, cfg)
        if self._line:
            motion.clock().register(self, lambda t: self.update(0, 0, 8, self.height()))   # 세로선 자리만 다시 그림
        else:
            motion.clock().unregister(self)

    def _apply_star_motion(self, u, cfg):
        from . import motion
        starred = bool(self.task.get('priority'))
        if starred and self._overdue:
            # 중요한데 놓친 업무: 빨간 테두리 별 (누를 수는 없음)
            import qtawesome as qta
            self.btn_star.setGraphicsEffect(None)
            self.btn_star.setText('')
            red = qta.icon('ph.star-bold', color=tc('#E5484D'))
            ic = QIcon()
            sz = int(self._star_sz * 0.8)
            for mode in (QIcon.Normal, QIcon.Disabled):      # 누를 수 없어도(기한 지남) 회색이 아닌 빨강
                ic.addPixmap(red.pixmap(sz, sz), mode)
            self.btn_star.setIcon(ic)
            self.btn_star.setIconSize(QSize(int(self._star_sz * 0.8), int(self._star_sz * 0.8)))
            self.btn_star.setToolTip('중요한 업무인데 기한이 지났어요')
            return
        if starred and u.get('star_glow'):
            eff = QGraphicsOpacityEffect(self.btn_star)
            self.btn_star.setGraphicsEffect(eff)
            period = u['breathe'] or 3.0

            def glow(t, eff=eff, period=period):      # 세로선과 같은 박자로 0%까지 옅어졌다 돌아옴
                try:
                    eff.setOpacity(motion.breath(t, period))
                except RuntimeError:
                    pass
            motion.clock().register(self.btn_star, glow)
        else:
            motion.clock().unregister(self.btn_star)
            if not self._overdue:
                self.btn_star.setGraphicsEffect(None)

    def paintEvent(self, e):
        from PyQt5.QtWidgets import QStyleOption, QStyle
        opt = QStyleOption()
        opt.initFrom(self)
        p = QPainter(self)
        self.style().drawPrimitive(QStyle.PE_Widget, opt, p, self)
        u = getattr(self, '_urg', {}) or {}
        line = getattr(self, '_line', None)
        if line or u.get('countdown'):
            p.setRenderHint(QPainter.Antialiasing)
            # 업무 이름 줄의 실제 높이 (붙임파일 목록이 펼쳐져 있으면 그 위까지)
            line_h = min(max(self.name_edit._fixed_height, 28), self.height() - self._attach_area_height())
            if line:
                # 급한 업무 왼쪽 세로선: 숨쉬듯 진해졌다 옅어짐 (급할수록 빠르게)
                from . import motion
                color, period = line
                c = QColor(tc(color))
                c.setAlphaF(motion.breath(motion.clock().now(), period))     # 0%까지 옅어졌다 돌아옴
                p.setPen(Qt.NoPen)
                p.setBrush(c)
                p.drawRoundedRect(QRectF(1, 4, 4, line_h - 8), 2, 2)
            if u.get('countdown'):
                # 마감까지 남은 시간 막대: 줄 아래 얇게, 차오를수록 진하게
                y = line_h - 2.5
                left, right = 30.0, self.width() - 10.0
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(0, 0, 0, 18))
                p.drawRoundedRect(QRectF(left, y, right - left, 2.5), 1.25, 1.25)
                bar = QColor(tc('#F76B15'))
                bar.setAlphaF(0.55 + 0.45 * u['progress'])
                p.setBrush(bar)
                p.drawRoundedRect(QRectF(left, y, max(3.0, (right - left) * u['progress']), 2.5), 1.25, 1.25)
        p.end()

    def _line_box(self, w):
        """w를 업무 이름 첫 줄 높이의 칸에 넣어 세로 가운데에 놓는다."""
        box = QWidget()
        box.setStyleSheet('background: transparent;')
        box.setFixedHeight(self._line1_h)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(w, 0, Qt.AlignVCenter | Qt.AlignHCenter)
        w._line_box_parent = box      # 칸은 늘 보이고, 안의 위젯이 숨으면 폭이 0이 됨
        return box

    def _set_row_highlight(self, on: bool):
        self._hovered = on
        self._apply_row_style()

    def _sync_row_controls(self):
        """⋮·메모 쓰기·＋날짜는 줄에 마우스를 올렸거나 선택했을 때만 보인다."""
        on = self._hovered or self._selected
        self.btn_menu.set_row_active(on)
        self.btn_note.set_row_hover(on)
        if hasattr(self, '_add_date_btn'):
            self._add_date_btn.set_row_hover(on)

    def enterEvent(self, e):
        self._hovered = True
        self._apply_row_style()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hovered = False
        self._apply_row_style()
        super().leaveEvent(e)

    def _apply_row_style(self):
        if getattr(self, '_drag_hover', False):
            style = f'QWidget#TaskRow {{ background: {rgba(self._accent, 0.16)}; border-radius: 6px; }}'
        elif self._selected:
            style = f'QWidget#TaskRow {{ background: {rgba(self._accent, 0.22)}; border-radius: 6px; }}'
        elif self._hovered:
            style = f'QWidget#TaskRow {{ background: {tc("rgba(0,0,0,12)")}; border-radius: 6px; }}'
        else:
            style = 'QWidget#TaskRow { background: transparent; }'
        self.setStyleSheet(style)
        if hasattr(self, 'btn_note'):
            self._sync_row_controls()

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
        """업무 ⋮ 메뉴: 업무 이름·마감 → 반복(한 줄 칩) → 완료 표시 · 삭제."""
        from PyQt5.QtWidgets import QWidgetAction
        from . import menu_style as ms
        menu = ms.make_menu(self)
        menu.setMinimumWidth(px(250) + 2 * ms.SHADOW)

        # ── 머리: 업무 이름 + 마감 ──
        head = QWidget()
        head.setStyleSheet('background: transparent;')
        hl = QVBoxLayout(head)
        hl.setContentsMargins(12, 6, 12, 6)
        hl.setSpacing(1)
        name = QLabel()
        nf = QFont(menu.font())
        nf.setWeight(QFont.DemiBold)
        name.setFont(nf)
        name.setStyleSheet(f'color: {ms.TEXT.name()};')
        name.setText(QFontMetrics(nf).elidedText(task.get('name', ''), Qt.ElideRight, px(230)))
        name.setToolTip(task.get('name', ''))
        hl.addWidget(name)
        sub = QLabel(self._menu_subtitle())
        sf = QFont(menu.font())
        sf.setPointSizeF(max(8.0, sf.pointSizeF() - 1.5))
        sub.setFont(sf)
        sub.setStyleSheet(f'color: {ms.MUTED.name()};')
        hl.addWidget(sub)
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(head)
        menu.addAction(wa)
        menu.addSeparator()

        # ── 반복: 한 줄 칩 ──
        cur = self.task.get('recurrence', '')
        rep = QWidget()
        rep.setStyleSheet('background: transparent;')
        rl = QVBoxLayout(rep)
        rl.setContentsMargins(12, 4, 12, 8)
        rl.setSpacing(6)
        cap = QLabel('반복')
        cap.setFont(sf)
        cap.setStyleSheet(f'color: {ms.MUTED.name()};')
        rl.addWidget(cap)
        chips = QHBoxLayout()
        chips.setSpacing(4)
        sel_bg = ms.TEXT.name()
        sel_fg = ms.BG.name()
        idle = f"""
            QPushButton {{ background: transparent; color: {ms.TEXT.name()}; border: 1px solid {ms.SEP.name(QColor.HexArgb)};
                          border-radius: 7px; padding: 3px 8px; }}
            QPushButton:hover {{ background: {ms.HOVER.name(QColor.HexArgb)}; }}"""
        chosen = f"""
            QPushButton {{ background: {sel_bg}; color: {sel_fg}; border: 1px solid {sel_bg};
                          border-radius: 7px; padding: 3px 8px; font-weight: 600; }}"""
        options = [('없음', ''), ('매주', 'weekly'), ('격주', 'biweekly'), ('매월', 'monthly'), ('매년', 'yearly')]
        for label, val in options:
            b = QPushButton(label)
            b.setFont(sf)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(chosen if cur == val else idle)
            b.clicked.connect(lambda _, v=val: (menu.close(), self._set_recurrence(v)))
            chips.addWidget(b)
        custom = cur.startswith('custom:')
        b = QPushButton('직접…')
        b.setFont(sf)
        b.setCursor(Qt.PointingHandCursor)
        b.setToolTip('예: 매월 둘째 주 화요일')
        b.setStyleSheet(chosen if custom else idle)
        b.clicked.connect(lambda: (menu.close(), self._open_custom_recurrence()))
        chips.addWidget(b)
        chips.addStretch()
        rl.addLayout(chips)
        wr = QWidgetAction(menu)
        wr.setDefaultWidget(rep)
        menu.addAction(wr)
        menu.addSeparator()

        # ── 완료 표시 · 삭제 ──
        act_strike = ms.add_action(menu, '완료 표시 해제' if self._strikethrough else '완료 표시',
                                   None, icon='ph.arrow-counter-clockwise' if self._strikethrough else 'ph.check-circle',
                                   shortcut='C')
        act_delete = ms.add_action(menu, '삭제', None, icon='ph.trash', shortcut='Del')
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

    def _menu_subtitle(self) -> str:
        """⋮ 메뉴 머리의 작은 글: '11월 10일 (화) 17:00 · D-3 · 매주'."""
        parts = []
        dl = self.task.get('deadline') or ''
        if dl:
            try:
                d = date.fromisoformat(dl[:10])
                parts.append(f"{d.month}월 {d.day}일 ({'월화수목금토일'[d.weekday()]})" + (f' {dl[11:16]}' if len(dl) > 10 else ''))
                parts.append(calc_dday(dl))
            except ValueError:
                pass
        else:
            parts.append('마감일 없음')
        rec = self.task.get('recurrence', '')
        names = {'weekly': '매주', 'biweekly': '격주', 'monthly': '매월', 'yearly': '매년'}
        if rec:
            parts.append(names.get(rec, '직접 정한 반복'))
        if self._strikethrough:
            parts.append('완료')
        return ' · '.join(parts)

    def _toggle_strikethrough(self):
        """완료 표시: 가운데 선이 왼쪽에서 오른쪽으로 그어진 뒤 적용 (움직임 끔이면 바로)."""
        from . import motion
        if not self._strikethrough and motion.load()['enabled'] and self.isVisible():
            self._animate_strike(self._apply_strikethrough)
        else:
            self._apply_strikethrough()

    def _animate_strike(self, then):
        from PyQt5.QtCore import QVariantAnimation, QEasingCurve
        ed = self.name_edit
        overlay = QWidget(ed)
        overlay.setAttribute(Qt.WA_TransparentForMouseEvents)
        overlay.setGeometry(ed.rect())
        fm = QFontMetrics(ed.font())
        text_w = min(ed.width() - 4, fm.horizontalAdvance(ed.toPlainText().split('\n')[0]) + 2)
        y = 2 + fm.ascent() - fm.xHeight() / 2.0 + 1
        state = {'v': 0.0}

        def paint(_e, ov=overlay):
            pp = QPainter(ov)
            pp.setRenderHint(QPainter.Antialiasing)
            pen = pp.pen()
            pen.setColor(QColor(tct('#333333')))
            pen.setWidthF(1.6)
            pp.setPen(pen)
            pp.drawLine(QPointF(1, y), QPointF(1 + text_w * state['v'], y))
            pp.end()
        overlay.paintEvent = paint
        overlay.show()
        anim = QVariantAnimation(self)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setDuration(280)
        anim.setEasingCurve(QEasingCurve.OutCubic)

        def step(v):
            state['v'] = v
            overlay.update()

        def done():
            overlay.hide()
            overlay.deleteLater()
            then()
        anim.valueChanged.connect(step)
        anim.finished.connect(done)
        self._strike_anim = anim
        anim.start()

    def _apply_strikethrough(self):
        self._strikethrough = not self._strikethrough
        self.task['strikethrough'] = int(self._strikethrough)
        font = self.name_edit.font()
        font.setStrikeOut(self._strikethrough)
        self.name_edit.setFont(font)
        self.name_edit._update_height()
        update_task(self.task['id'], self.task['name'], self.task['deadline'],
                    strikethrough=int(self._strikethrough),
                    priority=self.task.get('priority', 0),
                    recurrence=self.task.get('recurrence', ''))
        self.apply_urgency()          # 완료하면 숨쉬기·박동이 멈춤

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
            accent = f'QPushButton {{ border: none; color: {self._accent}; background: transparent; padding: 0; }}' \
                     f'QPushButton:pressed {{ color: {self._accent}; }}' if self._attach_expanded else \
                     f'QPushButton {{ border: none; color: {tct("#555555")}; background: transparent; padding: 0; }}' \
                     f'QPushButton:pressed {{ color: {self._accent}; }}'
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
        header.setStyleSheet(f'color: {tct("#666")}; background: transparent;')
        header.setFixedHeight(round(16 * self._scale))
        self._attach_layout.addWidget(header)
        for p in self._attachments:
            self._attach_layout.addWidget(self._make_attach_row(p))
        self._refresh_attach_btn()

    def _make_attach_row(self, path):
        s = self._scale
        row = QWidget()
        row.setFixedHeight(round(26 * s))
        row.setStyleSheet(f'QWidget {{ background: {tc("rgba(255,255,255,0.4)")}; border-radius: 4px; }}')
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
            f'QPushButton {{ border: none; background: transparent; color: {tct("#999")}; font-size: 11px; }}'
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
            from .calendar_view import P, _primary_button
            self._date_picker_popup = QWidget()
            self._date_picker_popup.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
            self._date_picker_popup.setAttribute(Qt.WA_TranslucentBackground)
            _outer = QVBoxLayout(self._date_picker_popup)
            _outer.setContentsMargins(0, 0, 0, 0)
            _card = QFrame()
            _card.setObjectName('PickerCard')
            _card.setStyleSheet(f"QFrame#PickerCard {{ background: {P.BG}; border: 1px solid {P.LINE}; border-radius: 12px; }}")
            _outer.addWidget(_card)

            layout = QVBoxLayout(_card)
            layout.setContentsMargins(10, 8, 10, 10)
            layout.setSpacing(8)

            self._cal_popup = CustomCalendarWidget()
            self._cal_popup.setMinimumSize(px(260), px(220))
            layout.addWidget(self._cal_popup)

            time_row = QHBoxLayout()
            time_row.setSpacing(6)
            
            self._time_checkbox = QCheckBox('시간')
            self._time_checkbox.setFont(pr_font(10))
            self._time_checkbox.setStyleSheet(f'color: {P.TEXT2}; background: transparent;')
            self._time_checkbox.stateChanged.connect(self._on_time_checkbox_changed)
            time_row.addWidget(self._time_checkbox)

            self._picker_time_edit = QTimeEdit(QTime(9, 0))
            self._picker_time_edit.setDisplayFormat('HH:mm')
            self._picker_time_edit.setFont(pr_font(11))
            self._picker_time_edit.setFixedWidth(QFontMetrics(self._picker_time_edit.font()).horizontalAdvance('00:00') + 28)
            self._picker_time_edit.setEnabled(False)
            self._picker_time_edit.setStyleSheet(f"""
                QTimeEdit {{
                    background: {P.CHIP}; border: 1px solid transparent; border-radius: 7px;
                    padding: 3px 8px; color: {P.TEXT}; qproperty-alignment: AlignCenter;
                }}
                QTimeEdit:focus {{ background: {P.BG}; border: 1px solid {P.TEXT2}; }}
                QTimeEdit:disabled {{ background: transparent; color: {P.FAINT}; }}
                QTimeEdit::up-button {{ width: 0px; }}
                QTimeEdit::down-button {{ width: 0px; }}
            """)
            time_row.addWidget(self._picker_time_edit)
            time_row.addStretch()

            self._confirm_btn = _primary_button('확인')
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
            self.btn_star.setIcon(_svg_icon(self._star_svg, tc('#D4A800'), self._star_sz))
            self.btn_star.setIconSize(QSize(self._star_sz, self._star_sz))
            self.btn_star.setStyleSheet("""
                QPushButton { border: none; padding: 0; }
                QPushButton:pressed { padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }
            """)
        else:
            self.btn_star.setIcon(QIcon())
            self.btn_star.setText(MI.STAR_BORDER)
            self.btn_star.setFont(mi_font(int(15 * self._scale)))
            self.btn_star.setStyleSheet(f"""
                QPushButton {{ color: {tc("rgba(0,0,0,0.25)")}; border: none; padding: 0; }}
                QPushButton:hover {{ color: {tc("rgba(0,0,0,0.5)")}; }}
                QPushButton:pressed {{ padding-top: 2px; padding-left: 1px; padding-bottom: 0px; padding-right: 0px; }}
            """)
        set_task_priority(self.task['id'], new)
        from . import motion
        if motion.load()['enabled'] and self.isVisible():
            self._pop_star(lambda: self._update_with_slide())
        else:
            self.on_update()

    def _pop_star(self, then):
        """별이 톡 커졌다 돌아온다 (0.22초)."""
        import math
        from PyQt5.QtCore import QVariantAnimation, QEasingCurve
        base = self.btn_star.iconSize() if not self.btn_star.icon().isNull() else QSize(self._star_sz, self._star_sz)
        font0 = self.btn_star.font()
        anim = QVariantAnimation(self)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setDuration(220)
        anim.setEasingCurve(QEasingCurve.OutQuad)

        def step(v):
            k = 1 + 0.35 * math.sin(math.pi * v)
            try:
                self.btn_star.setIconSize(QSize(int(base.width() * k), int(base.height() * k)))
                if self.btn_star.icon().isNull():
                    f = QFont(font0)
                    f.setPointSizeF(max(1.0, font0.pointSizeF() * k))
                    self.btn_star.setFont(f)
            except RuntimeError:
                pass
        anim.valueChanged.connect(step)
        anim.finished.connect(then)
        self._star_anim = anim
        anim.start()

    def _update_with_slide(self):
        """목록을 다시 그리면서, 이 업무가 옛 자리에서 새 자리로 미끄러져 가는 모습을 보여준다."""
        win = self.window()
        lst = self.parentWidget()
        try:
            win._pending_slide = (self.task['id'], self.grab(), self.mapTo(lst, QPoint(0, 0)))
        except Exception:
            win._pending_slide = None
        self.on_update()

    def _set_recurrence(self, recurrence):
        self.task['recurrence'] = recurrence
        set_task_recurrence(self.task['id'], recurrence)
        self.on_update()

    def _open_custom_recurrence(self):
        dlg = CustomRecurrenceDialog(self, current=self.task.get('recurrence', ''))
        if dlg.exec_() == QDialog.Accepted:
            self._set_recurrence(dlg.get_value())
