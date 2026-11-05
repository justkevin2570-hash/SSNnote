"""ui/widgets.py — 여러 화면에서 재사용하는 작은 위젯들 (라벨, 입력칸, 달력, 시간 선택 등)."""
from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QFrame, QCalendarWidget, QTextEdit, QListWidget, QAbstractItemView, QListWidgetItem
)
from PyQt5.QtCore import Qt, QDate, QTime, QTimer, QPoint, QSize, pyqtSignal
from PyQt5.QtGui import QFontMetrics, QColor, QPainter, QTextCharFormat, QTextCursor, QKeySequence
from .common import pr_font
from .theme import T, tc, tct, qcolor

class AddDateButton(QLabel):
    """deadline이 없는 태스크에서 날짜 자리에 표시되는 클릭 가능한 레이블."""
    clicked = pyqtSignal()

    _STYLE_NORMAL = (
        f'color: {tct("#bbb")}; background: transparent;'
        f'border: 1px dashed {tc("#ccc")}; border-radius: 3px;'
        'padding: 1px 5px;'
    )
    _STYLE_HOVER = (
        f'color: {tct("#888")}; background: {tc("rgba(212,184,0,0.1)")};'
        f'border: 1px dashed {T.ACCENT}; border-radius: 3px;'
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
    clicked = pyqtSignal()

    def __init__(self, dday_text, date_text, *args, **kwargs):
        super().__init__(dday_text, *args, **kwargs)
        self._dday_text = dday_text
        self._date_text = date_text
        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

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
        if T.DANGER in self._original_style:
            self.setStyleSheet(self._original_style.replace(T.DANGER, tc('#e96060')))
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setText(self._dday_text)
        f = self.font()
        f.setPointSizeF(f.pointSizeF() + 0.3)
        self.setFont(f)
        if hasattr(self, '_original_style'):
            self.setStyleSheet(self._original_style)
        super().leaveEvent(event)


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
            p.setBrush(qcolor('#000000', 30))  # 다크에선 흰 음영
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(self.rect(), 4, 4)

        p.translate(self.width() / 2.0, self.height() / 2.0)
        p.rotate(90)
        p.setFont(self.font())
        p.setPen(qcolor('#000000'))
        p.drawText(
            -self.height() // 2, -self.width() // 2,
            self.height(), self.width(),
            Qt.AlignCenter, self.text()
        )
        p.end()


class PlainPasteMemoEdit(QTextEdit):
    def insertFromMimeData(self, source):
        if not source.hasText():
            super().insertFromMimeData(source)
            return
        cursor = self.textCursor()
        start = cursor.position()
        cursor.insertText(source.text())
        cursor.setPosition(start, QTextCursor.KeepAnchor)
        char_fmt = QTextCharFormat()
        char_fmt.setFont(self.font())
        cursor.mergeCharFormat(char_fmt)
        cursor.clearSelection()
        self.setTextCursor(cursor)


class _SelectableLabel(QLabel):
    """마우스 드래그로 텍스트 선택 가능한 라벨. 선택 후 Ctrl+C 복사 지원."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        self.setFocusPolicy(Qt.ClickFocus)

    def keyPressEvent(self, e):
        if e.matches(QKeySequence.Copy) and self.hasSelectedText():
            QApplication.clipboard().setText(self.selectedText())
            return
        super().keyPressEvent(e)


class _FixedDotBar(QWidget):
    """드래그 불가 고정 구분선 (황금색 바 + 점 장식)."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(10)
        self._color = T.ACCENT

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
            painter.setPen(QColor(T.ACCENT))
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(rect.adjusted(1, 1, -2, -2))
            painter.restore()


class TimePickerPopup(QFrame):
    time_selected = pyqtSignal(QTime)

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.setStyleSheet(f"""
            QFrame {{
                background: {tc("#fffbe6")};
                border: 1px solid {T.ACCENT};
                border-radius: 6px;
            }}
            QListWidget {{
                background: transparent;
                border: none;
                font-family: 'Malgun Gothic';
                font-size: 11pt;
                color: {tct("#333")};
                outline: 0;
            }}
            QListWidget::item {{ padding: 3px 10px; border-radius: 3px; }}
            QListWidget::item:selected {{ background: {T.ACCENT}; color: {tct("white")}; }}
            QListWidget::item:hover {{ background: {tc("#fff3a0")}; }}
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
        popup.setStyleSheet(f"""
            QLabel {{
                background: {T.TEXT_ON_NOTE};
                color: {tct("white")};
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
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
        popup.setStyleSheet(f"""
            QLabel {{
                background: {T.TEXT_ON_NOTE};
                color: {tct("white")};
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
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
        popup.setStyleSheet(f"""
            QLabel {{
                background: {T.TEXT_ON_NOTE};
                color: {tct("white")};
                border-radius: 4px;
                padding: 3px 8px;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
        """)
        popup.adjustSize()
        popup.move(global_pos + QPoint(4, -popup.height() - 8))
        popup.show()
        QTimer.singleShot(500, popup.deleteLater)
