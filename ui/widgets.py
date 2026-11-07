"""ui/widgets.py — 여러 화면에서 재사용하는 작은 위젯들 (라벨, 입력칸, 달력, 시간 선택 등)."""
from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QFrame, QCalendarWidget, QTextEdit, QListWidget, QAbstractItemView, QListWidgetItem
)
from PyQt5.QtCore import Qt, QDate, QTime, QTimer, QPoint, QPointF, QRectF, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QColor, QPainter, QTextCharFormat, QTextCursor, QKeySequence
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

    _STYLE_HIDDEN = ('color: transparent; background: transparent;'
                     'border: 1px dashed transparent; border-radius: 3px; padding: 1px 5px;')

    def __init__(self, *args, **kwargs):
        super().__init__('＋날짜', *args, **kwargs)
        from .theme import current_accent, rgba
        ac = current_accent()      # 이 메모 창 색에 맞춘 포인트 색
        self._STYLE_HOVER = (
            f'color: {tct("#777")}; background: {rgba(ac, 0.12)};'
            f'border: 1px dashed {ac}; border-radius: 3px; padding: 1px 5px;'
        )
        self.setCursor(Qt.PointingHandCursor)
        self.setFont(pr_font(10))
        self._row_hover = False
        self._self_hover = False
        self._restyle()

    def set_row_hover(self, on: bool):
        """업무 줄에 마우스를 올렸을 때만 보인다 (평소엔 자리만 차지 — 줄이 움직이지 않게)."""
        self._row_hover = on
        self._restyle()

    def _restyle(self):
        if self._self_hover:
            self.setStyleSheet(self._STYLE_HOVER)
        elif self._row_hover:
            self.setStyleSheet(self._STYLE_NORMAL)
        else:
            self.setStyleSheet(self._STYLE_HIDDEN)

    def enterEvent(self, event):
        self._self_hover = True
        self._restyle()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._self_hover = False
        self._restyle()
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
    """D-day 표시. 급한 정도(tone)에 따라 둥근 바탕색으로 보여주고, 마우스를 올리면 날짜를 보여준다.

    tone: 'overdue'(지남, 빨강) · 'today'(오늘, 주황) · 'soon'(1~3일, 옅은 주황) ·
          'week'(4~7일, 옅은 호박색) · 'later'(그 뒤, 회색 글자만) · 'error'
    움직임(ui/motion.py): set_motion(breathe=주기초, beat=기한 지남 박동) — 둘레가 숨쉬듯 빛나거나
    왼쪽 빨간 점이 심장 박동처럼 뛴다.
    """
    clicked = pyqtSignal()

    # (바탕색, 글자색) — 바탕 None이면 글자만
    _TONES = {
        'overdue': ('#E5484D', '#FFFFFF'),
        'today':   ('#F76B15', '#FFFFFF'),
        'soon':    ('rgba(247,107,21,0.16)', '#C2410C'),
        'week':    ('rgba(176,112,0,0.15)', '#8A5A00'),
        'later':   (None, 'rgba(0,0,0,0.55)'),
        'error':   (None, 'rgba(0,0,0,0.35)'),
    }
    _GLOW = {'overdue': '#E5484D', 'today': '#F76B15', 'soon': '#F76B15', 'week': '#B07000',
             'later': '#8A8A8A', 'error': '#8A8A8A'}
    PAD = 5          # 숨쉬기 빛이 번질 자리

    def __init__(self, dday_text, date_text, *args, tone='later', align_left=False, **kwargs):
        super().__init__(dday_text, *args, **kwargs)
        self._align_left = align_left      # 목록 왼쪽에 둘 때(대시보드)는 왼쪽 정렬
        self._dday_text = dday_text
        self._date_text = date_text
        self._hover = False
        self._breathe = 0.0
        self._beat = False
        self.setCursor(Qt.PointingHandCursor)
        self.set_tone(tone)

    def set_tone(self, tone: str):
        self._tone = tone if tone in self._TONES else 'later'
        self.updateGeometry()
        self.update()

    def set_texts(self, dday_text, date_text=None):
        """남은 시간처럼 바뀌는 글자 (마우스를 올린 동안은 날짜를 그대로 보여줌)."""
        self._dday_text = dday_text
        if date_text is not None:
            self._date_text = date_text
        if not self._hover:
            self.setText(dday_text)
            self.updateGeometry()

    def set_motion(self, breathe: float = 0.0, beat: bool = False):
        from .motion import clock
        changed = (breathe, beat) != (self._breathe, self._beat)
        self._breathe, self._beat = breathe, beat
        if breathe or beat:
            clock().register(self, lambda t: self.update())
        else:
            clock().unregister(self)
        if changed:
            self.updateGeometry()
            self.update()

    def _colors(self):
        bg, fg = self._TONES[self._tone]

        def q(c, is_bg):
            if c is None:
                return None
            if c.startswith('rgba'):
                r, g, b, a = [x.strip() for x in c[5:-1].split(',')]
                col = QColor(int(r), int(g), int(b))
                col = QColor((tc if is_bg else tct)(col.name()))
                col.setAlphaF(float(a))
                return col
            if c == '#FFFFFF':
                return QColor(c)           # 진한 바탕 위의 흰 글자는 다크 모드에서도 흰색 그대로
            return QColor((tc if is_bg else tct)(c))
        return q(bg, True), q(fg, False)

    def _dot_space(self):
        return 12 if self._beat else 0

    def _pad(self):
        return self.PAD if (self._breathe or self._beat) else 1

    def sizeHint(self):
        fm = QFontMetrics(self.font())
        return QSize(fm.horizontalAdvance(self.text()) + 16 + 2 * self._pad() + self._dot_space(),
                     max(fm.height() + 4, 22) + 2 * self._pad())

    minimumSizeHint = sizeHint

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

    def paintEvent(self, event):
        from .motion import breath, clock, heartbeat
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        bg, fg = self._colors()
        fm = QFontMetrics(self.font())
        h = max(fm.height() + 2, 20)
        w = fm.horizontalAdvance(self.text()) + 14
        if self._align_left:
            x = self._pad() + self._dot_space()
        else:                              # 기본은 오른쪽 정렬 (줄마다 같은 끝에)
            x = self.width() - w - self._pad()
        y = (self.height() - h) / 2.0
        pill = QRectF(x, y, w, h)
        t = clock().now()
        if self._breathe and not self._hover:
            a = breath(t, self._breathe)
            # 표시가 살짝 부풀었다 돌아오고(최대 6%), 둘레에 같은 색 빛이 번진다
            k = 0.06 * a
            pill = pill.adjusted(-w * k / 2, -h * k / 2, w * k / 2, h * k / 2)
            glow = QColor(tc(self._GLOW[self._tone] if self._tone not in ('later', 'error') else '#E2A400'))
            for spread, alpha in ((self.PAD + 1.0, 0.22), (self.PAD * 0.6, 0.38)):   # 바깥은 옅게, 안쪽은 진하게
                glow.setAlphaF(alpha * a)
                p.setPen(Qt.NoPen)
                p.setBrush(glow)
                r = pill.adjusted(-spread * a, -spread * a, spread * a, spread * a)
                p.drawRoundedRect(r, r.height() / 2.0, r.height() / 2.0)
        if bg is not None:
            p.setPen(Qt.NoPen)
            p.setBrush(bg)
            p.drawRoundedRect(pill, pill.height() / 2.0, pill.height() / 2.0)
        if self._beat:
            hb = heartbeat(t)
            red = QColor(tc('#E5484D'))
            cx, cy = x - 7, y + h / 2.0
            red.setAlphaF(0.25 * hb)
            p.setPen(Qt.NoPen)
            p.setBrush(red)
            p.drawEllipse(QPointF(cx, cy), 3 + 3 * hb, 3 + 3 * hb)
            red.setAlphaF(1.0)
            p.setBrush(red)
            p.drawEllipse(QPointF(cx, cy), 2.6 + 0.8 * hb, 2.6 + 0.8 * hb)
        p.setFont(self.font())
        p.setPen(fg)
        p.drawText(pill, Qt.AlignCenter, self.text())
        p.end()

    def enterEvent(self, event):
        self._hover = True
        self.setText(self._date_text)
        self.updateGeometry()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.setText(self._dday_text)
        self.updateGeometry()
        super().leaveEvent(event)


class RotatedMenuButton(QPushButton):
    """'…' 을 90도 회전해서 세로로 표시하는 버튼"""
    hovered = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__('…', parent)
        self._hovered = False
        self._row_active = True     # 업무 줄에서는 TaskRow가 마우스 올림·선택 때만 켠다

    def set_row_active(self, on: bool):
        if on != self._row_active:
            self._row_active = on
            self.update()

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
        if not (self._row_active or self._hovered or self.isDown()):
            return                      # 평소엔 숨김 (자리는 유지)
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        if self._hovered:
            p.setBrush(qcolor('#000000', 30))  # 다크에선 흰 음영
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(self.rect(), 4, 4)

        p.translate(self.width() / 2.0, self.height() / 2.0)
        p.rotate(90)
        p.setFont(self.font())
        p.setPen(qcolor('#000000', 170))
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
    """작은 날짜 고르기 달력 (미니멀): 날짜를 직접 그린다.
    오늘은 빨간 동그라미, 고른 날은 진한 동그라미, 일요일 빨강·토요일 파랑, 다른 달 날짜는 숨김."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        style_mini_calendar(self)

    def paintCell(self, painter, rect, date):
        from .calendar_view import P, _ui_font
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(rect, QColor(P.BG))
        if date.month() != self.monthShown():          # 앞뒤 달 날짜는 아주 옅게
            painter.setFont(_ui_font(9.5))
            painter.setPen(QColor(P.FAINT))
            painter.drawText(rect, Qt.AlignCenter, str(date.day()))
            painter.restore()
            return
        today = date == QDate.currentDate()
        selected = date == self.selectedDate()
        d = min(rect.width(), rect.height()) - 6
        circle = QRectF(rect.center().x() - d / 2 + 0.5, rect.center().y() - d / 2 + 0.5, d, d)
        f = _ui_font(9.5, QFont.DemiBold if (today or selected) else QFont.Normal)
        if today or selected:
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(P.TODAY if today else P.PRIMARY))
            painter.drawEllipse(circle)
            color = QColor('#FFFFFF') if (today or P.PRIMARY_TEXT == '#FFFFFF') else QColor(P.PRIMARY_TEXT)
        elif date.dayOfWeek() == 7:
            color = QColor(P.SUN)
        elif date.dayOfWeek() == 6:
            color = QColor(P.SAT)
        else:
            color = QColor(P.TEXT)
        painter.setFont(f)
        painter.setPen(color)
        painter.drawText(rect, Qt.AlignCenter, str(date.day()))
        painter.restore()


def style_mini_calendar(cal):
    """QCalendarWidget을 미니멀하게: 흰 바탕, 격자 없음, 요일은 작은 회색 글자, 얇은 꺾쇠 화살표."""
    from .calendar_view import P, _ui_font
    from .common import icon_png
    cal.setGridVisible(False)
    cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
    cal.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
    from PyQt5.QtCore import QLocale
    cal.setLocale(QLocale(QLocale.Korean, QLocale.SouthKorea))
    cal.setFirstDayOfWeek(Qt.Sunday)
    hdr = QTextCharFormat()
    hdr.setForeground(QColor(P.MUTED))
    hdr.setFont(_ui_font(8.5, QFont.DemiBold))
    cal.setHeaderTextFormat(hdr)
    for day, color in ((Qt.Saturday, P.SAT), (Qt.Sunday, P.SUN)):
        f = QTextCharFormat()
        f.setForeground(QColor(color))
        f.setFont(_ui_font(8.5, QFont.DemiBold))
        cal.setWeekdayTextFormat(day, f)
    cal.setStyleSheet(f"""
        QCalendarWidget {{ background: {P.BG}; }}
        QCalendarWidget QWidget {{ background: {P.BG}; alternate-background-color: {P.BG}; }}
        QCalendarWidget QAbstractItemView {{
            background: {P.BG}; color: {P.TEXT}; selection-background-color: transparent;
            selection-color: {P.TEXT}; outline: 0; border: none;
        }}
        QCalendarWidget QWidget#qt_calendar_navigationbar {{
            background: {P.BG}; border-bottom: 1px solid {P.LINE}; padding: 4px 2px;
        }}
        QCalendarWidget QToolButton {{
            background: transparent; color: {P.TEXT}; border: none; border-radius: 6px;
            font-family: 'Pretendard GOV', 'Pretendard', 'Malgun Gothic'; font-size: 10pt; font-weight: 600;
            padding: 4px 6px;
        }}
        QCalendarWidget QToolButton:hover {{ background: {P.CHIP}; }}
        QCalendarWidget QToolButton::menu-indicator {{ image: none; width: 0; }}
        QCalendarWidget QToolButton#qt_calendar_prevmonth {{ qproperty-icon: url({icon_png('ph.caret-left', P.TEXT2, 16)}); }}
        QCalendarWidget QToolButton#qt_calendar_nextmonth {{ qproperty-icon: url({icon_png('ph.caret-right', P.TEXT2, 16)}); }}
        QCalendarWidget QSpinBox {{ background: {P.BG}; color: {P.TEXT}; border: 1px solid {P.LINE}; border-radius: 5px; }}
        QCalendarWidget QMenu {{ background: {P.BG}; color: {P.TEXT}; border: 1px solid {P.LINE}; }}
    """)


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
