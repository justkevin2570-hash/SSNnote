import os
from datetime import date
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QApplication,
    QScrollArea, QFrame, QDateEdit, QDateTimeEdit, QMenu, QAction, QWidgetAction,
    QMessageBox, QDialog, QGridLayout, QCalendarWidget, QToolButton,
    QPlainTextEdit, QSizePolicy, QSplitter, QGraphicsColorizeEffect
)
from PyQt5.QtCore import Qt, QDate, QTime, QEvent, QTimer, QDateTime, QPoint, QPointF, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QColor, QPainter, QTextCharFormat, QPalette, QTextOption, QTextLayout, QIcon, QPixmap
from db import (update_window, delete_window, get_tasks, add_task, delete_task, update_task,
                add_task_history, get_task_history, delete_task_history,
                get_documents, add_document, update_document, delete_document)
from autostart import is_enabled as autostart_is_enabled, set_enabled as autostart_set
from capture import ScreenCaptureOverlay, run_ocr, grab_fullscreen, OcrWorker, _normalize_doc_number

TITLE_BAR_HEIGHT = 40
TITLE_COLOR      = '#f7c948'
SNAP_THRESHOLD   = 20  # 완전히 붙는 거리(px)
SNAP_ZONE        = 50  # 당기기 시작하는 거리(px)

PALETTE = [
    '#FDD663', '#FEFFA7', '#CCFF90', '#A8F0E8',
    '#AECBFA', '#D7AEFB', '#FDCFE8', '#E6C9A8', '#AAAAAA', '#DDDDDD', '#FFFFFF',
]



class AddDateButton(QLabel):
    """deadline이 없는 태스크에서 날짜 자리에 표시되는 클릭 가능한 레이블."""
    clicked = pyqtSignal()

    _STYLE_NORMAL = (
        'color: #bbb; background: transparent;'
        'border: 1px dashed #ccc; border-radius: 3px;'
        'padding: 1px 5px; margin-bottom: 3px;'
    )
    _STYLE_HOVER = (
        'color: #888; background: rgba(212,184,0,0.1);'
        'border: 1px dashed #d4b800; border-radius: 3px;'
        'padding: 1px 5px; margin-bottom: 3px;'
    )

    def __init__(self, *args, **kwargs):
        super().__init__('＋날짜', *args, **kwargs)
        self.setCursor(Qt.PointingHandCursor)
        self.setFont(QFont('Malgun Gothic', 9))
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


class DdayLabel(QLabel):
    def __init__(self, dday_text, date_text, *args, **kwargs):
        super().__init__(dday_text, *args, **kwargs)
        self._dday_text = dday_text
        self._date_text = date_text

    def enterEvent(self, event):
        self.setText(self._date_text)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setText(self._dday_text)
        super().leaveEvent(event)


def calc_dday(deadline_str):
    try:
        deadline = date.fromisoformat(deadline_str)
    except ValueError:
        return '날짜오류'
    diff = (deadline - date.today()).days
    if diff == 0:
        return 'D-day'
    elif diff > 0:
        return f'D-{diff}'
    else:
        return f'D+{abs(diff)}'


class EdgeHandle(QWidget):
    EDGE  = 6   # 가장자리 감지 두께(px)
    CORNER = 16  # 모서리 감지 크기(px)
    MIN_W = 200
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
            self.grabKeyboard()

    def mouseMoveEvent(self, e):
        if not (e.buttons() == Qt.LeftButton and self._drag_start):
            return
        d = e.globalPos() - self._drag_start
        g = self._start_geom
        win = self.window()

        if self.edge == 'left':
            new_w = g.width() - d.x()
            if new_w >= self.MIN_W:
                win.setGeometry(g.x() + d.x(), g.y(), new_w, g.height())

        elif self.edge == 'right':
            win.resize(max(self.MIN_W, g.width() + d.x()), g.height())

        elif self.edge == 'bottom':
            win.resize(g.width(), max(self.MIN_H, g.height() + d.y()))

        elif self.edge == 'bottom-right':
            win.resize(
                max(self.MIN_W, g.width() + d.x()),
                max(self.MIN_H, g.height() + d.y()),
            )

        elif self.edge == 'bottom-left':
            new_w = g.width() - d.x()
            new_h = max(self.MIN_H, g.height() + d.y())
            if new_w >= self.MIN_W:
                win.setGeometry(g.x() + d.x(), g.y(), new_w, new_h)
            else:
                win.resize(g.width(), new_h)

        elif self.edge == 'top':
            new_h = g.height() - d.y()
            if new_h >= self.MIN_H:
                win.setGeometry(g.x(), g.y() + d.y(), g.width(), new_h)

        elif self.edge == 'top-left':
            new_w = g.width() - d.x()
            new_h = g.height() - d.y()
            new_x = g.x() + d.x() if new_w >= self.MIN_W else g.x()
            new_y = g.y() + d.y() if new_h >= self.MIN_H else g.y()
            new_w = max(self.MIN_W, new_w)
            new_h = max(self.MIN_H, new_h)
            win.setGeometry(new_x, new_y, new_w, new_h)

        elif self.edge == 'top-right':
            new_w = max(self.MIN_W, g.width() + d.x())
            new_h = g.height() - d.y()
            if new_h >= self.MIN_H:
                win.setGeometry(g.x(), g.y() + d.y(), new_w, new_h)
            else:
                win.resize(new_w, g.height())

    def mouseReleaseEvent(self, e):
        self.releaseKeyboard()
        self._drag_start = None
        self._start_geom = None
        self.window().save_state()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape and self._drag_start:
            self.window().setGeometry(self._start_geom)
            self.releaseKeyboard()
            self._drag_start = None
            self._start_geom = None


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
                font-size: 19px;
                padding: 4px 6px 0px 6px;
                border-radius: 4px;
            }}
            QPushButton:hover {{ background: rgba(0,0,0,0.12); border-radius: 3px; }}
            QPushButton:pressed {{ background: rgba(0,0,0,0.22); border-radius: 3px; padding: 2px 0 0 2px; }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 8, 0)
        layout.setSpacing(2)

        self.label = QLabel('서서니 노트 &nbsp;&nbsp;<span style="font-size:10pt; font-weight:bold; font-family:Consolas; font-style:italic;">v1.6</span>')
        self.label.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        self.label.setStyleSheet('color: #5a4000; background: transparent;')

        btn_menu  = QPushButton('…')
        btn_menu.setToolTip('메뉴')
        btn_menu.setStyleSheet('font-size: 19px; letter-spacing: 0px; padding: 4px 6px 0px 6px;')
        self.btn_pin   = QPushButton('📌')
        self.btn_pin.setToolTip('항상 위 꺼짐')
        _gray = QGraphicsColorizeEffect()
        _gray.setColor(QColor('#888888'))
        _gray.setStrength(1.0)
        self.btn_pin.setGraphicsEffect(_gray)
        _x_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '엑스아이콘.png')
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

        layout.addWidget(self.label)
        layout.addStretch()
        layout.addWidget(btn_menu)
        layout.addWidget(self.btn_pin)
        layout.addWidget(self.btn_close)

        btn_menu.clicked.connect(lambda: parent.show_menu(btn_menu))

        self.btn_pin.clicked.connect(parent.toggle_always_on_top)
        self.btn_close.clicked.connect(parent.close)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPos() - self.parent.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.LeftButton and self._drag_pos:
            new_pos = e.globalPos() - self._drag_pos
            self.parent.move(self.parent._snap_to_screen(new_pos))

    def mouseReleaseEvent(self, e):
        self._drag_pos = None
        self.parent.save_state()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.parent.toggle_shade()


class RotatedMenuButton(QPushButton):
    """'…' 을 90도 회전해서 세로로 표시하는 버튼"""
    def __init__(self, parent=None):
        super().__init__('…', parent)
        self._hovered = False

    def enterEvent(self, e):
        self._hovered = True
        self.update()

    def leaveEvent(self, e):
        self._hovered = False
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
        self.document().contentsChanged.connect(self._update_height)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.clearFocus()
        else:
            super().keyPressEvent(e)

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
        self.setFixedHeight(max(h, 28))


class TaskRow(QWidget):
    def __init__(self, task, on_delete, on_update):
        super().__init__()
        self.task      = task
        self.on_update = on_update
        self.setStyleSheet('background: transparent;')

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 8, 0)
        layout.setSpacing(4)

        deadline = task['deadline']
        overdue = False
        if deadline:
            dday = calc_dday(deadline)
            if dday.startswith('D+'):
                overdue = True
                dday_color = '#aaa'
            elif dday == 'D-day':
                dday_color = '#e74c3c'
            elif dday == '날짜오류':
                dday_color = '#aaa'
            else:
                days_left = int(dday[2:])
                dday_color = '#e74c3c' if days_left <= 5 else '#333'
            suffix = f'({dday})'
        else:
            dday_color = '#333'
            suffix = ''

        self._strikethrough = bool(task.get('strikethrough', 0))
        name_color = '#aaa' if overdue else '#111'

        self.name_edit = _AutoHeightEdit(task['name'])
        name_font = QFont('Malgun Gothic', 12)
        name_font.setPointSizeF(10.5 if overdue else 12)
        if overdue:
            name_font.setItalic(True)
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
                d = date.fromisoformat(deadline)
                if suffix == '(D-day)':
                    date_text = '언능 하소!'
                else:
                    weekday = ['월','화','수','목','금','토','일'][d.weekday()]
                    date_text = f'{d.year}. {d.month}. {d.day}.({weekday})'
            except ValueError:
                date_text = suffix
            dday_lbl = DdayLabel(suffix, date_text)
            dday_font = QFont('Malgun Gothic', 12, QFont.Bold)
            dday_font.setPointSizeF(10.5 if overdue else 12)
            if overdue:
                dday_font.setItalic(True)
            dday_lbl.setFont(dday_font)
            dday_lbl.setStyleSheet(f'color: {dday_color}; background: transparent; margin-bottom: 3px;')
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

        layout.addWidget(self.name_edit, 1)
        if dday_lbl:
            layout.addWidget(dday_lbl, 0, Qt.AlignTop)
        elif hasattr(self, '_add_date_btn'):
            layout.addWidget(self._add_date_btn, 0, Qt.AlignTop)
        layout.addWidget(btn_menu, 0, Qt.AlignTop)

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
            update_task(self.task['id'], name, self.task['deadline'], strikethrough=int(self._strikethrough))
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
        update_task(self.task['id'], self.task['name'], self.task['deadline'], strikethrough=int(self._strikethrough))

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
                    strikethrough=int(self._strikethrough))
        self.on_update()


from PyQt5.QtWidgets import QSplitterHandle

class DragHandleSplitter(QSplitter):
    """가운데에 드래그 핸들 점을 그리는 커스텀 스플리터."""
    def createHandle(self):
        return _DotHandle(self.orientation(), self)

class _DotHandle(QSplitterHandle):
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        dot_color = QColor('#a07000')
        painter.setBrush(dot_color)
        painter.setPen(Qt.NoPen)
        cx = self.width() // 2
        cy = self.height() // 2
        r = 2  # 점 반지름
        gap = 6  # 점 간격
        for i in (-2, -1, 0, 1, 2):
            painter.drawEllipse(cx + i * gap - r, cy - r, r * 2, r * 2)
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
        _edit_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '수정 아이콘.png')
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
            return

        import os
        from PyQt5.QtGui import QIcon, QPixmap
        from PyQt5.QtCore import QSize
        _del_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '엑스아이콘.png')
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


class MemoWindow(QMainWindow):
    def __init__(self, window_id, on_new=None, open_windows=None):
        super().__init__()
        self.window_id       = window_id
        self.on_new          = on_new or (lambda **kw: None)
        self._open_windows   = open_windows if open_windows is not None else []
        self.collapsed       = False
        self.expanded_height = 400
        self.pin_active      = False
        self._bg_color       = '#FEFFA7'  # 기본값
        self._capture_hint_shown = False

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self._build_ui()
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self._pin_shortcut = QShortcut(QKeySequence('Ctrl+Shift+F'), self)
        self._pin_shortcut.activated.connect(self.toggle_always_on_top)
        self._shade_shortcut = QShortcut(QKeySequence('Ctrl+Shift+R'), self)
        self._shade_shortcut.activated.connect(self.toggle_shade)
        self._shortcuts = [self._pin_shortcut, self._shade_shortcut]

    def apply_state(self, x, y, width, height, collapsed, color=''):
        self.expanded_height = height
        self.collapsed       = collapsed
        self.move(x, y)
        self.resize(width, TITLE_BAR_HEIGHT if collapsed else height)
        if collapsed:
            self.content.hide()
            self.setFixedHeight(TITLE_BAR_HEIGHT)
        if color:
            self._apply_color(color)

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

        self.content = QWidget()
        self.content.setStyleSheet('background: transparent;')
        _outer_layout = QVBoxLayout(self.content)
        _outer_layout.setContentsMargins(0, 0, 8, 0)
        _outer_layout.setSpacing(0)

        _splitter = DragHandleSplitter(Qt.Vertical)
        _splitter.setStyleSheet("""
            QSplitter::handle {
                background: #d4b800;
                height: 10px;
            }
        """)
        _splitter.setChildrenCollapsible(False)
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
        field_style_date = """
            QDateEdit {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
            }
            QDateEdit::drop-down { width: 0px; border: none; background: transparent; }
            QDateEdit::up-button { width: 0px; }
            QDateEdit::down-button { width: 0px; }
        """
        lbl_font = QFont('Malgun Gothic', 11)
        lbl_css  = 'color: #5a4000; background: transparent;'

        # 1행: 업무
        row1 = QHBoxLayout()
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
        self.input_name.installEventFilter(self)
        self.input_name.returnPressed.connect(self._add_task)
        self.input_name.textChanged.connect(lambda t: self.input_date.setDate(QDate(1900, 1, 1)) or setattr(self, '_date_touched', False) if not t else None)
        row1.addWidget(lbl_name)
        row1.addWidget(self.input_name, 1)
        content_layout.addLayout(row1)

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
        self.input_date.setFixedWidth(145)
        self.input_date.setFont(QFont('Malgun Gothic', 11))
        self.input_date.setStyleSheet(field_style_date)
        self.input_date.dateChanged.connect(lambda: setattr(self, '_date_touched', True))
        self.input_date.installEventFilter(self)
        for child in self.input_date.findChildren(QWidget):
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

        row2.addWidget(lbl_date)
        row2.addWidget(self.input_date)
        row2.addStretch()
        row2.addWidget(btn_history)
        content_layout.addWidget(self.date_row_widget)

        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet('color: #d4b800; background: #d4b800; max-height: 1px; margin: 2px 8px;')
        content_layout.addWidget(line)

        # 업무 목록
        self.task_list_widget = QWidget()
        self.task_list_widget.setStyleSheet('background: transparent;')
        self.task_list_layout = QVBoxLayout(self.task_list_widget)
        self.task_list_layout.setContentsMargins(0, 0, 0, 0)
        self.task_list_layout.setSpacing(0)
        self.task_list_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setWidget(self.task_list_widget)
        scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical { width: 6px; background: transparent; }
            QScrollBar::handle:vertical { background: #d4b800; border-radius: 3px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)
        content_layout.addWidget(scroll)
        _splitter.addWidget(_top_panel)

        # ── 하단 패널: 공문번호 섹션 ──────────────────────────────
        _bottom_panel = QWidget()
        _bottom_panel.setStyleSheet('background: transparent;')
        bottom_layout = QVBoxLayout(_bottom_panel)
        bottom_layout.setContentsMargins(0, 4, 0, 4)
        bottom_layout.setSpacing(2)

        # 헤더: "공문번호" 레이블 + 캡처 버튼
        doc_header = QHBoxLayout()
        doc_header.setContentsMargins(8, 0, 8, 0)
        lbl_doc = QLabel('공문 캡쳐')
        lbl_doc.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        lbl_doc.setStyleSheet('color: #5a4000; background: transparent;')
        self.lbl_capture_result = ClickToCopyLabel('')
        self.lbl_capture_result.setFont(QFont('Malgun Gothic', 10))
        self.lbl_capture_result.setStyleSheet("""
            QLabel {
                background: rgba(255,255,255,0.25);
                border: 1px solid #c0a800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #5a4000;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """)
        self.lbl_capture_result.setCursor(Qt.PointingHandCursor)
        btn_capture = QPushButton('캡처')
        btn_capture.setFont(QFont('Malgun Gothic', 10))
        btn_capture.setFixedWidth(50)
        btn_capture.setStyleSheet("""
            QPushButton {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #5a4000;
            }
            QPushButton:hover { background: rgba(212,184,0,0.25); }
        """)
        btn_capture.clicked.connect(self._start_capture)
        btn_shortcut = QPushButton('단축키 ON')
        btn_shortcut.setFont(QFont('Malgun Gothic', 9, QFont.Bold))
        btn_shortcut.setFixedWidth(85)
        btn_shortcut.setCheckable(True)
        btn_shortcut.setChecked(True)
        btn_shortcut.setStyleSheet("""
            QPushButton {
                background: rgba(255,255,255,0.5);
                border: 1px solid #d4b800;
                border-radius: 4px;
                padding: 2px 4px;
                color: #999999;
                font-weight: normal;
            }
            QPushButton:hover { background: rgba(212,184,0,0.25); }
            QPushButton:checked {
                background: rgba(212,184,0,0.35);
                border: 1px solid #a08800;
                color: #3a2800;
                font-weight: bold;
            }
        """)
        def _toggle_shortcut(checked):
            for sc in self._shortcuts:
                sc.setEnabled(checked)
            btn_shortcut.setText('단축키 ON' if checked else '단축키 OFF')
        btn_shortcut.toggled.connect(_toggle_shortcut)
        self._btn_shortcut = btn_shortcut

        self.lbl_capture_status = QLabel('')
        self.lbl_capture_status.setFont(QFont('Malgun Gothic', 10, QFont.Bold))
        self.lbl_capture_status.setStyleSheet('color: #1a56cc; background: transparent;')
        self.lbl_capture_status.hide()
        doc_header.addWidget(lbl_doc)
        doc_header.addSpacing(6)
        doc_header.addWidget(self.lbl_capture_status)
        doc_header.addStretch()
        doc_header.addWidget(btn_shortcut)
        doc_header.addSpacing(4)
        doc_header.addWidget(btn_capture)
        bottom_layout.addLayout(doc_header)

        doc_line = QFrame()
        doc_line.setFrameShape(QFrame.HLine)
        doc_line.setStyleSheet('color: #d4b800; background: #d4b800; max-height: 1px; margin: 0 8px;')
        bottom_layout.addWidget(doc_line)

        # 공문 목록 스크롤 영역
        self.doc_list_widget = QWidget()
        self.doc_list_widget.setStyleSheet('background: transparent;')
        self.doc_list_layout = QVBoxLayout(self.doc_list_widget)
        self.doc_list_layout.setContentsMargins(0, 0, 0, 0)
        self.doc_list_layout.setSpacing(0)
        self.doc_list_layout.addStretch()

        doc_scroll = QScrollArea()
        doc_scroll.setWidgetResizable(True)
        doc_scroll.setWidget(self.doc_list_widget)
        doc_scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical { width: 6px; background: transparent; }
            QScrollBar::handle:vertical { background: #d4b800; border-radius: 3px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)
        bottom_layout.addWidget(doc_scroll)

        _splitter.addWidget(_bottom_panel)
        _splitter.setStretchFactor(0, 2)
        _splitter.setStretchFactor(1, 1)

        root.addWidget(self.content)
        self._apply_color(self._bg_color)
        self._refresh_tasks()
        self._refresh_documents()
        self._schedule_midnight_refresh()
        self._setup_resize_handles()

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
                    self._date_touched = False
                    self.input_name.setFocus()
                    return True
                if event.key() == Qt.Key_Tab:
                    if self.input_date.currentSection() == QDateTimeEdit.DaySection:
                        self.input_name.setFocus()
                        return True
        return super().eventFilter(obj, event)

    def _add_task(self):
        name = self.input_name.text().strip()
        if not name:
            return
        deadline = self.input_date.date().toString('yyyy-MM-dd') if self._date_touched else ''
        add_task(self.window_id, name, deadline)
        self.input_name.clear()
        self.input_date.setDate(QDate(1900, 1, 1))
        self._date_touched = False
        self.input_name.setFocus()
        self._refresh_tasks()

    def _delete_task(self, task):
        add_task_history(self.window_id, task['name'], task['deadline'], strikethrough=task.get('strikethrough', 0))
        delete_task(task['id'])
        self._refresh_tasks()

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

    def _on_midnight(self):
        self._refresh_tasks()
        self._schedule_midnight_refresh()  # 다음 자정 예약

    def _refresh_tasks(self):
        while self.task_list_layout.count() > 1:
            item = self.task_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for task in get_tasks(self.window_id):
            row = TaskRow(task, self._delete_task, self._refresh_tasks)
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
                self.raise_()
                self.activateWindow()
            else:
                self.lbl_capture_result.setText('')
                QMessageBox.information(self, '알림', 'OCR 텍스트를 인식하지 못했습니다.')

        def _on_error(msg):
            QMessageBox.warning(self, 'OCR 오류', msg)

        self._ocr_worker = run_ocr(pixmap, _after_ocr, _on_error)

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
        autostart_on = autostart_is_enabled()
        act_auto   = QAction(('✅' if autostart_on else '☐') + ' 시작 시 자동실행', self)
        act_help   = QAction('💡 도움말', self)
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

        menu.addAction(act_new)
        menu.addMenu(color_menu)
        menu.addAction(act_auto)
        menu.addAction(act_help)
        menu.addSeparator()
        menu.addAction(act_delete)

        act_new.triggered.connect(lambda: self.on_new(offset_from=self))
        act_auto.triggered.connect(lambda: autostart_set(not autostart_on))
        act_help.triggered.connect(self.show_help)
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
                font-size: 19px;
                padding: 4px 6px 0px 6px;
                border-radius: 4px;
            }}
            QPushButton:hover {{ background: rgba(0,0,0,0.12); }}
        """)

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
                add_task(self.window_id, r['name'], r['deadline'], strikethrough=r.get('strikethrough', 0))
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
        self._date_touched = True
        self._cal_popup.hide()

    def show_help(self):
        msg = QMessageBox(self)
        msg.setWindowTitle('도움말')
        msg.setText(
            '<p style="line-height: 195%; font-family: Malgun Gothic; font-size: 11pt;">'
            '① Tab키, Enter키만 잘 쓰면 편하게 쓰실 수 있습니다.<br>'
            '② 단축키: <b>Ctrl + Shift</b> 를 누른 후<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>X</b> : 화면 캡처<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>S</b> : 포커싱<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>R</b> : 롤업<br>'
            '&nbsp;&nbsp;&nbsp;&nbsp;- <b>F</b> : 항상 위<br>'
            '③ 캡처는 공문을 전체화면으로 키워야 정확합니다.<br>'
            '④ 캡처하고 공문 제목과 공문 번호란 클릭하시면 붙여넣기 됩니다.<br>'
            '⑤ 공문 제목과 공문 번호란을 클릭하시면 복사가 됩니다.'
            '</p>'
        )
        msg.setStyleSheet(
            "font-family: 'Malgun Gothic'; font-size: 11pt;"
            "QMessageBox { min-width: 370px; }"
        )
        msg.exec_()

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

    def _snap_to_screen(self, pos):
        screen = QApplication.screenAt(pos) or QApplication.primaryScreen()
        rect = screen.availableGeometry()
        x, y = float(pos.x()), float(pos.y())
        w, h = self.width(), self.height()

        def snap_axis(val, lo, hi, size):
            dist_lo = val - lo          # 왼쪽/상단 끝까지의 거리
            dist_hi = hi - (val + size) # 오른쪽/하단 끝까지의 거리
            if 0 <= dist_lo < SNAP_ZONE:
                if dist_lo < SNAP_THRESHOLD:
                    return float(lo)
                t = 1.0 - dist_lo / SNAP_ZONE   # 가까울수록 t가 커짐
                return val - dist_lo * t * t     # 2차 곡선으로 서서히 당김
            if 0 <= dist_hi < SNAP_ZONE:
                if dist_hi < SNAP_THRESHOLD:
                    return float(hi - size)
                t = 1.0 - dist_hi / SNAP_ZONE
                return val + dist_hi * t * t
            return val

        nx = snap_axis(x, rect.left(), rect.right(),  w)
        ny = snap_axis(y, rect.top(),  rect.bottom(), h)
        return QPoint(int(nx), int(ny))

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

    def save_state(self):
        pos = self.pos()
        update_window(
            self.window_id,
            x=pos.x(), y=pos.y(),
            width=self.width(),
            height=self.expanded_height,
            collapsed=self.collapsed,
            color=self._bg_color,
        )

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, '_first_shown', False):
            self._first_shown = True
            QTimer.singleShot(0, lambda: (self.activateWindow(), self.input_name.setFocus()))

    def closeEvent(self, e):
        if not getattr(self, '_deleted', False):
            self.save_state()
        e.accept()
