"""ui/popup_toggle.py — 메뉴를 연 단추를 다시 누르면 닫히기만 하고 다시 열리지 않는다."""
from unittest import mock

from PyQt5.QtCore import Qt, QEvent, QPoint, QPointF
from PyQt5.QtGui import QMouseEvent
from PyQt5.QtWidgets import QApplication, QMenu, QPushButton, QWidget, QVBoxLayout

from ui import popup_toggle


def _press(w, typ=QEvent.MouseButtonPress):
    p = QPointF(5, 5)
    return QMouseEvent(typ, p, QPointF(w.mapToGlobal(QPoint(5, 5))), Qt.LeftButton,
                       Qt.LeftButton if typ != QEvent.MouseButtonRelease else Qt.NoButton, Qt.NoModifier)


def _setup():
    app = QApplication.instance()
    f = popup_toggle.PopupToggleFilter()
    app.installEventFilter(f)
    host = QWidget(); lay = QVBoxLayout(host)
    btn = QPushButton('⋮'); other = QPushButton('다른 단추')
    lay.addWidget(btn); lay.addWidget(other)
    host.show()
    opened = []
    menu = QMenu(host); menu.addAction('항목')

    def open_menu():
        opened.append(1)
        menu.popup(btn.mapToGlobal(btn.rect().bottomLeft()))
    btn.clicked.connect(open_menu)
    return app, f, host, btn, other, menu, opened


def _click(app, w):
    app.sendEvent(w, _press(w))
    app.sendEvent(w, _press(w, QEvent.MouseButtonRelease))
    app.processEvents()


def _close_by_click_on(app, menu, target):
    """팝업 바깥(target 위)을 눌러 닫히는 상황: 누른 채 닫힘 → Qt가 target에 누름을 다시 보냄."""
    with mock.patch.object(QApplication, 'mouseButtons', return_value=Qt.LeftButton), \
         mock.patch.object(QApplication, 'widgetAt', return_value=target):
        menu.hide()
        app.processEvents()
    _click(app, target)   # replay


def test_same_button_closes_without_reopening():
    app, f, host, btn, other, menu, opened = _setup()
    try:
        _click(app, btn)
        assert opened == [1] and menu.isVisible()
        _close_by_click_on(app, menu, btn)
        assert opened == [1]          # 다시 열리지 않음
        assert not menu.isVisible()
        _click(app, btn)              # 그다음 누름은 다시 연다
        assert opened == [1, 1]
    finally:
        menu.hide(); app.removeEventFilter(f); host.close()


def test_other_button_still_gets_click():
    app, f, host, btn, other, menu, opened = _setup()
    hits = []
    other.clicked.connect(lambda: hits.append(1))
    try:
        _click(app, btn)
        _close_by_click_on(app, menu, other)
        assert hits == [1]            # 다른 곳 누름은 그대로 전달
    finally:
        menu.hide(); app.removeEventFilter(f); host.close()


def test_closed_by_keyboard_does_not_swallow_next_click():
    app, f, host, btn, other, menu, opened = _setup()
    try:
        _click(app, btn)
        with mock.patch.object(QApplication, 'mouseButtons', return_value=Qt.NoButton):
            menu.hide()               # Esc 등으로 닫힘
            app.processEvents()
        _click(app, btn)
        assert opened == [1, 1]
    finally:
        menu.hide(); app.removeEventFilter(f); host.close()
