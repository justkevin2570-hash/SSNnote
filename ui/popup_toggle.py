"""ui/popup_toggle.py — 메뉴 · 달력 같은 팝업을 '한 번 누르면 열리고, 다시 누르면 닫히게'.

Qt는 팝업이 열린 채 바깥을 누르면 팝업을 닫고, 그 누름을 아래 위젯에 한 번 더 보낸다(replay).
그래서 메뉴를 연 단추를 다시 누르면 '닫힘 → 단추가 눌림 → 다시 열림'이 되어 닫히지 않는다.

앱 전체 이벤트 필터 하나로 고친다:
  · 팝업이 처음 뜰 때, 바로 전에 눌린 위젯(= 연 단추)을 기억한다.
  · 마우스를 누른 채로 팝업이 닫혔고, 그 자리가 연 단추 위라면
    곧바로 그 단추로 다시 오는 누름(과 떼기)을 삼킨다 → 다시 열리지 않는다.
다른 곳을 누르면 지금처럼 팝업만 닫히고 그 누름은 그대로 전달된다.
"""
import time

from PyQt5.QtCore import QObject, QEvent, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QWidget

SWALLOW_SEC = 0.5

_FILTER = None


def _related(a, b) -> bool:
    """b가 a 자신이거나 a 안의 자식인가."""
    if not isinstance(a, QWidget) or not isinstance(b, QWidget):
        return False   # 창(QWindow) 같은 위젯이 아닌 것은 비교하지 않음
    try:
        return a is b or a.isAncestorOf(b)   # b가 a(연 단추) 안에 있을 때만
    except Exception:   # 이미 지워진 위젯 등
        return False


class PopupToggleFilter(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._last_pressed = None   # 마지막으로 마우스가 눌린 위젯
        self._opener = None         # 지금 열린 팝업을 연 위젯
        self._swallow = None        # (위젯, 시각) — 이 위젯으로 오는 다음 누름을 삼킨다
        self._eat_release = None

    @staticmethod
    def _is_popup(obj) -> bool:
        return isinstance(obj, QWidget) and obj.isWindow() and obj.windowType() == Qt.Popup

    def eventFilter(self, obj, event):
        # 여기서 오류가 나면 PyQt가 앱을 통째로 끈다 → 어떤 경우에도 예외를 밖으로 내보내지 않는다
        try:
            return self._filter(obj, event)
        except Exception:
            self._swallow = self._eat_release = None
            return False

    def _filter(self, obj, event):
        t = event.type()
        if t in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick):
            if not isinstance(obj, QWidget):
                return False
            sw = self._swallow
            if sw is not None:
                self._swallow = None
                w, at = sw
                if time.monotonic() - at < SWALLOW_SEC and _related(w, obj):
                    self._eat_release = (obj, time.monotonic())
                    return True
            self._eat_release = None   # 새로 누름이 왔으면 남은 '떼기 삼키기'는 버림 (이번 누름의 떼기는 그대로)
            if not self._is_popup(obj.window()):
                self._last_pressed = obj
        elif t == QEvent.MouseButtonRelease:
            if self._eat_release is not None and isinstance(obj, QWidget):
                w, at = self._eat_release
                self._eat_release = None   # 남아 있지 않게 바로 지움 (떼기가 먼저 와 버린 경우 대비)
                if time.monotonic() - at < SWALLOW_SEC and _related(w, obj):
                    return True
        elif t == QEvent.Show:
            if self._is_popup(obj):
                active = QApplication.activePopupWidget()
                if active is None or active is obj:   # 하위 메뉴가 아니라 처음 뜬 팝업
                    self._opener = self._last_pressed
        elif t == QEvent.Hide:
            if self._is_popup(obj) and self._opener is not None:
                opener = self._opener
                try:
                    alive = opener.isVisible()
                except RuntimeError:
                    alive = False
                if alive and QApplication.mouseButtons() != Qt.NoButton:
                    under = QApplication.widgetAt(QCursor.pos())
                    if _related(opener, under):
                        self._swallow = (opener, time.monotonic())
        return False


def install(app=None):
    """앱에 한 번만 설치."""
    global _FILTER
    app = app or QApplication.instance()
    if app is None or _FILTER is not None:
        return _FILTER
    _FILTER = PopupToggleFilter(app)
    app.installEventFilter(_FILTER)
    return _FILTER
