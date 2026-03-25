import sys
import os
import socket
import ctypes
import ctypes.wintypes

from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtCore import Qt, QAbstractNativeEventFilter
from db import init_db, get_all_windows, create_window
from window import MemoWindow

_HOTKEY_CAPTURE = 1
_HOTKEY_RESTORE = 2
_MOD_CONTROL    = 0x0002
_MOD_SHIFT      = 0x0004
_VK_X           = 0x58
_VK_S           = 0x53
_WM_HOTKEY      = 0x0312


class _GlobalHotkeyFilter(QAbstractNativeEventFilter):
    # hotkeys: {id: (mod, vk, callback)}
    def __init__(self, hotkeys):
        super().__init__()
        self._hotkeys = hotkeys
        for hid, (mod, vk, _) in hotkeys.items():
            ctypes.windll.user32.RegisterHotKey(None, hid, mod, vk)

    def nativeEventFilter(self, eventType, message):
        if eventType == b'windows_generic_MSG':
            msg = ctypes.wintypes.MSG.from_address(int(message))
            if msg.message == _WM_HOTKEY and msg.wParam in self._hotkeys:
                self._hotkeys[msg.wParam][2]()
        return False, 0

    def unregister(self):
        for hid in self._hotkeys:
            ctypes.windll.user32.UnregisterHotKey(None, hid)

_SINGLE_INSTANCE_PORT = 47391  # 임의의 고정 포트

_open_windows = []
_last_active_window = None


def _make_tray_icon():
    pix = QPixmap(32, 32)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor('#f7c948'))
    p.setPen(QColor('#c8a000'))
    p.drawRoundedRect(2, 2, 28, 28, 4, 4)
    p.setPen(QColor('#7a6000'))
    p.drawLine(7, 10, 25, 10)
    p.drawLine(7, 16, 25, 16)
    p.drawLine(7, 22, 18, 22)
    p.end()
    return QIcon(pix)


def new_window(offset_from=None):
    x, y = 130, 130
    if offset_from:
        pos = offset_from.pos()
        x, y = pos.x() + 30, pos.y() + 30
    wid = create_window(x=x, y=y)
    _launch_window(wid, x, y, 320, 400, collapsed=False)


def _launch_window(wid, x, y, width, height, collapsed, color=''):
    win = MemoWindow(window_id=wid, on_new=new_window, open_windows=_open_windows)
    win.apply_state(x, y, width, height, collapsed, color)
    win.show()
    _open_windows.append(win)
    win.destroyed.connect(lambda _: _open_windows.remove(win) if win in _open_windows else None)


if __name__ == '__main__':
    # 중복 실행 방지: 이미 실행 중이면 조용히 종료
    _lock_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    _lock_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
    try:
        _lock_sock.bind(('127.0.0.1', _SINGLE_INSTANCE_PORT))
    except OSError:
        sys.exit(0)

    init_db()
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    def _on_focus_changed(old, new):
        global _last_active_window
        if new:
            top = new.window()
            if isinstance(top, MemoWindow):
                _last_active_window = top

    def _trigger_capture():
        if not _open_windows:
            return
        target = _last_active_window if _last_active_window in _open_windows else _open_windows[0]
        target._start_capture()

    def _show_or_restore():
        if _open_windows:
            for win in _open_windows:
                win.show()
                win.raise_()
                win.activateWindow()
        else:
            _restore_all_windows()

    app.focusChanged.connect(_on_focus_changed)
    _hotkey_filter = _GlobalHotkeyFilter({
        _HOTKEY_CAPTURE: (_MOD_CONTROL | _MOD_SHIFT, _VK_X, _trigger_capture),
        _HOTKEY_RESTORE: (_MOD_CONTROL | _MOD_SHIFT, _VK_S, _show_or_restore),
    })
    app.installNativeEventFilter(_hotkey_filter)
    app.aboutToQuit.connect(_hotkey_filter.unregister)

    _app_icon = _make_tray_icon()
    app.setWindowIcon(_app_icon)

    tray = QSystemTrayIcon(_app_icon, app)
    tray.setToolTip('서서니 메모')

    menu = QMenu()
    act_new  = QAction('새 메모')
    act_quit = QAction('종료')
    act_new.triggered.connect(lambda: new_window())
    act_quit.triggered.connect(app.quit)
    menu.addAction(act_new)
    menu.addSeparator()
    menu.addAction(act_quit)
    def _restore_all_windows():
        for w in get_all_windows():
            _launch_window(w['id'], w['x'], w['y'], w['width'], w['height'], bool(w['collapsed']), w.get('color', ''))
        if not _open_windows:
            new_window()

    def on_tray_activated(reason):
        if reason == QSystemTrayIcon.Trigger:  # 좌클릭
            _show_or_restore()

    tray.activated.connect(on_tray_activated)
    tray.setContextMenu(menu)
    tray.show()

    for w in get_all_windows():
        _launch_window(w['id'], w['x'], w['y'], w['width'], w['height'], bool(w['collapsed']), w.get('color', ''))

    sys.exit(app.exec_())
