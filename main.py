import sys
import os
import json
import socket
import ctypes
import ctypes.wintypes
import threading

_SETTINGS_PATH = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'settings.json')

def _load_settings():
    try:
        with open(_SETTINGS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _save_settings(data):
    try:
        os.makedirs(os.path.dirname(_SETTINGS_PATH), exist_ok=True)
        with open(_SETTINGS_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMessageBox
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtCore import Qt, QAbstractNativeEventFilter, QTimer, qInstallMessageHandler, QtMsgType


def _qt_message_handler(msg_type, context, message):
    if 'Unknown property cursor' in message:
        return
    if msg_type in (QtMsgType.QtDebugMsg, QtMsgType.QtInfoMsg):
        return
    print(message, file=sys.stderr)

qInstallMessageHandler(_qt_message_handler)
from db import init_db, get_all_windows, get_tasks, create_window
from window import MemoWindow
import auth
import sync
import updater

_WELCOME_FLAG = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'welcome_shown.txt')

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
        self._enabled = True
        for hid, (mod, vk, _) in hotkeys.items():
            ctypes.windll.user32.RegisterHotKey(None, hid, mod, vk)

    def nativeEventFilter(self, eventType, message):
        if eventType == b'windows_generic_MSG':
            msg = ctypes.wintypes.MSG.from_address(int(message))
            if msg.message == _WM_HOTKEY and msg.wParam in self._hotkeys and self._enabled:
                self._hotkeys[msg.wParam][2]()
        return False, 0

    def unregister(self):
        for hid in self._hotkeys:
            ctypes.windll.user32.UnregisterHotKey(None, hid)

    def set_enabled(self, enabled: bool):
        if enabled == self._enabled:
            return
        self._enabled = enabled
        if enabled:
            for hid, (mod, vk, _) in self._hotkeys.items():
                ctypes.windll.user32.RegisterHotKey(None, hid, mod, vk)
        else:
            for hid in self._hotkeys:
                ctypes.windll.user32.UnregisterHotKey(None, hid)

_SINGLE_INSTANCE_PORT = 47391  # 임의의 고정 포트

_open_windows = []
_last_active_window = None


def _make_tray_icon():
    from PyQt5.QtGui import QPen
    pix = QPixmap(32, 32)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    # 노트 본체 (노란 배경)
    p.setBrush(QColor('#f7c948'))
    p.setPen(QColor('#c8a000'))
    p.drawRoundedRect(2, 2, 28, 28, 4, 4)
    # 노트 선 (어두운 색)
    pen = QPen(QColor('#7a6000'))
    pen.setWidth(2)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    p.drawLine(7, 10, 25, 10)
    p.drawLine(7, 16, 25, 16)
    p.drawLine(7, 22, 18, 22)
    p.end()
    return QIcon(pix)


def new_window(offset_from=None, on_toggle_hotkey=None, on_shortcut_change=None, get_shortcut_enabled=None):
    x, y = 130, 130
    if offset_from:
        pos = offset_from.pos()
        x, y = pos.x() + 30, pos.y() + 30
    wid = create_window(x=x, y=y)
    _launch_window(wid, x, y, 320, 400, collapsed=False, on_toggle_hotkey=on_toggle_hotkey,
                   on_shortcut_change=on_shortcut_change, get_shortcut_enabled=get_shortcut_enabled)


_alarm_callbacks = {}  # 정의 전 참조 문제 우회용

def _launch_window(wid, x, y, width, height, collapsed, color='', scale=1.0, on_toggle_hotkey=None,
                   on_shortcut_change=None, get_shortcut_enabled=None):
    win = MemoWindow(
        window_id=wid,
        on_new=new_window,
        open_windows=_open_windows,
        on_toggle_hotkey=on_toggle_hotkey,
        on_alarm_interval_change=lambda m: _alarm_callbacks.get('set', lambda m: None)(m),
        get_alarm_interval=lambda: _alarm_callbacks.get('get', lambda: 180)(),
        on_timed_alarm_change=lambda v: _alarm_callbacks.get('set_timed', lambda v: None)(v),
        get_timed_alarm_enabled=lambda: _alarm_callbacks.get('get_timed', lambda: True)(),
        on_shortcut_change=on_shortcut_change,
        get_shortcut_enabled=get_shortcut_enabled,
    )
    win.apply_state(x, y, width, height, collapsed, color, scale)
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
    import autostart
    autostart.refresh_if_enabled()
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    QApplication.setAttribute(Qt.AA_ShareOpenGLContexts, True)
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    def _on_focus_changed(old, new):
        global _last_active_window
        if new:
            top = new.window()
            if isinstance(top, MemoWindow):
                _last_active_window = top

    def _trigger_capture():
        # 열린 공문 작성 창 찾기
        doc_editor = None
        for w in QApplication.topLevelWidgets():
            if type(w).__name__ == 'DocumentEditorWindow' and w.isVisible():
                doc_editor = w
                break

        if doc_editor is None:
            # 공문 작성 창이 없으면 열기
            if not _open_windows:
                return
            target = _last_active_window if _last_active_window in _open_windows else _open_windows[0]
            target._open_document_editor()
            doc_editor = getattr(target, '_doc_editor', None)
            if doc_editor is None:
                return
            # 창이 뜨고 나서 캡처 시작
            QTimer.singleShot(300, doc_editor._start_capture)
        else:
            doc_editor._start_capture()

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

    # ── 단축키 켬/끔 상태 관리 ──────────────────────────────────────
    _sc_settings = _load_settings()
    _shortcut_enabled = [_sc_settings.get('shortcut_enabled', True)]

    def _get_shortcut_enabled():
        return _shortcut_enabled[0]

    def _set_shortcut_enabled(enabled):
        _shortcut_enabled[0] = enabled
        _hotkey_filter.set_enabled(enabled)
        for win in _open_windows:
            win._set_local_shortcuts_enabled(enabled)
        s = _load_settings()
        s['shortcut_enabled'] = enabled
        _save_settings(s)

    # 저장된 상태가 끔이었다면 초기 적용
    if not _shortcut_enabled[0]:
        _hotkey_filter.set_enabled(False)

    _app_icon = _make_tray_icon()
    app.setWindowIcon(_app_icon)

    tray = QSystemTrayIcon(_app_icon, app)
    tray.setToolTip('서서니 노트')

    menu = QMenu()
    act_quit = QAction('종료')
    act_quit.triggered.connect(app.quit)
    menu.addAction(act_quit)
    def _restore_all_windows():
        for w in get_all_windows():
            _launch_window(w['id'], w['x'], w['y'], w['width'], w['height'], bool(w['collapsed']), w.get('color', ''), w.get('scale', 1.0),
                           on_toggle_hotkey=_hotkey_filter.set_enabled,
                           on_shortcut_change=_set_shortcut_enabled,
                           get_shortcut_enabled=_get_shortcut_enabled)
        if not _open_windows:
            new_window(on_toggle_hotkey=_hotkey_filter.set_enabled,
                       on_shortcut_change=_set_shortcut_enabled,
                       get_shortcut_enabled=_get_shortcut_enabled)

    def on_tray_activated(reason):
        if reason == QSystemTrayIcon.Trigger:  # 좌클릭
            _show_or_restore()

    tray.activated.connect(on_tray_activated)
    tray.setContextMenu(menu)
    tray.show()

    for w in get_all_windows():
        _launch_window(w['id'], w['x'], w['y'], w['width'], w['height'], bool(w['collapsed']), w.get('color', ''), w.get('scale', 1.0),
                       on_toggle_hotkey=_hotkey_filter.set_enabled,
                       on_shortcut_change=_set_shortcut_enabled,
                       get_shortcut_enabled=_get_shortcut_enabled)

    def _cloud_init():
        """백그라운드: 익명 인증 후 전체 데이터 Supabase 동기화."""
        session = auth.init()
        if not session:
            return
        windows = get_all_windows()
        tasks_by_window = {w['id']: get_tasks(w['id']) for w in windows}
        sync.push_all(windows, tasks_by_window)

    threading.Thread(target=_cloud_init, daemon=True).start()

    def _show_welcome_if_first_run():
        if os.path.exists(_WELCOME_FLAG):
            return
        msg = QMessageBox()
        msg.setWindowTitle('서서니 노트')
        msg.setText(
            '반갑습니다~ 노트 앱이 한 분에게라도\n'
            '도움이 되었으면 하는 마음입니다.\n\n'
            '꼭 도움말 읽어보시고, 단축키 잘 활용하세요!'
        )
        msg.setStandardButtons(QMessageBox.Ok)
        btn = msg.button(QMessageBox.Ok)
        if btn:
            btn.setText('확인')
        msg.exec_()
        os.makedirs(os.path.dirname(_WELCOME_FLAG), exist_ok=True)
        with open(_WELCOME_FLAG, 'w', encoding='utf-8') as f:
            f.write('shown')

    QTimer.singleShot(300, _show_welcome_if_first_run)

    # ── 마감 임박 알람 ────────────────────────────────────────────
    _urgent_toast = None  # GC 방지용 참조
    _shown_timed_alerts = set()  # (task_id, hours) 중복 방지

    def _check_urgent_tasks():
        global _urgent_toast
        from window import calc_dday, UrgentToast
        urgent = []
        seen = set()
        for win_data in get_all_windows():
            for task in get_tasks(win_data['id']):
                if task.get('strikethrough') or not task.get('deadline'):
                    continue
                dday = calc_dday(task['deadline'])
                if dday.startswith('D+') or dday == '날짜오류':
                    continue
                is_urgent = (
                    dday == 'D-day' or '시간' in dday or
                    (dday.startswith('D-') and int(dday[2:]) <= 5)
                )
                if is_urgent and task['id'] not in seen:
                    urgent.append(task)
                    seen.add(task['id'])
        if urgent:
            _urgent_toast = UrgentToast(urgent)
            _urgent_toast.show()

    _settings_cache = _load_settings()
    _timed_alarm_enabled = [_settings_cache.get('timed_alarm_enabled', True)]

    def _check_timed_tasks():
        global _urgent_toast
        if not _timed_alarm_enabled[0]:
            return
        from datetime import datetime
        from window import UrgentToast
        urgent = []
        now = datetime.now()

        for win_data in get_all_windows():
            for task in get_tasks(win_data['id']):
                if task.get('strikethrough') or not task.get('deadline'):
                    continue
                dl = task['deadline']
                if len(dl) <= 10:
                    continue
                try:
                    deadline_dt = datetime.strptime(dl, '%Y-%m-%d %H:%M')
                except ValueError:
                    continue
                remaining_min = (deadline_dt - now).total_seconds() / 60
                alerts = [(3, '3시간 후 마감'), (2, '2시간 후 마감'), (1, '1시간 후 마감'), (0.5, '30분 후 마감')]
                for threshold, label in alerts:
                    key = (task['id'], threshold)
                    target_min = threshold * 60
                    if abs(remaining_min - target_min) <= 5 and key not in _shown_timed_alerts:
                        _shown_timed_alerts.add(key)
                        t = dict(task)
                        t['notice'] = label
                        urgent.append(t)
                        break

        # 지난 마감 항목 정리
        to_remove = set()
        for key in _shown_timed_alerts:
            task_id, _ = key
            found = any(
                t['id'] == task_id
                for w in get_all_windows()
                for t in get_tasks(w['id'])
            )
            if not found:
                to_remove.add(key)
        _shown_timed_alerts.difference_update(to_remove)

        if urgent:
            _urgent_toast = UrgentToast(urgent)
            _urgent_toast.show()

    _alarm_interval = [_load_settings().get('alarm_interval_minutes', 180)]

    def _get_alarm_interval():
        return _alarm_interval[0]

    def _set_alarm_interval(minutes):
        _alarm_interval[0] = minutes
        if minutes == 0:
            _alarm_timer.stop()
        else:
            _alarm_timer.setInterval(minutes * 60 * 1000)
            if not _alarm_timer.isActive():
                _alarm_timer.start()
        s = _load_settings()
        s['alarm_interval_minutes'] = minutes
        _save_settings(s)

    def _get_timed_alarm_enabled():
        return _timed_alarm_enabled[0]

    def _set_timed_alarm_enabled(enabled):
        _timed_alarm_enabled[0] = enabled
        s = _load_settings()
        s['timed_alarm_enabled'] = enabled
        _save_settings(s)

    _alarm_callbacks['get'] = _get_alarm_interval
    _alarm_callbacks['set'] = _set_alarm_interval
    _alarm_callbacks['get_timed'] = _get_timed_alarm_enabled
    _alarm_callbacks['set_timed'] = _set_timed_alarm_enabled

    _alarm_timer = QTimer()
    _alarm_timer.timeout.connect(_check_urgent_tasks)
    if _alarm_interval[0] != 0:
        _alarm_timer.start(_alarm_interval[0] * 60 * 1000)
        QTimer.singleShot(10000, _check_urgent_tasks)    # 앱 시작 10초 후 첫 체크

    _timed_alarm_timer = QTimer()
    _timed_alarm_timer.timeout.connect(_check_timed_tasks)
    _timed_alarm_timer.start(60 * 1000)              # 1분마다

    _update_notifier = updater.UpdateNotifier()
    _banner_notifier = updater.StatusBannerNotifier()

    def _on_update_status(has_update: bool):
        msg = '새로운 업데이트가 준비되었습니다.' if has_update else '최신 버전입니다.'
        for win in _open_windows:
            win.title_bar.status_label.setText(msg)
        QTimer.singleShot(5000, _restore_titles)

    def _restore_titles():
        for win in _open_windows:
            win.title_bar.status_label.setText('')

    _banner_notifier.status_signal.connect(_on_update_status, Qt.QueuedConnection)

    threading.Thread(
        target=updater.check_for_update_on_startup,
        args=(_update_notifier, _banner_notifier),
        daemon=True,
    ).start()

    sys.exit(app.exec_())
