"""캡처할 때 브라우저 잠깐 확대 (Windows 전용).

공문을 100%로 띄운 채 캡처하면 글자가 작아 글자 인식이 자주 틀린다. 그래서 캡처 직전에
맨 앞 창이 브라우저(엣지 · 크롬 등)면 Ctrl+'+'를 몇 번 보내 화면을 키우고, 캡처가 끝나면(취소해도)
Ctrl+'-'를 같은 횟수만큼 보내 되돌린다.

- zoom_then(callback): 키울 수 있으면 키운 뒤 다시 그려질 시간을 두고 callback, 아니면 바로 callback.
- restore(): 키워 둔 게 있으면 되돌림 (여러 번 불러도 한 번만).
- 설정: db 'capture_zoom' ('0'이면 끔, 기본 켬). 브라우저가 아니거나 Windows가 아니면 아무것도 안 함.
"""
import sys
import time

SETTING_KEY = 'capture_zoom'
STEPS = 3                      # 크롬 · 엣지 확대 단계: 100 → 110 → 125 → 150%
RENDER_WAIT_MS = 450           # 키운 화면이 다시 그려질 시간
BROWSERS = ('msedge.exe', 'chrome.exe', 'whale.exe', 'firefox.exe', 'brave.exe')

_zoomed = None                 # (hwnd, 보낸 횟수) — 되돌릴 것이 있을 때만

_VK_CONTROL, _VK_SHIFT, _VK_MENU, _VK_X = 0x11, 0x10, 0x12, 0x58
_VK_OEM_PLUS, _VK_OEM_MINUS = 0xBB, 0xBD
_KEYEVENTF_KEYUP = 0x0002


def enabled() -> bool:
    try:
        from db import get_setting
        return get_setting(SETTING_KEY, '1') != '0'
    except Exception:
        return True


def set_enabled(on: bool):
    from db import set_setting
    set_setting(SETTING_KEY, '1' if on else '0')


# ── Windows 호출 (테스트에서는 이 함수들을 바꿔 끼움) ─────────────────
def _process_name(hwnd) -> str:
    import ctypes
    from ctypes import wintypes
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    if not pid.value:
        return ''
    kernel32.OpenProcess.restype = wintypes.HANDLE
    h = kernel32.OpenProcess(0x1000, False, pid.value)      # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return ''
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(1024)
        if not kernel32.QueryFullProcessImageNameW(wintypes.HANDLE(h), 0, buf, ctypes.byref(size)):
            return ''
        return buf.value.replace('/', '\\').rsplit('\\', 1)[-1].lower()
    finally:
        kernel32.CloseHandle(wintypes.HANDLE(h))


def _foreground_window() -> int:
    import ctypes
    ctypes.windll.user32.GetForegroundWindow.restype = ctypes.c_void_p
    return ctypes.windll.user32.GetForegroundWindow() or 0


def _top_other_window() -> int:
    """우리 앱 창을 뺀, 화면 맨 위(Z 순서)의 보이는 일반 창."""
    import ctypes
    import os
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    me = os.getpid()
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd):
            return True
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == me:
            return True
        if user32.GetWindowLongW(hwnd, -20) & 0x00000080:    # WS_EX_TOOLWINDOW
            return True
        if user32.GetWindowTextLengthW(hwnd) == 0:
            return True
        rc = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rc))
        if rc.right - rc.left < 200 or rc.bottom - rc.top < 150:
            return True
        found.append(hwnd)
        return False                                         # 첫 창에서 멈춤

    user32.EnumWindows(cb, 0)
    return int(found[0]) if found else 0


def _set_foreground(hwnd) -> bool:
    import ctypes
    from ctypes import wintypes
    return bool(ctypes.windll.user32.SetForegroundWindow(wintypes.HWND(hwnd)))


def _key_down(vk) -> bool:
    import ctypes
    return bool(ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000)


def _send_ctrl(vk, times):
    import ctypes
    kb = ctypes.windll.user32.keybd_event
    kb(_VK_CONTROL, 0, 0, 0)
    for _ in range(times):
        kb(vk, 0, 0, 0)
        kb(vk, 0, _KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)
    kb(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)
    time.sleep(0.06)                                         # 브라우저가 키를 받아 갈 틈


def _wait_keys_released(timeout=0.8):
    """단축키(Ctrl+Shift+X)를 아직 누르고 있으면 뗄 때까지 잠깐 기다린다 — 눌린 채로 보내면 다른 단축키가 됨."""
    end = time.time() + timeout
    while time.time() < end:
        if not any(_key_down(vk) for vk in (_VK_SHIFT, _VK_MENU, _VK_X, _VK_CONTROL)):
            return True
        time.sleep(0.02)
    return False


# ── 공개 함수 ────────────────────────────────────────────────────
def find_browser() -> int:
    """지금 화면 맨 앞에 있는 브라우저 창 (없으면 0)."""
    if sys.platform != 'win32':
        return 0
    try:
        for hwnd in (_foreground_window(), _top_other_window()):
            if hwnd and _process_name(hwnd) in BROWSERS:
                return hwnd
    except Exception:
        pass
    return 0


def zoom_in() -> bool:
    """맨 앞 브라우저를 STEPS단계 키운다. 키웠으면 True."""
    global _zoomed
    if _zoomed is not None or not enabled():
        return False
    hwnd = find_browser()
    if not hwnd:
        return False
    try:
        if not _wait_keys_released():
            return False
        _set_foreground(hwnd)
        time.sleep(0.05)
        if _foreground_window() != hwnd:
            return False                                     # 앞으로 못 가져왔으면 엉뚱한 창에 키를 보내지 않음
        _send_ctrl(_VK_OEM_PLUS, STEPS)
        _zoomed = (hwnd, STEPS)
        return True
    except Exception:
        return False


def restore():
    """zoom_in으로 키운 것을 되돌린다 (키운 게 없으면 아무것도 안 함)."""
    global _zoomed
    if _zoomed is None:
        return False
    hwnd, steps = _zoomed
    _zoomed = None
    try:
        _set_foreground(hwnd)
        time.sleep(0.05)
        if _foreground_window() != hwnd:
            return False
        _send_ctrl(_VK_OEM_MINUS, steps)
        return True
    except Exception:
        return False


def zoom_then(callback):
    """브라우저를 키울 수 있으면 키우고 다시 그려진 뒤 callback, 아니면 바로 callback."""
    if not zoom_in():
        callback()
        return
    from PyQt5.QtCore import QTimer

    def run():
        try:
            callback()
        except Exception:          # 그 사이 창이 닫혔거나 캡처 화면을 못 띄웠으면 키운 것만 되돌림
            restore()
            raise
    QTimer.singleShot(RENDER_WAIT_MS, run)
