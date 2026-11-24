"""캡처할 때 브라우저 잠깐 확대 (Windows 전용).

공문을 100%로 띄운 채 캡처하면 글자가 작아 글자 인식이 자주 틀린다. 그래서 캡처 직전에
맨 앞 창이 브라우저(엣지 · 크롬 등)면 그 창 가운데에 Ctrl+휠을 몇 칸 굴려 화면을 키우고, 캡처가 끝나면(취소해도)
반대로 같은 칸만큼 굴려 되돌린다 (v3.59의 Ctrl+'+' 키 방식은 실제 PC에서 안 먹어서 휠로 바꿈).

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

_zoomed = None                 # (hwnd, 굴린 칸 수) — 되돌릴 것이 있을 때만
last_status = ''               # 이번 캡처에서 키웠는지 · 못 키운 까닭 (캡처 화면 위쪽에 보여 줌)

_VK_CONTROL, _VK_SHIFT, _VK_MENU, _VK_X = 0x11, 0x10, 0x12, 0x58
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


def _window_center(hwnd):
    import ctypes
    from ctypes import wintypes
    rc = wintypes.RECT()
    if not ctypes.windll.user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(rc)):
        return None
    return (rc.left + rc.right) // 2, (rc.top + rc.bottom) // 2


def _root_at(x, y) -> int:
    """그 점에 보이는 창의 맨 위 부모 창."""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    user32.WindowFromPoint.restype = ctypes.c_void_p
    user32.GetAncestor.restype = ctypes.c_void_p
    h = user32.WindowFromPoint(wintypes.POINT(x, y))
    return (user32.GetAncestor(wintypes.HWND(h), 2) or 0) if h else 0     # GA_ROOT


def _ctrl_wheel(x, y, notches):
    """마우스를 (x, y)에 잠깐 두고 Ctrl+휠을 notches칸 (양수: 확대) 굴린 뒤 마우스를 제자리로."""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    old = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(old))
    user32.SetCursorPos(x, y)
    time.sleep(0.03)
    user32.keybd_event(_VK_CONTROL, 0, 0, 0)
    try:
        step = 120 if notches > 0 else -120
        for _ in range(abs(notches)):
            user32.mouse_event(0x0800, 0, 0, ctypes.c_ulong(step & 0xFFFFFFFF), 0)   # MOUSEEVENTF_WHEEL
            time.sleep(0.05)
    finally:
        user32.keybd_event(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)
    time.sleep(0.08)                                         # 브라우저가 휠을 받아 갈 틈
    user32.SetCursorPos(old.x, old.y)


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
    """지금 화면 맨 앞에 있는 브라우저 창 (없으면 0). 못 찾은 까닭은 last_status에."""
    global last_status
    if sys.platform != 'win32':
        return 0
    seen = ''
    try:
        for hwnd in (_foreground_window(), _top_other_window()):
            if not hwnd:
                continue
            name = _process_name(hwnd)
            if name in BROWSERS:
                return hwnd
            seen = seen or name
    except Exception as e:
        last_status = f'확대 안 함: 창 확인 오류 ({type(e).__name__})'
        return 0
    last_status = f'확대 안 함: 맨 앞 창이 브라우저가 아니에요 ({seen or "창 없음"})'
    return 0


def _wheel_at(hwnd, notches) -> bool:
    global last_status
    pt = _window_center(hwnd)
    if pt is None:
        last_status = '확대 안 함: 브라우저 창 위치를 못 읽었어요'
        return False
    try:
        _set_foreground(hwnd)                                # 되면 좋고, 안 돼도 휠은 마우스 아래 창으로 감
    except Exception:
        pass
    time.sleep(0.05)
    if _root_at(*pt) != hwnd:
        last_status = '확대 안 함: 브라우저가 다른 창에 가려져 있어요'
        return False
    _ctrl_wheel(pt[0], pt[1], notches)
    return True


def zoom_in() -> bool:
    """맨 앞 브라우저에 Ctrl+휠을 STEPS칸 굴려 키운다. 키웠으면 True."""
    global _zoomed, last_status
    last_status = ''
    if _zoomed is not None or not enabled():
        return False
    hwnd = find_browser()
    if not hwnd:
        return False
    try:
        if not _wait_keys_released():
            last_status = '확대 안 함: 단축키를 아직 누르고 있어요 (누르고 바로 떼 주세요)'
            return False
        if not _wheel_at(hwnd, STEPS):
            return False
        _zoomed = (hwnd, STEPS)
        last_status = '🔍 화면을 잠깐 키웠어요 — 끝나면 원래 크기로 돌아가요'
        return True
    except Exception as e:
        last_status = f'확대 안 함: 오류 ({type(e).__name__})'
        return False


def restore():
    """zoom_in으로 키운 것을 되돌린다 (키운 게 없으면 아무것도 안 함). 캡처 화면을 숨긴 뒤에 부를 것."""
    global _zoomed
    if _zoomed is None:
        return False
    hwnd, steps = _zoomed
    _zoomed = None
    try:
        return _wheel_at(hwnd, -steps)
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
