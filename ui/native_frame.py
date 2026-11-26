"""ui/native_frame.py — Windows 제목표시줄을 숨기고 앱 머리줄이 대신하게 하는 도우미 (공문 작성 · 달력 창).

창은 win._is_caption_point(pos)로 「머리줄 빈 곳」을 알려 준다. 크기 조절을 막으려면
win._lock_frame = True (가장자리 · 머리줄 끌기 모두 꺼짐 — 달력 「바탕화면에 붙이기」).
"""
import sys

# ── Windows 제목표시줄 숨기기 도우미 (Windows에서만, 다른 곳에서는 기본 제목표시줄 그대로) ──
# Qt에는 제목표시줄 없는 창(FramelessWindowHint)이라고 알리고, Windows 창 스타일에는 테두리 · 제목 비트를
# 다시 켜 둔다. 그러면 Windows가 크기 조절 · 화면 가장자리에 붙이기 · 최소화 애니메이션 · 그림자를 그대로
# 해 주고, 창 전체가 앱 화면이 된다 (WM_NCCALCSIZE → 0). 머리줄 빈 곳은 HTCAPTION, 가장자리는 크기 조절.
CAPTION_HIDDEN = sys.platform == 'win32'
WM_NCCALCSIZE, WM_NCHITTEST = 0x0083, 0x0084
HT = {'left': 10, 'right': 11, 'top': 12, 'topleft': 13, 'topright': 14,
       'bottom': 15, 'bottomleft': 16, 'bottomright': 17}
HTCAPTION = 2


def setup_native_frame(win):
    """창 스타일에 테두리 · 제목 비트를 켜고, 그림자 · 둥근 모서리를 Windows에 맡긴다."""
    try:
        import ctypes
        from ctypes import wintypes
        hwnd = int(win.winId())
        user32 = ctypes.windll.user32
        GWL_STYLE = -16
        user32.GetWindowLongW.restype = ctypes.c_long
        style = user32.GetWindowLongW(hwnd, GWL_STYLE)
        style |= 0x00C00000 | 0x00040000 | 0x00020000 | 0x00010000   # CAPTION | THICKFRAME | MINIMIZEBOX | MAXIMIZEBOX
        user32.SetWindowLongW(hwnd, GWL_STYLE, style)

        class MARGINS(ctypes.Structure):
            _fields_ = [('l', ctypes.c_int), ('r', ctypes.c_int), ('t', ctypes.c_int), ('b', ctypes.c_int)]
        dwm = ctypes.windll.dwmapi
        dwm.DwmExtendFrameIntoClientArea(hwnd, ctypes.byref(MARGINS(1, 1, 1, 1)))   # 창 그림자
        pref = ctypes.c_int(2)                                                       # Windows 11 둥근 모서리
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(pref), ctypes.sizeof(pref))
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)   # FRAMECHANGED
    except Exception:
        pass


def caption_native_event(win, message):
    """WM_NCCALCSIZE · WM_NCHITTEST만 처리해 결과값을 돌려준다. 나머지는 None (Qt · Windows 기본 처리)."""
    import ctypes
    from ctypes import wintypes
    msg = wintypes.MSG.from_address(int(message))
    user32 = ctypes.windll.user32
    if msg.message == WM_NCCALCSIZE and msg.wParam:
        if user32.IsZoomed(msg.hWnd):          # 최대화하면 창이 화면 밖으로 테두리만큼 나가므로 그만큼 안쪽으로
            rect = wintypes.RECT.from_address(msg.lParam)
            pad = user32.GetSystemMetrics(92)                   # SM_CXPADDEDBORDER
            fx, fy = user32.GetSystemMetrics(32) + pad, user32.GetSystemMetrics(33) + pad   # SM_CXFRAME · SM_CYFRAME
            rect.left += fx
            rect.right -= fx
            rect.top += fy
            rect.bottom -= fy
        return 0                                                # 창 전체가 앱 화면 (제목표시줄 없음)
    if msg.message == WM_NCHITTEST:
        from PyQt5.QtGui import QCursor
        return hit_test(win, win.mapFromGlobal(QCursor.pos()), bool(user32.IsZoomed(msg.hWnd)))
    return None


def hit_test(win, pos, zoomed=False):
    """창 기준 pos → Windows 판정값 (가장자리 크기 조절 · 머리줄 빈 곳은 제목표시줄) 또는 None(앱 화면)."""
    if not zoomed:
        if getattr(win, '_lock_frame', False):    # 위치 · 크기 잠금
            return None
        e = 6
        w, h = win.width(), win.height()
        fixed_h = win.minimumHeight() == win.maximumHeight()
        fixed_w = win.minimumWidth() == win.maximumWidth()
        left, right = (pos.x() < e) and not fixed_w, (pos.x() >= w - e) and not fixed_w
        top, bottom = (pos.y() < e) and not fixed_h, (pos.y() >= h - e) and not fixed_h
        key = ('top' if top else 'bottom' if bottom else '') + ('left' if left else 'right' if right else '')
        if key:
            return HT[key]
    if getattr(win, '_lock_frame', False):
        return None
    return HTCAPTION if win._is_caption_point(pos) else None
