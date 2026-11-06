"""ui/theme.py — 디자인 토큰 (색상·글꼴의 단일 출처)과 라이트/다크 테마.

화면 코드에 색상 hex 값을 직접 쓰지 말고 여기 토큰을 쓴다.
  - 정적 문자열:   QColor(T.ACCENT)
  - 스타일시트:    f"border: 1px solid {T.ACCENT};"   (f-string 안에서 QSS 중괄호는 {{ }})
  - 한 번만 쓰는 색: tc('#888888')  — 라이트에선 그대로, 다크에선 자동 변환한 색

테마는 앱을 켤 때 한 번 정해진다 (설정: db.app_settings 'theme' = 'light' | 'dark').
화면들이 스타일을 만들 때 토큰을 읽으므로, 바꾼 테마는 다음 실행부터 적용된다.

다크 변환 규칙 (dark_color):
  - 무채색·파스텔·진한 글자색: 밝기를 뒤집고 색상(hue)은 유지 → 흰 바탕은 어두운 웜그레이,
    진한 갈색 글자는 밝은 금색, 파스텔 메모 색은 같은 색의 깊은 톤
  - rgba(0,0,0,a) 음영 → rgba(255,255,255,a) 밝게,  rgba(255,255,255,a) 유리막 → rgba(0,0,0,a/2)
  - 핵심 토큰은 _DARK에서 직접 지정
"""

import colorsys
import re


class T:
    # ── 메모(노란 스티커) 계열 ─────────────────────────────────────
    ACCENT               = '#d4b800'  # 입력칸 테두리, 포커스, 스크롤바, 주요 버튼
    ACCENT_HOVER         = '#b89e00'
    ACCENT_SOFT          = '#f1e7b3'  # 같은 노랑 계열의 옅은 단계 (집행률 막대 바탕)
    DIVIDER              = '#D4AF37'  # 제목/날짜/하단 구분선
    BORDER_SOFT          = '#c8b87c'  # 옅은 입력 테두리
    INPUT_BG             = '#fffacd'  # 메모 위 입력칸 배경
    SURFACE              = '#ffffff'  # 카드·입력칸 바탕
    TEXT_ON_NOTE         = '#5a4000'  # 노란 바탕 위 기본 글자
    INK                  = '#1a1410'  # 제목 등 가장 진한 글자
    TEXT_WARM            = '#5a4a3a'
    TEXT_WARM_MUTED      = '#7a6a5a'

    # ── 달력·상세 창 (따뜻한 회색 계열) ───────────────────────────
    SURFACE_WARM         = '#e8ddd5'  # 버튼 배경, 구분선
    SURFACE_WARM_HOVER   = '#ddd0c5'
    SURFACE_WARM_PRESSED = '#d0c0b5'
    BORDER_WARM          = '#f0e8e0'  # 칸 테두리
    BG_WARM              = '#faf6f0'  # 창 배경
    BG_WARM_ALT          = '#f5efe8'

    # ── 상태 색 ──────────────────────────────────────────────────
    HIGHLIGHT            = '#e85d3a'  # 오늘, 중요 업무
    HIGHLIGHT_BG         = '#fff3ee'
    DANGER               = '#e74c3c'  # 마감 임박 D-day, 삭제 hover
    SUNDAY               = '#cc0000'
    SATURDAY             = '#0055cc'
    SATURDAY_SOFT        = '#3a8ee8'

    # ── 중립 회색 스케일 (예산·공문 작성 창, Tailwind slate) ───────
    SLATE_50             = '#f8fafc'
    SLATE_100            = '#f1f5f9'
    SLATE_200            = '#e2e8f0'
    SLATE_300            = '#cbd5e1'
    SLATE_400            = '#94a3b8'
    SLATE_500            = '#64748b'
    SLATE_600            = '#475569'
    SLATE_700            = '#334155'
    SLATE_800            = '#1e293b'
    SLATE_900            = '#0f172a'
    GRAY_200             = '#e5e7eb'

    # ── 포인트 색 스케일 (예산·공문 작성 창, Tailwind indigo) ──────
    INDIGO_300           = '#a5b4fc'
    INDIGO_500           = '#6366f1'
    INDIGO_600           = '#4f46e5'
    INDIGO_700           = '#4338ca'
    INDIGO_900           = '#312e81'
    INDIGO_950           = '#1e1b4b'

    # ── 글꼴 ─────────────────────────────────────────────────────
    FONT_UI              = 'Pretendard'
    FONT_FALLBACK        = 'Malgun Gothic'


# 메모 창 배경 색상표 (메뉴의 색상 선택).
PALETTE = [
    '#FDD663', '#FEFFA7', '#CCFF90', '#A8F0E8',
    '#AECBFA', '#D7AEFB', '#FDCFE8', '#E6C9A8', '#AAAAAA', '#DDDDDD', '#FFFFFF',
]


# ── 다크 테마 ──────────────────────────────────────────────────

THEMES = ('light', 'dark')
THEME_KEY = 'theme'

# 핵심 토큰은 직접 고른 값 (나머지는 dark_color로 자동 변환)
_DARK = {
    'ACCENT': '#e0c42a', 'ACCENT_HOVER': '#f0d64a', 'ACCENT_SOFT': '#4a4118',
    'DIVIDER': '#6b5a1e', 'BORDER_SOFT': '#6e6440', 'INPUT_BG': '#2f2b1c', 'SURFACE': '#2a2621',
    'TEXT_ON_NOTE': '#f2dc8a', 'INK': '#f3ede4', 'TEXT_WARM': '#d9cfc2', 'TEXT_WARM_MUTED': '#a89c8c',
    'SURFACE_WARM': '#38322b', 'SURFACE_WARM_HOVER': '#433c33', 'SURFACE_WARM_PRESSED': '#4e463b',
    'BORDER_WARM': '#3a342c', 'BG_WARM': '#1c1915', 'BG_WARM_ALT': '#25211c',
    'HIGHLIGHT': '#ff7a55', 'HIGHLIGHT_BG': '#3a2219', 'DANGER': '#ff6b5b',
    'SUNDAY': '#ff6b6b', 'SATURDAY': '#6fa8ff', 'SATURDAY_SOFT': '#6fa8ff',
}

_NAMED = {'white': (255, 255, 255), 'black': (0, 0, 0)}
_RGBA = re.compile(r'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)')


def _parse(s: str):
    s = s.strip()
    low = s.lower()
    if low in _NAMED:
        return _NAMED[low] + (None,)
    if low.startswith('#'):
        h = low[1:]
        if len(h) == 3:
            h = ''.join(ch * 2 for ch in h)
        if len(h) == 6:
            return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), None
        return None
    m = _RGBA.fullmatch(low)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3)), m.group(4)
    return None


def _invert(r, g, b):
    """밝기 뒤집기 (색상 유지). 무채색은 살짝 따뜻한 회색으로."""
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if s < 0.15 or max(r, g, b) - min(r, g, b) < 12:
        l2 = 0.92 - l * 0.81            # 흰색 → #211d18 근처, #333 → #e8e2d9 근처
        h2, s2 = 0.09, 0.10
    elif 0.3 < l < 0.7:
        l2, h2, s2 = max(l, 1.0 - l), h, s          # 중간 밝기 강조색(파랑·초록·노랑)은 어둡게 하지 않음
    else:
        l2 = min(max(1.0 - l, 0.16), 0.84)
        h2, s2 = h, (s * 0.55 if l > 0.8 else s)   # 파스텔 바탕은 채도를 낮춘 깊은 톤
    r2, g2, b2 = colorsys.hls_to_rgb(h2, l2, s2)
    return round(r2 * 255), round(g2 * 255), round(b2 * 255)


def dark_color(s: str) -> str:
    """라이트 테마 색 문자열 → 다크 테마 색 문자열. 못 읽는 값은 그대로."""
    p = _parse(s) if isinstance(s, str) else None
    if p is None:
        return s
    r, g, b, a = p
    if a is not None:
        if (r, g, b) == (0, 0, 0):
            return f'rgba(255,255,255,{a})'
        if (r, g, b) == (255, 255, 255):
            return f'rgba(0,0,0,{float(a) / 2:g})'
        r, g, b = _invert(r, g, b)
        return f'rgba({r},{g},{b},{a})'
    return '#%02x%02x%02x' % _invert(r, g, b)


def dark_text_color(s: str) -> str:
    """글자색용 다크 변환. 진한 글자는 밝게(dark_color와 같음), 옅은 파스텔 글자는
    어두워져 사라지지 않도록 중간 밝기로 (예: 옅은 분홍 '일' 머리글 → 차분한 분홍)."""
    p = _parse(s) if isinstance(s, str) else None
    if p is None or p[3] is not None:
        return dark_color(s)
    r, g, b, _ = p
    h, l, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if l <= 0.7:
        return dark_color(s)
    if sat < 0.15 or max(r, g, b) - min(r, g, b) < 12:
        h2, l2, s2 = 0.09, 0.42, 0.06               # 아주 옅은 회색 글자 → 흐린 회색
    else:
        h2, l2, s2 = h, 0.62, min(sat, 0.6)         # 파스텔 글자 → 같은 색의 중간 톤
    r2, g2, b2 = colorsys.hls_to_rgb(h2, l2, s2)
    return '#%02x%02x%02x' % (round(r2 * 255), round(g2 * 255), round(b2 * 255))


def tct(color: str) -> str:
    """글자색. 라이트에선 그대로, 다크에선 dark_text_color."""
    return dark_text_color(color) if DARK else color


def note_color(hex_color: str) -> str:
    """메모 창 배경색(저장값)을 현재 테마의 표시색으로. 다크에선 같은 색상의 깊고 차분한 톤."""
    if not DARK:
        return hex_color
    p = _parse(hex_color) if isinstance(hex_color, str) else None
    if p is None:
        return hex_color
    h, l, s = colorsys.rgb_to_hls(p[0] / 255, p[1] / 255, p[2] / 255)
    if s < 0.15:                                    # 흰색·회색 메모 → 어두운 회색 단계
        l2, s2, h = 0.13 + (1.0 - l) * 0.25, 0.06, 0.09
    else:
        l2, s2 = 0.21, min(s, 0.9) * 0.45
    r, g, b = colorsys.hls_to_rgb(h, l2, s2)
    return '#%02x%02x%02x' % (round(r * 255), round(g * 255), round(b * 255))


def _load_mode() -> str:
    try:
        from db import get_setting
        mode = get_setting(THEME_KEY, 'light')
    except Exception:  # 첫 실행(설정 표 없음) 등 → 라이트
        mode = 'light'
    return mode if mode in THEMES else 'light'


MODE = _load_mode()
DARK = MODE == 'dark'


def tc(color: str) -> str:
    """한 번만 쓰는 색. 라이트에선 입력 그대로, 다크에선 dark_color로 변환."""
    return dark_color(color) if DARK else color


def qcolor(color: str, alpha: int = 255):
    """tc()를 거친 QColor (알파 지정). 숫자로 그리던 QColor(85, 85, 85, 80) 같은 곳에 쓴다."""
    from PyQt5.QtGui import QColor
    q = QColor(tc(color))
    q.setAlpha(alpha)
    return q


def _palette(dark: bool):
    from PyQt5.QtGui import QColor, QPalette
    src = T if dark else TL
    pal = QPalette()
    roles = {
        QPalette.Window: src.BG_WARM, QPalette.WindowText: src.INK,
        QPalette.Base: src.SURFACE, QPalette.AlternateBase: src.BG_WARM_ALT,
        QPalette.Text: src.INK, QPalette.Button: src.SURFACE_WARM, QPalette.ButtonText: src.INK,
        QPalette.ToolTipBase: src.SURFACE, QPalette.ToolTipText: src.INK,
        QPalette.Highlight: src.ACCENT, QPalette.HighlightedText: src.BG_WARM if dark else '#ffffff',
        QPalette.PlaceholderText: src.TEXT_WARM_MUTED, QPalette.Link: src.SATURDAY,
    }
    for role, col in roles.items():
        pal.setColor(role, QColor(col))
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        pal.setColor(QPalette.Disabled, role, QColor(src.TEXT_WARM_MUTED))
    return pal


_DEFAULT_PALETTE = None   # 다크로 바꾸기 전 Qt 기본 팔레트 (라이트 전용 창에 되돌려 줌)


def apply_app_palette(app):
    """다크 테마면 앱 기본 팔레트를 어둡게 (스타일시트가 없는 바탕·글자도 어둡게). 라이트면 아무것도 안 함."""
    global _DEFAULT_PALETTE
    if DARK:
        from PyQt5.QtGui import QPalette
        # 기본 팔레트는 '명시된 색이 없음' 상태라 그대로 복사해 setPalette하면 앱(다크) 팔레트를 따라간다.
        # 모든 그룹·역할을 setColor로 옮겨 적어 '명시된' 라이트 팔레트로 만든다.
        src, pal = app.palette(), QPalette()
        for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
            for role in range(QPalette.NColorRoles):
                if role == QPalette.NoRole:
                    continue
                pal.setColor(group, QPalette.ColorRole(role), src.color(group, QPalette.ColorRole(role)))
        _DEFAULT_PALETTE = pal
        app.setPalette(_palette(True))


def keep_light(widget):
    """아직 다크를 지원하지 않는 창(예산, 공문 작성)은 다크 테마에서도 원래(라이트) 팔레트로.

    화면을 다 만든 뒤 호출한다. 스타일시트가 있는 자식은 앱 팔레트를 따라가므로 자식마다 지정하고,
    이 창이 띄우는 대화상자에도 이어지도록 WA_WindowPropagation을 켠다.
    """
    if not DARK or _DEFAULT_PALETTE is None:
        return
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QWidget
    widget.setAttribute(Qt.WA_WindowPropagation)
    widget.setPalette(_DEFAULT_PALETTE)
    for ch in widget.findChildren(QWidget):
        ch.setPalette(_DEFAULT_PALETTE)


def set_theme(mode: str):
    """다음 실행부터 적용."""
    from db import set_setting
    set_setting(THEME_KEY, mode if mode in THEMES else 'light')


class TL:
    """항상 라이트 값 — 아직 다크 테마를 지원하지 않는 창(예산, 공문 작성)용."""


for _k, _v in list(vars(T).items()):
    if _k.isupper():
        setattr(TL, _k, _v)
        if DARK and isinstance(_v, str) and not _k.startswith('FONT'):
            setattr(T, _k, _DARK.get(_k, dark_color(_v)))
del _k, _v


# ── 메모 색에 맞춘 포인트 색 ─────────────────────────────────────
# 메모 창마다 바탕색에서 포인트 색(입력칸 테두리·선택한 줄·스크롤바 등)을 만든다.
# 업무 줄은 만들어질 때 지금 값(current_accent)을 받아 둔다 — 창이 여러 개여도 각자 자기 색.
_CURRENT_ACCENT = None


def accent_for_note(shown_hex: str) -> str:
    """메모 바탕(표시색)과 같은 색 계열의 진한 포인트 색. 무채색 메모는 회색."""
    from PyQt5.QtGui import QColor
    c = QColor(shown_hex)
    h, s, v, _ = c.getHsv()
    if h < 0 or s < 12:                       # 흰색·회색 메모
        return QColor.fromHsv(0, 0, min(220, v + 90) if DARK else max(100, v - 110)).name()
    if DARK:
        return QColor.fromHsv(h, min(255, s + 40), min(255, v + 120)).name()
    if 40 <= h <= 70:                         # 노랑 계열은 황금빛으로 (올리브색 방지)
        h = 48
    return QColor.fromHsv(h, min(255, max(s + 90, 150)), max(120, v - 75)).name()


def set_current_accent(hex_color):
    global _CURRENT_ACCENT
    _CURRENT_ACCENT = hex_color


def current_accent() -> str:
    return _CURRENT_ACCENT or T.ACCENT


def rgba(hex_color: str, alpha: float) -> str:
    """'#d4b800', 0.25 → 'rgba(212,184,0,0.25)' (QSS용)."""
    from PyQt5.QtGui import QColor
    c = QColor(hex_color)
    return f'rgba({c.red()},{c.green()},{c.blue()},{alpha})'
