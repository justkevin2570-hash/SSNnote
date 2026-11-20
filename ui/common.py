"""ui/common.py — 공통 상수·경로·아이콘 폰트(Fluent/Material/Pretendard) 로더."""
import os
import sys
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor, QPainter, QIcon, QPixmap, QFontDatabase, QFontMetricsF

def _base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    # 이 파일은 ui/ 안에 있으므로 프로젝트 루트(assets/가 있는 곳)는 한 단계 위
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TITLE_BAR_HEIGHT = 38
TITLE_COLOR      = '#f7c948'
SNAP_THRESHOLD   = 12  # 완전히 붙는 거리(px)
SNAP_ZONE        = 30  # 당기기 시작하는 거리(px)
ALIGN_THRESHOLD  = 12  # 정렬 스냅 확정 거리(px)
ALIGN_ZONE       = 30  # 정렬 스냅 인력 시작 거리(px)
SNAP_RESIST      = 18  # 정렬 스냅 탈출 거리(px)
BREAKAWAY_DIST   = 25  # 인접 스냅 탈출 거리(px) - 유클리드 거리 기준

# ── Fluent System Icons 폰트 로드 ────────────────────────────────
_FI_FONT_ID        = -1
_FI_FILLED_FONT_ID = -1
_MAT_FONT_ID       = -1
_PR_FONT_ID        = -1
_PRETENDARD_LOADED = False

def _load_fluent_icons():
    global _FI_FONT_ID, _FI_FILLED_FONT_ID, _MAT_FONT_ID
    base = _base_path()
    if _FI_FONT_ID == -1:
        _FI_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'font', 'FluentSystemIcons-Regular.ttf'))
    if _FI_FILLED_FONT_ID == -1:
        _FI_FILLED_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'font', 'FluentSystemIcons-Filled.ttf'))
    if _MAT_FONT_ID == -1:
        _MAT_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'font', 'MaterialIcons-Regular.ttf'))

def mi_font(size=14):
    """Fluent System Icons Regular QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_FI_FONT_ID)
    return QFont(families[0] if families else 'FluentSystemIcons-Regular', size)

def mi_font_filled(size=14):
    """Fluent System Icons Filled QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_FI_FILLED_FONT_ID)
    return QFont(families[0] if families else 'FluentSystemIcons-Filled', size)

def mat_font(size=14):
    """Material Icons QFont 반환."""
    _load_fluent_icons()
    families = QFontDatabase.applicationFontFamilies(_MAT_FONT_ID)
    return QFont(families[0] if families else 'Material Icons', size)

def _load_pretendard():
    """assets/font 안의 Pretendard 정적 폰트(여러 굵기)를 앱 폰트로 등록한다.

    QSS의 font-family: 'Pretendard' 가 이 등록된 패밀리를 쓰고,
    굵은 글씨는 합성 볼드 대신 실제 굵기 파일로 그려진다.
    Regular 파일의 id는 pr_font()용으로 _PR_FONT_ID에 남긴다.
    """
    global _PR_FONT_ID, _PRETENDARD_LOADED
    if _PRETENDARD_LOADED:
        return
    _PRETENDARD_LOADED = True
    for dirpath, _dirs, files in os.walk(os.path.join(_base_path(), 'assets', 'font')):
        for fn in sorted(files):
            low = fn.lower()
            if low.startswith('pretendard') and low.endswith(('.ttf', '.otf')):
                fid = QFontDatabase.addApplicationFont(os.path.join(dirpath, fn))
                if fid != -1 and low == 'pretendard-regular.ttf':
                    _PR_FONT_ID = fid

def pr_font(size=11):
    """Pretendard Regular QFont 반환."""
    _load_pretendard()
    families = QFontDatabase.applicationFontFamilies(_PR_FONT_ID)
    return QFont(families[0] if families else 'Pretendard', size)


# ── 업무창·메모장 글꼴 (창마다 고름) ─────────────────────────────
# 리틱체·맑은 고딕은 앱에 넣지 않고 이 PC(Windows)에 설치된 글꼴을 쓴다.
# (리틱체 라이선스: 글꼴 파일 재배포 금지 → 설치 파일에 넣지 않음)
NOTE_FONTS = (('pretendard', 'Pretendard'), ('reatic', '리틱체'), ('reatic_narrow', '리틱체 Narrow'),
              ('malgun', '맑은 고딕'))
DEFAULT_NOTE_FONT = 'pretendard'
# 글꼴 고르기 값 → 이 PC에서 찾을 (글꼴 이름, 모양 이름) 후보. 모양 이름이 ''이면 보통 모양.
# 리틱체 Narrow(폭 좁은 판)는 Windows에 'REATIC Narrow' 같은 글꼴 이름으로, 다른 곳에서는 'REATIC'의 'Narrow' 모양으로 보인다.
_NOTE_SPECS = {
    'reatic': (('리틱체', ''), ('REATIC', '')),
    'reatic_narrow': (('REATIC Narrow', ''), ('리틱체 Narrow', ''), ('REATIC', 'Narrow'), ('리틱체', 'Narrow')),
    'malgun': (('맑은 고딕', ''), ('Malgun Gothic', '')),
}
_CURRENT_NOTE_FONT = DEFAULT_NOTE_FONT
# 같은 pt에서도 글꼴마다 실제 글자 크기가 달라, 한글 글자 크기를 Pretendard에 맞춘다.
# 미리 정한 숫자 대신 이 PC에 설치된 글꼴의 한글 글자 상자(높이·폭)를 직접 재서 정한다 (note_scale).
# 폭 좁은 판(Narrow)은 같은 집안 보통 글꼴의 배율을 그대로 써서 좁은 느낌을 살린다.
NOTE_SCALE_BASE = {'reatic_narrow': 'reatic'}
NOTE_SCALE_SAMPLE = '가나다라마바사아자차카타파하'
NOTE_SCALE_RANGE = (0.7, 1.15)
_MEASURED_SCALE = {}
# 설치 여부를 파일로도 찾을 글꼴: (꼭 있어야 하는 파일, 있으면 함께 쓰는 굵은 파일)
NOTE_FONT_FILES = {'reatic': ('REATIC-Regular', 'REATIC-Bold'),
                   'reatic_narrow': ('REATIC-Narrow', 'REATIC-NarrowBold')}
# 설치되지 않았을 때 안내할 공식 받는 곳 (재배포 금지 글꼴이라 앱·저장소에 넣지 않고 공식 사이트에서 받게 함)
REATIC_SITE = 'https://www.reaticindustry.com/'
REATIC_DOWNLOAD = ('https://www.reaticindustry.com/_files/archives/'
                   '3ae519_5c5d974ad6304c8eb9d143cd9d8a6479.zip?dn=%EB%A6%AC%ED%8B%B1%EC%B2%B4.zip')
NOTE_FONT_DOWNLOAD = {'reatic': REATIC_DOWNLOAD, 'reatic_narrow': REATIC_DOWNLOAD}


def _pretendard_family():
    _load_pretendard()
    families = QFontDatabase.applicationFontFamilies(_PR_FONT_ID)
    return families[0] if families else 'Pretendard'


_INSTALLED = {}                  # 글꼴 고르기 값 → (설치된 글꼴 이름, 모양) (업무 줄마다 찾지 않게 기억)
_LOADED_FILES = set()            # 앱이 직접 불러온 글꼴 파일 (앱을 켠 뒤 설치한 글꼴)


def refresh_installed_fonts():
    """글꼴 메뉴를 열 때: 앱을 켠 뒤 새로 설치한 글꼴도 찾게 기억을 비운다."""
    _INSTALLED.clear()
    _MEASURED_SCALE.clear()


def _match_spec(key):
    """Qt 글꼴 목록에서 (글꼴 이름, 모양 이름) 찾기."""
    cands = _NOTE_SPECS.get(key) or ()
    if not cands:
        return ('', '')
    fdb = QFontDatabase()
    have = set(fdb.families())
    for fam, style in cands:
        if fam not in have:
            continue
        styles = fdb.styles(fam)
        if style and style not in styles:
            continue
        if not style and styles and all(st.lower().startswith('narrow') for st in styles):
            continue                   # 'REATIC'에 Narrow 모양만 있으면 일반 리틱체가 아님
        return (fam, style)
    return ('', '')


def _load_installed_files(key):
    """Windows에 설치된 글꼴 파일을 찾아 앱에 불러온다 (앱을 켠 뒤 설치해도 다시 켤 필요 없게). 불러왔으면 True."""
    stems = NOTE_FONT_FILES.get(key)
    if not stems:
        return False
    try:
        from .font_detect import find_files
        found = find_files(stems)
    except Exception:
        return False
    if stems[0] not in found:
        return False
    loaded = False
    for stem in stems:
        path = found.get(stem)
        if path and path not in _LOADED_FILES:
            if QFontDatabase.addApplicationFont(path) != -1:
                _LOADED_FILES.add(path)
                loaded = True
    return loaded


def _installed_spec(key):
    """이 PC에 설치된 (글꼴 이름, 모양 이름). 없으면 ('', '')."""
    if key not in _INSTALLED:
        spec = _match_spec(key)
        if not spec[0] and _load_installed_files(key):
            spec = _match_spec(key)
        _INSTALLED[key] = spec
    return _INSTALLED[key]


def _installed_family(key):
    """이 PC에 설치된 글꼴 이름 (없으면 '')."""
    return _installed_spec(key)[0]


def note_font_available(key) -> bool:
    return key == 'pretendard' or bool(_installed_family(key))


def note_family(key=None) -> str:
    """글꼴 고르기 값 → 실제 글꼴 이름. 그 PC에 없으면 Pretendard."""
    key = key or _CURRENT_NOTE_FONT
    return (_installed_family(key) if key != 'pretendard' else '') or _pretendard_family()


def set_current_note_font(key):
    """지금 그리는 업무창의 글꼴 (업무 줄을 만들기 전에 정한다 — set_current_accent와 같은 방식)."""
    global _CURRENT_NOTE_FONT
    _CURRENT_NOTE_FONT = key if key in dict(NOTE_FONTS) else DEFAULT_NOTE_FONT


def current_note_font() -> str:
    return _CURRENT_NOTE_FONT


def _hangul_box(family, style=''):
    """48pt 한글 견본의 (글자 높이, 한 글자 폭) — 정밀하게 재려고 큰 크기로."""
    f = QFont(family)
    if style:
        f.setStyleName(style)
    f.setPointSizeF(48)
    m = QFontMetricsF(f)
    r = m.tightBoundingRect(NOTE_SCALE_SAMPLE)
    return r.height(), m.horizontalAdvance(NOTE_SCALE_SAMPLE) / len(NOTE_SCALE_SAMPLE)


def _measure_scale(key):
    """이 글꼴의 한글이 Pretendard와 같은 크기로 보이는 배율 = √(높이 비 × 폭 비)."""
    fam, style = _installed_spec(key)
    if not fam:
        return 1.0
    base = NOTE_SCALE_BASE.get(key)
    if base and _installed_family(base):
        return note_scale(base)                  # Narrow: 보통 글꼴과 같은 배율 (좁은 폭은 그대로)
    ph, pw = _hangul_box(_pretendard_family())
    h, w = _hangul_box(fam, style)
    if h <= 0 or w <= 0:
        return 1.0
    hr = ph / h
    ratio = hr if base else (hr * (pw / w)) ** 0.5   # Narrow만 있으면 높이만 맞춤
    lo, hi = NOTE_SCALE_RANGE
    return min(hi, max(lo, ratio))


def note_scale(key=None) -> float:
    """고른 글꼴의 크기 보정 배율 (그 글꼴이 이 PC에 없어 Pretendard로 그리면 1)."""
    key = key or _CURRENT_NOTE_FONT
    if key == 'pretendard' or not _installed_family(key):
        return 1.0
    if key not in _MEASURED_SCALE:
        _MEASURED_SCALE[key] = _measure_scale(key)
    return _MEASURED_SCALE[key]


def note_font(size=11, key=None):
    """업무창·메모장 글자용 QFont (고른 글꼴, 없으면 Pretendard). 글꼴별 크기 보정을 넣는다."""
    key = key or _CURRENT_NOTE_FONT
    f = QFont(note_family(key))
    style = _installed_spec(key)[1] if key != 'pretendard' else ''
    if style:
        f.setStyleName(style)
    f.setPointSizeF(size * note_scale(key))
    return f


def set_note_bold(font, on=True):
    """굵게. 'Narrow'처럼 모양 이름으로 고른 글꼴은 굵은 모양('Narrow Bold')으로 바꾼다."""
    style = font.styleName()
    if style:
        base = style.replace(' Bold', '').replace('Bold', '').strip() or 'Regular'
        font.setStyleName(f'{base} Bold' if on and base != 'Regular' else ('Bold' if on else base))
    else:
        font.setBold(on)
    return font

def mi_icon(codepoint, size=18, color=None):
    """Fluent Icons 코드포인트를 QIcon으로 변환. 색을 안 주면 테마의 기본 아이콘 회색."""
    if color is None:
        from .theme import tc
        color = tc('#555555')
    _load_fluent_icons()
    pix = QPixmap(size, size)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setFont(mi_font(int(size * 0.82)))
    painter.setPen(QColor(color))
    painter.drawText(pix.rect(), Qt.AlignCenter, codepoint)
    painter.end()
    return QIcon(pix)


# Fluent System Icons 코드포인트 (20px Regular 기준)
class MI:
    STAR          = '\uebaa'   # star_emphasis_20_regular (삐침 별) — Regular 폰트용
    STAR_BORDER   = '\uf70f'   # star_20_regular (아웃라인 별) — Regular 폰트용
    CHECK         = '\uf294'   # checkmark

# Material Icons 코드포인트
class MAT:
    REPEAT        = '\ue040'   # repeat



def themed_pixmap(path) -> QPixmap:
    """아이콘 PNG를 불러온다. 다크 테마에선 검은 단색 아이콘을 밝은 색으로 칠한다 (투명도 유지).

    색이 있는 그림(금색 화살표, 설명서 이미지 등)은 그대로 둔다.
    """
    pm = QPixmap(path)
    from .theme import DARK, T
    if not DARK or pm.isNull():
        return pm
    img = pm.toImage()
    # 불투명 픽셀 표본으로 '어두운 단색' 아이콘인지 판단
    step = max(1, min(img.width(), img.height()) // 24)
    dark = total = 0
    for y in range(0, img.height(), step):
        for x in range(0, img.width(), step):
            c = img.pixelColor(x, y)
            if c.alpha() > 128:
                total += 1
                if c.lightness() < 110 and c.saturation() < 60:
                    dark += 1
    if not total or dark / total < 0.8:
        return pm
    tinted = QPixmap(pm.size())
    tinted.fill(Qt.transparent)
    p = QPainter(tinted)
    p.drawPixmap(0, 0, pm)
    p.setCompositionMode(QPainter.CompositionMode_SourceIn)
    p.fillRect(tinted.rect(), QColor(T.INK))
    p.end()
    return tinted


def _render_svg(data: str, size: int) -> 'QIcon':
    """SVG를 화면 배율(devicePixelRatio)만큼 크게, 가장자리를 부드럽게(안티에일리어싱) 그린 아이콘.
    예전에는 1배 크기로 거칠게 그려 별이 깨져 보였다 (200% 화면에서는 늘려서 흐릿)."""
    from PyQt5.QtWidgets import QApplication
    renderer = QSvgRenderer(data.encode('utf-8'))
    app = QApplication.instance()
    dpr = max(1.0, app.devicePixelRatio()) if app else 1.0
    big = int(round(size * dpr * 2))                 # 두 배로 그린 뒤 부드럽게 줄임
    pm = QPixmap(big, big)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.SmoothPixmapTransform)
    renderer.render(p)
    p.end()
    pm = pm.scaled(int(round(size * dpr)), int(round(size * dpr)), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    pm.setDevicePixelRatio(dpr)
    return QIcon(pm)


def _svg_icon(svg_path, color, size) -> 'QIcon':
    with open(svg_path, 'r', encoding='utf-8') as f:
        data = f.read()
    return _render_svg(data.replace('#000000', color), size)


_STAR_BODY = ('M9.104 2.9c0.367-0.744 1.427-0.744 1.794 0l1.93 3.91 4.317 0.628c0.82 0.12 1.148 1.127 0.554 '
              '1.706l-3.124 3.044 0.738 4.3c0.14 0.816-0.717 1.44-1.451 1.054l-3.86-2.03-3.862 2.03c-0.733 0.385-1.59'
              '-0.238-1.45-1.055l0.737-4.299-3.124-3.044C1.71 8.565 2.037 7.557 2.857 7.438l4.317-0.627 1.93-3.912z')


def outline_star_icon(color, size, stroke=1.9) -> 'QIcon':
    """삐침 없는 빈 별 (테두리만, 굵기 stroke · 둥근 모서리). 화면 배율만큼 크게 그려 선명하게."""
    from PyQt5.QtWidgets import QApplication
    data = (f'<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="-1 -1 22 22">'
            f'<path d="{_STAR_BODY}" fill="none" stroke="{color}" stroke-width="{stroke}" '
            f'stroke-linejoin="round"/></svg>')
    return _render_svg(data, size)


# ── 화면 크기에 맞추기 ───────────────────────────────────────────
# 노트북처럼 작은 화면이나, 큰 모니터에서 쓰던 창 크기·위치를 그대로 열 때 화면 밖으로 나가지 않게 한다.

def setup_high_dpi():
    """화면 배율 설정. QApplication을 만들기 전에 부른다.

    배율은 정수로만(Floor: 125%·150% → 1배, 200% → 2배) 적용하고, 글꼴은 Windows가 배율에 맞는
    크기로 직접 그리게 둔다. Qt 5는 1.25배 같은 소수점 배율에서 글자를 늘려 그리느라 Windows의
    글꼴 다듬기(힌팅·ClearType)를 못 써서 글씨가 흐리고 깨져 보인다 (v2.75에서 겪음).
    대신 글꼴이 커져도 잘리지 않도록 칸·버튼 폭은 고정 숫자 대신 글자 길이로 정한다.
    """
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
    policy = getattr(Qt, 'HighDpiScaleFactorRoundingPolicy', None)
    if policy is not None and hasattr(QApplication, 'setHighDpiScaleFactorRoundingPolicy'):
        QApplication.setHighDpiScaleFactorRoundingPolicy(policy.Floor)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)


def _screen_for(x, y, w):
    from PyQt5.QtCore import QPoint
    from PyQt5.QtWidgets import QApplication
    for pt in (QPoint(x + w // 2, y + 12), QPoint(x + 12, y + 12)):   # 제목 줄 가운데 → 왼쪽 위
        s = QApplication.screenAt(pt)
        if s is not None:
            return s
    return QApplication.primaryScreen()


def fit_to_screen(x, y, w, h, max_frac=0.92, min_w=0, min_h=0):
    """(x, y, w, h)를 그 창이 있는 화면(없으면 주 화면) 안에 들어가게 줄이고 옮긴다.
    화면에서 사라진 모니터에 있던 창은 주 화면 안으로 들어온다."""
    screen = _screen_for(x, y, w)
    if screen is None:
        return x, y, w, h
    ag = screen.availableGeometry()
    w = max(min_w, min(w, int(ag.width() * max_frac)))
    h = max(min_h, min(h, int(ag.height() * max_frac)))
    x = max(ag.left(), min(x, ag.left() + ag.width() - w))
    y = max(ag.top(), min(y, ag.top() + ag.height() - h))
    return x, y, w, h


def default_geometry(pref_w, pref_h, frac_w=0.72, frac_h=0.82):
    """처음 여는 창의 크기: 원하는 크기와 화면 비율 중 작은 쪽, 화면 가운데."""
    from PyQt5.QtWidgets import QApplication
    screen = QApplication.primaryScreen()
    if screen is None:
        return 150, 150, pref_w, pref_h
    ag = screen.availableGeometry()
    w = min(pref_w, int(ag.width() * frac_w))
    h = min(pref_h, int(ag.height() * frac_h))
    return ag.left() + (ag.width() - w) // 2, ag.top() + (ag.height() - h) // 2, w, h


def ui_scale() -> float:
    """글꼴 배율 (Windows 125% → 1.25). 화면은 정수 배율로 그리고 글꼴만 커지므로,
    픽셀로 정한 창·칸 폭에 곱해서 글자가 잘리지 않게 한다."""
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance()
    screen = app.primaryScreen() if app else None
    if screen is None:
        return 1.0
    return max(1.0, min(2.5, screen.logicalDotsPerInch() / 96.0))


def px(n) -> int:
    """픽셀 값을 글꼴 배율에 맞춰 늘린다: px(440) → 125%에서 550."""
    return int(round(n * ui_scale()))


def _render_icon(icon, size: int):
    """QIcon을 size×size 픽셀 그림으로 (화면 배율과 상관없이 딱 그 픽셀 수로 그림)."""
    from PyQt5.QtCore import Qt, QRect
    from PyQt5.QtGui import QPixmap, QPainter
    pm = QPixmap(size, size)
    pm.setDevicePixelRatio(1)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.SmoothPixmapTransform)
    icon.paint(p, QRect(0, 0, size, size))
    p.end()
    return pm


def icon_png(name: str, color: str, size: int = 16) -> str:
    """qtawesome 아이콘을 임시 PNG로 (QSS의 image: url(...)용). 실패하면 ''.

    QSS 그림은 PNG 크기 그대로 그려지므로 글꼴 배율(px)만큼 키워서 저장하고,
    200%·300% 화면용 @2x·@3x 파일도 함께 만든다 (Qt가 화면 배율에 맞는 파일을 골라 씀).
    예전에는 16px 한 장만 만들고 다시 만들지 않아서, 배율이 다른 화면에서는 흐리게 늘어나 보였다.
    """
    try:
        import os
        import tempfile
        import qtawesome as qta
        s = px(size)
        safe = name.replace('.', '_') + '_' + color.strip('#')
        base = os.path.join(tempfile.gettempdir(), f'ssnnote_{safe}_{s}px')
        icon = None
        for n, suffix in ((1, ''), (2, '@2x'), (3, '@3x')):
            path = f'{base}{suffix}.png'
            if not os.path.exists(path):
                icon = icon or qta.icon(name, color=color)
                _render_icon(icon, s * n).save(path)
        return f'{base}.png'.replace('\\', '/')
    except Exception:
        return ''


def fit_pixmap(pixmap, w: int, h: int, widget=None):
    """그림을 w×h(논리 크기) 안에 맞춰 줄이되, 화면 배율만큼 픽셀을 더 담아 선명하게.
    (그냥 scaled(w, h)를 쓰면 200% 화면에서 늘어나 흐려지고, 고배율 캡처는 반 크기로 보인다.)"""
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
    if pixmap is None or pixmap.isNull():
        return pixmap
    r = widget.devicePixelRatioF() if widget is not None else QApplication.instance().devicePixelRatio()
    r = max(1.0, r)
    out = pixmap.scaled(int(round(w * r)), int(round(h * r)), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    out.setDevicePixelRatio(r)
    return out


def scaled_size(n: int):
    """아이콘 크기 QSize — 글꼴 배율(125%·150%)만큼 키워서 글자와 크기를 맞춘다."""
    from PyQt5.QtCore import QSize
    s = px(n)
    return QSize(s, s)


def icon_label(icon, size: int = 16, parent=None, top: int = 0):
    """아이콘만 보여 주는 라벨. 그릴 때마다 벡터로 그려서 어느 배율에서도 또렷하다.
    (QLabel.setPixmap(icon.pixmap(16, 16))은 한 번 만든 그림을 늘려 써서 흐려졌다.)"""
    from PyQt5.QtGui import QPainter
    from PyQt5.QtWidgets import QLabel

    class _IconLabel(QLabel):
        def __init__(self):
            super().__init__(parent)
            self._icon = icon
            s = px(size)
            self._s = s
            t = px(top)                        # 위 여백: 옆 글자 첫 줄에 맞출 때
            self.setContentsMargins(0, t, 0, 0)
            self.setFixedSize(s, s + t)

        def setIcon(self, ic):
            self._icon = ic
            self.update()

        def paintEvent(self, _e):
            if self._icon is None:
                return
            p = QPainter(self)
            p.setRenderHint(QPainter.Antialiasing)
            p.setRenderHint(QPainter.SmoothPixmapTransform)
            r = self.contentsRect()
            s = min(self._s, r.width(), r.height())
            self._icon.paint(p, r.x() + (r.width() - s) // 2, r.y(), s, s)

    return _IconLabel()


# ── 앱 아이콘 (노란 메모지 + 줄 세 개) ─────────────────────────────
APP_ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)


def draw_app_icon(n: int):
    """앱 아이콘을 n×n 픽셀로 그린다 (32px 도안을 비율대로 · 작은 크기에서도 선이 뭉개지지 않게)."""
    from PyQt5.QtCore import Qt, QRectF, QLineF
    from PyQt5.QtGui import QPixmap, QPainter, QColor, QPen
    k = n / 32
    pm = QPixmap(n, n)
    pm.setDevicePixelRatio(1)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    edge = max(1.0, round(k))                       # 테두리 굵기 (16px에서도 1px)
    border = QPen(QColor('#c8a000'))
    border.setWidthF(edge)
    p.setPen(border)
    p.setBrush(QColor('#f7c948'))
    m = 2 * k + edge / 2
    p.drawRoundedRect(QRectF(m, m, n - 2 * m, n - 2 * m), 4 * k, 4 * k)
    lw = max(1.0, round(2 * k))                     # 줄 굵기 (픽셀 경계에 맞춰 또렷하게)
    pen = QPen(QColor('#7a6000'))
    pen.setWidthF(lw)
    pen.setCapStyle(Qt.RoundCap if n >= 24 else Qt.FlatCap)
    p.setPen(pen)
    off = (lw % 2) / 2                              # 홀수 굵기는 반 픽셀 옮겨야 흐리지 않음
    for y, x2 in ((10, 25), (16, 25), (22, 18)):
        yy = round(y * k) + off
        p.drawLine(QLineF(round(7 * k), yy, round(x2 * k), yy))
    p.end()
    return pm


def app_icon():
    """창 · 작업 표시줄 · 트레이 아이콘: 크기마다 따로 그려 넣어 어느 배율에서도 또렷하게.
    (예전에는 32px 한 장을 늘려 써서 고배율 화면 · 큰 작업 표시줄 아이콘이 흐렸다.)"""
    from PyQt5.QtGui import QIcon
    icon = QIcon()
    for n in APP_ICON_SIZES:
        icon.addPixmap(draw_app_icon(n))
    return icon

