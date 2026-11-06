"""ui/common.py — 공통 상수·경로·아이콘 폰트(Fluent/Material/Pretendard) 로더."""
import os
import sys
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor, QPainter, QIcon, QPixmap, QFontDatabase

def _base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    # 이 파일은 ui/ 안에 있으므로 프로젝트 루트(assets/가 있는 곳)는 한 단계 위
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TITLE_BAR_HEIGHT = 33
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
            os.path.join(base, 'assets', 'FluentSystemIcons-Regular.ttf'))
    if _FI_FILLED_FONT_ID == -1:
        _FI_FILLED_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'FluentSystemIcons-Filled.ttf'))
    if _MAT_FONT_ID == -1:
        _MAT_FONT_ID = QFontDatabase.addApplicationFont(
            os.path.join(base, 'assets', 'MaterialIcons-Regular.ttf'))

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
    """assets 안의 Pretendard 정적 폰트(여러 굵기)를 앱 폰트로 등록한다.

    QSS의 font-family: 'Pretendard' 가 이 등록된 패밀리를 쓰고,
    굵은 글씨는 합성 볼드 대신 실제 굵기 파일로 그려진다.
    Regular 파일의 id는 pr_font()용으로 _PR_FONT_ID에 남긴다.
    """
    global _PR_FONT_ID, _PRETENDARD_LOADED
    if _PRETENDARD_LOADED:
        return
    _PRETENDARD_LOADED = True
    for dirpath, _dirs, files in os.walk(os.path.join(_base_path(), 'assets')):
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


def _svg_icon(svg_path, color, size) -> 'QIcon':
    with open(svg_path, 'r', encoding='utf-8') as f:
        data = f.read()
    data = data.replace('#000000', color)
    renderer = QSvgRenderer(data.encode('utf-8'))
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    renderer.render(p)
    p.end()
    return QIcon(pm)


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
