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

def mi_icon(codepoint, size=18, color='#555555'):
    """Fluent Icons 코드포인트를 QIcon으로 변환."""
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
