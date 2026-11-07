"""ui/menu_style.py — 메뉴 모양 (Notion·Linear·Apple Notes 같은 미니멀한 팝업 메뉴).

QSS 대신 QProxyStyle로 직접 그린다 — QSS로는 단축키 글자색·둥근 그림자·체크 표시를 따로 다룰 수 없어서.
  · 흰 바탕(다크: 짙은 회색) + 아주 옅은 테두리 + 부드러운 그림자 + 둥근 모서리
  · 마우스를 올린 항목은 둥근 알약 모양으로 옅게 칠함
  · 왼쪽 선 아이콘, 오른쪽 옅은 단축키, 하위 메뉴는 얇은 꺾쇠
  · '삭제'가 들어간 항목은 빨간 글씨, 누를 수 없는 항목은 작은 회색 머리글

    menu = make_menu(parent)              # 또는 style_menu(QMenu(...))
    add_action(menu, '빠른 찾기', slot, icon='ph.magnifying-glass', shortcut='Ctrl+K')
    sub = add_submenu(menu, '설정', icon='ph.gear-six')
"""

from PyQt5.QtCore import QPointF, QRect, QRectF, QSize, Qt
from PyQt5.QtGui import QColor, QFont, QFontMetrics, QIcon, QPainter, QPen
from PyQt5.QtWidgets import QAction, QMenu, QProxyStyle, QStyle, QStyleOptionMenuItem

from .theme import DARK

SHADOW = 10          # 그림자 자리 (메뉴 둘레 여백)
RADIUS = 10
ITEM_RADIUS = 6

if DARK:
    BG = QColor('#232326')
    BORDER = QColor(255, 255, 255, 22)
    TEXT = QColor('#E8E8EA')
    MUTED = QColor('#8E8E93')
    HOVER = QColor(255, 255, 255, 18)
    SEP = QColor(255, 255, 255, 20)
    DANGER = QColor('#FF6B6B')
    ICON = QColor('#B4B4B8')
else:
    BG = QColor('#FFFFFF')
    BORDER = QColor(0, 0, 0, 20)
    TEXT = QColor('#1F2328')
    MUTED = QColor('#8A8F98')
    HOVER = QColor(0, 0, 0, 12)
    SEP = QColor(0, 0, 0, 16)
    DANGER = QColor('#D93025')
    ICON = QColor('#5F6368')


def _font():
    f = QFont('Pretendard GOV', 10)
    f.setFamilies(['Pretendard GOV', 'Pretendard', 'Malgun Gothic'])
    return f


def _is_danger(label: str) -> bool:
    return '삭제' in label or label.startswith('종료')


class MenuStyle(QProxyStyle):
    def pixelMetric(self, metric, option=None, widget=None):
        if metric == QStyle.PM_MenuPanelWidth:
            return 0
        if metric == QStyle.PM_MenuHMargin:
            return SHADOW + 4
        if metric == QStyle.PM_MenuVMargin:
            return SHADOW + 5
        if metric == QStyle.PM_SmallIconSize:
            return 16
        if metric == QStyle.PM_SubMenuOverlap:
            return -(SHADOW * 2 - 2)          # 그림자 자리만큼 겹쳐서 하위 메뉴가 바로 옆에 붙게
        if metric == QStyle.PM_MenuDesktopFrameWidth:
            return 0
        return super().pixelMetric(metric, option, widget)

    def styleHint(self, hint, option=None, widget=None, returnData=None):
        if hint == QStyle.SH_Menu_Scrollable:
            return 0
        if hint == QStyle.SH_Menu_SubMenuPopupDelay:
            return 120
        return super().styleHint(hint, option, widget, returnData)

    # ── 크기 ────────────────────────────────────────────────────
    def sizeFromContents(self, ct, opt, size, widget=None):
        if ct != QStyle.CT_MenuItem or not isinstance(opt, QStyleOptionMenuItem):
            return super().sizeFromContents(ct, opt, size, widget)
        fm = QFontMetrics(opt.font)
        if opt.menuItemType == QStyleOptionMenuItem.Separator:
            return QSize(size.width(), 9)
        label, _, shortcut = opt.text.partition('\t')
        header = not (opt.state & QStyle.State_Enabled) and opt.icon.isNull()
        h = max(fm.height() + 14, 30)
        if header:
            h = max(fm.height() + 8, 24)
        w = 12 + (26 if opt.maxIconWidth > 0 else 0) + fm.horizontalAdvance(label) + 14
        if shortcut:
            w += fm.horizontalAdvance(shortcut) + 28
        if opt.menuItemType == QStyleOptionMenuItem.SubMenu:
            w += 22
        return QSize(max(w + 8, 150), h)

    # ── 바탕 ────────────────────────────────────────────────────
    def drawPrimitive(self, element, opt, painter, widget=None):
        if element == QStyle.PE_PanelMenu:
            painter.save()
            painter.setRenderHint(QPainter.Antialiasing)
            r = QRectF(opt.rect).adjusted(SHADOW, SHADOW - 2, -SHADOW, -SHADOW - 2)
            # 부드러운 그림자: 바깥으로 갈수록 옅게
            for i in range(SHADOW, 0, -1):
                a = int(12 * (1 - i / SHADOW) ** 2) + 1
                painter.setPen(Qt.NoPen)
                painter.setBrush(QColor(0, 0, 0, a if not DARK else a * 2))
                painter.drawRoundedRect(r.adjusted(-i, -i + 3, i, i + 3), RADIUS + i, RADIUS + i)
            painter.setBrush(BG)
            painter.setPen(QPen(BORDER, 1))
            painter.drawRoundedRect(r.adjusted(0.5, 0.5, -0.5, -0.5), RADIUS, RADIUS)
            painter.restore()
            return
        if element == QStyle.PE_FrameMenu:
            return
        super().drawPrimitive(element, opt, painter, widget)

    # ── 항목 ────────────────────────────────────────────────────
    def drawControl(self, element, opt, painter, widget=None):
        if element == QStyle.CE_MenuEmptyArea:
            return
        if element != QStyle.CE_MenuItem or not isinstance(opt, QStyleOptionMenuItem):
            return super().drawControl(element, opt, painter, widget)
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        rect = QRect(opt.rect).adjusted(4, 1, -4, -1)
        enabled = bool(opt.state & QStyle.State_Enabled)
        selected = bool(opt.state & QStyle.State_Selected) and enabled

        if opt.menuItemType == QStyleOptionMenuItem.Separator:
            y = opt.rect.center().y() + 0.5
            painter.setPen(QPen(SEP, 1))
            painter.drawLine(QPointF(opt.rect.left() + 10, y), QPointF(opt.rect.right() - 10, y))
            painter.restore()
            return

        label, _, shortcut = opt.text.partition('\t')
        f = QFont(opt.font)

        if not enabled and opt.icon.isNull():
            # 누를 수 없는 항목 = 작은 회색 머리글 (예: '배경색', 'D-day 5일 이내 알림')
            f.setPointSizeF(max(8.0, f.pointSizeF() - 1.5))
            painter.setFont(f)
            painter.setPen(MUTED)
            painter.drawText(rect.adjusted(8, 0, -8, 0), Qt.AlignLeft | Qt.AlignBottom, label)
            painter.restore()
            return

        if selected:
            painter.setPen(Qt.NoPen)
            painter.setBrush(HOVER)
            painter.drawRoundedRect(QRectF(rect), ITEM_RADIUS, ITEM_RADIUS)

        danger = _is_danger(label)
        text_color = DANGER if danger else (TEXT if enabled else MUTED)
        x = rect.left() + 8
        if opt.maxIconWidth > 0:
            if not opt.icon.isNull():
                mode = QIcon.Normal if enabled else QIcon.Disabled
                pm = opt.icon.pixmap(QSize(16, 16), mode)
                painter.drawPixmap(x + 1, rect.center().y() - 8, pm)
            elif opt.checkType != QStyleOptionMenuItem.NotCheckable and opt.checked:
                _draw_check(painter, QRectF(x + 1, rect.center().y() - 8, 16, 16), text_color)
            x += 26

        painter.setFont(f)
        painter.setPen(text_color)
        right = rect.right() - 10
        if opt.menuItemType == QStyleOptionMenuItem.SubMenu:
            cx, cy = rect.right() - 12, rect.center().y() + 0.5
            painter.setPen(QPen(MUTED, 1.4, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            painter.drawPolyline(QPointF(cx - 2.5, cy - 4), QPointF(cx + 1.5, cy), QPointF(cx - 2.5, cy + 4))
            right -= 18
        if shortcut:
            sf = QFont(f)
            sf.setPointSizeF(max(8.0, f.pointSizeF() - 1))
            painter.setFont(sf)
            painter.setPen(MUTED)
            painter.drawText(QRect(x, rect.top(), right - x, rect.height()), Qt.AlignRight | Qt.AlignVCenter, shortcut)
            painter.setFont(f)
        painter.setPen(text_color)
        painter.drawText(QRect(x, rect.top(), right - x, rect.height()), Qt.AlignLeft | Qt.AlignVCenter, label)
        painter.restore()


def _draw_check(painter, r: QRectF, color):
    painter.save()
    painter.setPen(QPen(color, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    painter.drawPolyline(QPointF(r.left() + r.width() * 0.2, r.top() + r.height() * 0.52),
                         QPointF(r.left() + r.width() * 0.42, r.top() + r.height() * 0.74),
                         QPointF(r.left() + r.width() * 0.82, r.top() + r.height() * 0.28))
    painter.restore()


_STYLE = None


def _style():
    global _STYLE
    if _STYLE is None:
        _STYLE = MenuStyle()        # 앱이 끝날 때까지 하나를 같이 쓴다 (메뉴보다 먼저 사라지면 안 됨)
    return _STYLE


def style_menu(menu: QMenu) -> QMenu:
    """기존 QMenu에 새 모양을 입힌다 (하위 메뉴는 add_submenu로 만들면 자동)."""
    menu.setStyle(_style())
    menu.setStyleSheet('')
    menu.setAttribute(Qt.WA_TranslucentBackground, True)
    menu.setWindowFlags(menu.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
    menu.setFont(_font())
    menu.setToolTipsVisible(True)
    return menu


def make_menu(parent=None, title='') -> QMenu:
    return style_menu(QMenu(title, parent) if title else QMenu(parent))


def line_icon(name: str, color=None) -> QIcon:
    try:
        import qtawesome as qta
        return qta.icon(name, color=color or ICON)
    except Exception:
        return QIcon()


def check_icon(checked: bool) -> QIcon:
    """켜짐은 ✓, 꺼짐은 빈 칸 (글자 줄이 맞도록 같은 크기)."""
    from PyQt5.QtGui import QPixmap
    pm = QPixmap(32, 32)
    pm.setDevicePixelRatio(2)
    pm.fill(Qt.transparent)
    if checked:
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        _draw_check(p, QRectF(0, 0, 16, 16), TEXT)
        p.end()
    return QIcon(pm)


def add_action(menu: QMenu, text, slot=None, icon=None, shortcut='', checked=None, tip=''):
    """항목 추가. icon: 'ph.…' 이름, checked: True/False면 체크 항목(✓)."""
    act = QAction(text + (f'\t{shortcut}' if shortcut else ''), menu)
    if checked is not None:
        act.setIcon(check_icon(checked))
    elif icon:
        act.setIcon(line_icon(icon, DANGER if _is_danger(text) else None))
    if slot is not None:
        act.triggered.connect(lambda *_: slot())
    if tip:
        act.setToolTip(tip)
    menu.addAction(act)
    return act


def add_header(menu: QMenu, text):
    """누를 수 없는 작은 머리글."""
    act = QAction(text, menu)
    act.setEnabled(False)
    menu.addAction(act)
    return act


def add_submenu(menu: QMenu, title, icon=None) -> QMenu:
    sub = make_menu(menu, title)
    if icon:
        sub.setIcon(line_icon(icon))
    menu.addMenu(sub)
    return sub
