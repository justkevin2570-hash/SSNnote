"""ui/chrome.py — 메모 창 외곽 요소: 가장자리 핸들, 병합 버튼, 제목 표시줄, 스냅 가이드, 긴급 알림 토스트."""
import os
import ctypes
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QApplication,
    QFrame, QGraphicsOpacityEffect
)
from PyQt5.QtCore import Qt, QTimer, QPoint, QSize, QPropertyAnimation, QEasingCurve, QRectF
from PyQt5.QtGui import QFont, QColor, QPainter, QPixmap, QCursor
import qtawesome as qta
from .common import _base_path, BREAKAWAY_DIST, SNAP_RESIST, SNAP_THRESHOLD, TITLE_BAR_HEIGHT, TITLE_COLOR, px
from .dates import calc_dday
from .theme import T, tc, tct

class EdgeHandle(QWidget):
    EDGE  = 6   # 가장자리 감지 두께(px)
    CORNER = 16  # 모서리 감지 크기(px)
    MIN_W = 320
    MIN_H = 100

    _CURSORS = {
        'left':         Qt.SizeHorCursor,
        'right':        Qt.SizeHorCursor,
        'bottom':       Qt.SizeVerCursor,
        'bottom-left':  Qt.SizeBDiagCursor,
        'bottom-right': Qt.SizeFDiagCursor,
        'top':          Qt.SizeVerCursor,
        'top-left':     Qt.SizeFDiagCursor,
        'top-right':    Qt.SizeBDiagCursor,
    }

    def __init__(self, parent, edge):
        super().__init__(parent)
        self.edge         = edge
        self._drag_start  = None
        self._start_geom  = None
        self.setCursor(self._CURSORS[edge])
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.setStyleSheet('background: transparent;')

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_start = e.globalPos()
            self._start_geom = self.window().geometry()
            g = self._start_geom
            self._right_anchor = g.x() + g.width()   # 왼쪽 드래그 시 고정되는 오른쪽 끝
            self._bottom_anchor = g.y() + g.height()  # 위쪽 드래그 시 고정되는 아래쪽 끝
            self.grabKeyboard()
            # 리사이즈 중 합체 파트너가 딸려오지 않도록 양쪽 동기화 일시 정지
            win = self.window()
            if getattr(win, '_link_group', None):
                for m in win._link_group:
                    m._is_syncing_move = True

    def _anchor_partner(self, win):
        """파트너가 리사이즈로 움직이는 쪽 모서리에 붙어 있을 때만 따라붙게 한다.
        고정된 모서리 쪽에 파트너가 있으면 아무것도 하지 않는다."""
        if not win._link_group:
            return
        # 리사이즈 방향에 인접한 그룹 멤버를 파트너로 선택
        partner = None
        if self.edge in ('right', 'top-right', 'bottom-right'):
            cands = [m for m in win._link_group if m is not win and m.x() > win.x()]
            if cands: partner = min(cands, key=lambda m: m.x())
        elif self.edge in ('left', 'top-left', 'bottom-left'):
            cands = [m for m in win._link_group if m is not win and m.x() < win.x()]
            if cands: partner = max(cands, key=lambda m: m.x())
        elif self.edge in ('bottom', 'bottom-left', 'bottom-right'):
            cands = [m for m in win._link_group if m is not win and m.y() > win.y()
                     and m.x() < win.x() + win.width() and m.x() + m.width() > win.x()]
            if cands: partner = min(cands, key=lambda m: m.y())
        elif self.edge in ('top', 'top-left', 'top-right'):
            cands = [m for m in win._link_group if m is not win and m.y() < win.y()
                     and m.x() < win.x() + win.width() and m.x() + m.width() > win.x()]
            if cands: partner = max(cands, key=lambda m: m.y())
        if partner is None:
            return
        partner_on_right = win.x() < partner.x()
        partner_on_bottom = win.y() < partner.y()
        nx, ny = partner.x(), partner.y()
        need_move = False
        if self.edge in ('right', 'top-right', 'bottom-right') and partner_on_right:
            nx = win.x() + win.width()
            need_move = True
        elif self.edge in ('left', 'top-left', 'bottom-left') and not partner_on_right:
            nx = win.x() - partner.width()
            need_move = True
        if self.edge in ('bottom', 'bottom-left', 'bottom-right') and partner_on_bottom:
            ny = win.y() + win.height()
            need_move = True
        elif self.edge in ('top', 'top-left', 'top-right') and not partner_on_bottom:
            ny = win.y() - partner.height()
            need_move = True
        if need_move:
            partner.move(nx, ny)

    def mouseMoveEvent(self, e):
        if not (e.buttons() == Qt.LeftButton and self._drag_start):
            return
        d   = e.globalPos() - self._drag_start
        g   = self._start_geom
        win = self.window()

        nx, ny = g.x(), g.y()
        nw, nh = g.width(), g.height()

        if self.edge == 'left':
            nx = min(g.x() + d.x(), self._right_anchor - self.MIN_W)
            nw = self._right_anchor - nx
        elif self.edge == 'right':
            nw = g.width() + d.x()
        elif self.edge == 'bottom':
            nh = g.height() + d.y()
        elif self.edge == 'top':
            ny = min(g.y() + d.y(), self._bottom_anchor - self.MIN_H)
            nh = self._bottom_anchor - ny
        elif self.edge == 'bottom-right':
            nw = g.width() + d.x()
            nh = g.height() + d.y()
        elif self.edge == 'bottom-left':
            nx = min(g.x() + d.x(), self._right_anchor - self.MIN_W)
            nw = self._right_anchor - nx
            nh = g.height() + d.y()
        elif self.edge == 'top-right':
            nw = g.width() + d.x()
            ny = min(g.y() + d.y(), self._bottom_anchor - self.MIN_H)
            nh = self._bottom_anchor - ny
        elif self.edge == 'top-left':
            nx = min(g.x() + d.x(), self._right_anchor - self.MIN_W)
            nw = self._right_anchor - nx
            ny = min(g.y() + d.y(), self._bottom_anchor - self.MIN_H)
            nh = self._bottom_anchor - ny

        nw = max(self.MIN_W, nw)
        nh = max(self.MIN_H, nh)

        nx, ny, nw, nh = win._snap_resize(nx, ny, nw, nh, self.edge)

        win.setGeometry(nx, ny, nw, nh)
        self._anchor_partner(win)

    def mouseReleaseEvent(self, e):
        self.releaseKeyboard()
        self._drag_start = None
        self._start_geom = None
        # 리사이즈 완료 후 offset 재계산 후 동기화 재개
        win = self.window()
        if getattr(win, '_link_group', None):
            for m in win._link_group:
                m._is_syncing_move = False
        win.save_state()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape and self._drag_start:
            self.window().setGeometry(self._start_geom)
            self.releaseKeyboard()
            self._drag_start = None
            self._start_geom = None


class MergeButton(QWidget):
    """두 창 경계선에 떠있는 합체/분리 토글 버튼. 호버 시에만 표시."""
    SIZE_W = 45
    SIZE_H = 56
    _HOVER_RADIUS = 38

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFixedSize(self.SIZE_W, self.SIZE_H)
        self._merged  = False
        self._win_a   = None
        self._win_b   = None
        self._visible = False

        self._raise_timer = QTimer(self)
        self._raise_timer.setInterval(150)
        self._raise_timer.timeout.connect(self._do_raise)

        self._opacity_eff = QGraphicsOpacityEffect(self)
        self._opacity_eff.setOpacity(0.0)
        self.setGraphicsEffect(self._opacity_eff)
        self._opacity_anim = QPropertyAnimation(self._opacity_eff, b'opacity')
        self._opacity_anim.setDuration(150)

        self._hover_timer = QTimer(self)
        self._hover_timer.setInterval(80)
        self._hover_timer.timeout.connect(self._check_proximity)

        self._hide_delay = QTimer(self)
        self._hide_delay.setSingleShot(True)
        self._hide_delay.setInterval(1000)
        self._hide_delay.timeout.connect(self._do_fade_out)

        self._auto_hide_timer = QTimer(self)
        self._auto_hide_timer.setSingleShot(True)
        self._auto_hide_timer.setInterval(2000)
        self._auto_hide_timer.timeout.connect(self._auto_hide)

    def set_windows(self, win_a, win_b):
        self._win_a = win_a
        self._win_b = win_b

    def set_merged(self, merged: bool):
        self._merged = merged
        self.update()

    def update_position(self, win_a, win_b):
        ga = win_a.geometry()
        gb = win_b.geometry()
        if ga.right() <= gb.left() + SNAP_THRESHOLD:
            bx = ga.right()
        else:
            bx = gb.right()
        overlap_top    = max(ga.top(), gb.top())
        overlap_bottom = min(ga.bottom(), gb.bottom())
        center_y = (overlap_top + overlap_bottom) // 2
        self.move(bx - self.SIZE_W // 2, center_y - self.SIZE_H // 2)

    def _do_fade_out(self):
        cursor = QCursor.pos()
        center = self.mapToGlobal(QPoint(self.SIZE_W // 2, self.SIZE_H // 2))
        dx = cursor.x() - center.x()
        dy = cursor.y() - center.y()
        if (dx * dx + dy * dy) > self._HOVER_RADIUS ** 2:
            self._visible = False
            self._opacity_anim.stop()
            self._opacity_anim.setEndValue(0.0)
            self._opacity_anim.start()

    def _check_proximity(self):
        cursor = QCursor.pos()
        center = self.mapToGlobal(QPoint(self.SIZE_W // 2, self.SIZE_H // 2))
        dx = cursor.x() - center.x()
        dy = cursor.y() - center.y()
        near = (dx * dx + dy * dy) <= self._HOVER_RADIUS ** 2
        if near and not self._visible:
            self._hide_delay.stop()
            self._visible = True
            self._opacity_anim.stop()
            self._opacity_anim.setEndValue(1.0)
            self._opacity_anim.start()
            self.setCursor(Qt.PointingHandCursor)
        elif not near and self._visible and not self._hide_delay.isActive():
            self._hide_delay.start()
            self._opacity_anim.start()

    def _do_raise(self):
        try:
            ctypes.windll.user32.SetWindowPos(
                int(self.winId()), 0, 0, 0, 0, 0,
                0x0010 | 0x0001 | 0x0002)
        except Exception:
            pass

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        bw, bh = 16, 16
        by = (self.SIZE_H - bh) // 2
        f = QFont('Arial', 10, QFont.Bold)
        p.setFont(f)
        if self._merged:
            nx, sx = 6, 22
        else:
            nx, sx = 2, 27
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(T.DANGER))
        p.drawRoundedRect(nx, by, bw, bh, 4, 4)
        p.setBrush(QColor(tc('#3498db')))
        p.drawRoundedRect(sx, by, bw, bh, 4, 4)
        if not self._merged:
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(nx + bw, by, sx - (nx + bw), bh)
        p.setPen(QColor(255, 255, 255))
        p.drawText(QRectF(nx, by, bw, bh), Qt.AlignCenter, 'N')
        p.drawText(QRectF(sx, by, bw, bh), Qt.AlignCenter, 'S')
        p.end()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._win_a and self._win_b:
            self._win_a.toggle_merge(self._win_b)

    def _auto_hide(self):
        cursor = QCursor.pos()
        center = self.mapToGlobal(QPoint(self.SIZE_W // 2, self.SIZE_H // 2))
        dx = cursor.x() - center.x()
        dy = cursor.y() - center.y()
        if (dx * dx + dy * dy) > self._HOVER_RADIUS ** 2:
            self._visible = False
            self._opacity_anim.stop()
            self._opacity_anim.setEndValue(0.0)
            self._opacity_anim.start()

    def showEvent(self, event):
        super().showEvent(event)
        self._visible = True
        self._opacity_anim.stop()
        self._opacity_anim.setEndValue(1.0)
        self._opacity_anim.start()
        self._hover_timer.start()
        QTimer.singleShot(2000, self._auto_hide)

    def hideEvent(self, event):
        super().hideEvent(event)
        self._raise_timer.stop()
        self._hover_timer.stop()
        self._visible = False


class UrgentBadge(QWidget):
    """제목 줄의 급한 업무 배지: '지남 1 · 오늘 2'. 기한 지난 업무가 있으면 빨강, 아니면 주황.
    천천히 숨쉬듯 빛나서 메모 창을 접어 두거나 멀리서 봐도 '급한 게 있다'를 알 수 있다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._text = ''
        self._color = '#F76B15'
        self._breathe = 0.0
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        f = QFont('Pretendard GOV', 9)
        f.setFamilies(['Pretendard GOV', 'Pretendard', 'Malgun Gothic'])
        f.setWeight(QFont.DemiBold)
        self.setFont(f)
        self.hide()

    def set_counts(self, overdue: int, today: int, animate: bool):
        from .motion import clock
        parts = ([f'지남 {overdue}'] if overdue else []) + ([f'오늘 {today}'] if today else [])
        self._text = ' · '.join(parts)
        self._color = '#E5484D' if overdue else '#F76B15'
        self._breathe = (2.5 if overdue else 4.0) if (animate and parts) else 0.0
        self.setVisible(bool(parts))
        self.setToolTip('기한이 지났거나 오늘 마감인 업무' if parts else '')
        if self._breathe:
            clock().register(self, lambda t: self.update())
        else:
            clock().unregister(self)
        self.updateGeometry()
        self.update()

    def sizeHint(self):
        from PyQt5.QtGui import QFontMetrics
        fm = QFontMetrics(self.font())
        return QSize(fm.horizontalAdvance(self._text) + 18 + 8, max(fm.height() + 4, 18) + 8)

    def paintEvent(self, _):
        from PyQt5.QtGui import QFontMetrics
        from .motion import breath, clock
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        fm = QFontMetrics(self.font())
        h = max(fm.height() + 2, 18)
        w = fm.horizontalAdvance(self._text) + 16
        r = QRectF((self.width() - w) / 2.0, (self.height() - h) / 2.0, w, h)
        col = QColor(tc(self._color))
        if self._breathe:
            a = breath(clock().now(), self._breathe)
            glow = QColor(col)
            glow.setAlphaF(0.22 * a)
            p.setPen(Qt.NoPen)
            p.setBrush(glow)
            g = r.adjusted(-3 * a, -3 * a, 3 * a, 3 * a)
            p.drawRoundedRect(g, g.height() / 2, g.height() / 2)
        p.setPen(Qt.NoPen)
        p.setBrush(col)
        p.drawRoundedRect(r, h / 2, h / 2)
        p.setPen(QColor('#FFFFFF'))
        p.drawText(r, Qt.AlignCenter, self._text)
        p.end()


class TitleBar(QWidget):
    """메모 창 제목 줄: 이름 · 오늘 날짜 | 예산 · 메모 모드 · 메뉴 · 항상 위 · 닫기.
    아이콘은 모두 같은 얇은 선 아이콘(같은 크기·같은 색)이고, 켜진 버튼(메모 모드·항상 위)만 바탕을 칠한다."""

    _ICON_PX = 20

    def __init__(self, parent):
        super().__init__(parent)
        self.parent    = parent
        self._drag_pos = None
        self.setFixedHeight(TITLE_BAR_HEIGHT)
        self.apply_style(TITLE_COLOR)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 0, 6, 0)
        layout.setSpacing(1)

        self.label = QLabel('서서니 노트')
        f = QFont('Pretendard GOV', 10)
        f.setFamilies(['Pretendard GOV', 'Pretendard', 'Malgun Gothic'])
        f.setWeight(QFont.DemiBold)
        self.label.setFont(f)
        self.label.setStyleSheet(f'color: {T.TEXT_ON_NOTE}; background: transparent;')
        self.date_label = QLabel('')
        fd = QFont(f)
        fd.setWeight(QFont.Normal)
        fd.setPointSizeF(9.5)
        self.date_label.setFont(fd)
        self.date_label.setStyleSheet(f'color: {tc("rgba(0,0,0,0.45)")}; background: transparent; padding-left: 6px;')
        self._update_date()
        self._date_timer = QTimer(self)
        self._date_timer.timeout.connect(self._update_date)
        self._date_timer.start(60 * 1000)

        btn_menu = self._icon_button('ph.dots-three-bold', '메뉴')
        self.btn_menu = btn_menu
        self.btn_pin = self._icon_button('ph.push-pin-bold', '항상 위 꺼짐')
        self.btn_close = self._icon_button('ph.x-bold', '닫기')
        self.btn_mode = self._icon_button('ph.note-pencil-bold', '메모 모드로 전환')
        self.btn_mode.setObjectName('btn_mode')
        self.btn_mode.clicked.connect(parent._toggle_memo_mode)
        self.btn_budget = self._icon_button('ph.currency-krw-bold', '예산 관리')
        self.btn_budget.setObjectName('btn_budget')
        self.btn_budget.clicked.connect(parent._open_budget_window)

        layout.addWidget(self.label)
        layout.addWidget(self.date_label)
        self.badge = UrgentBadge(self)
        layout.addSpacing(4)
        layout.addWidget(self.badge, 0, Qt.AlignVCenter)
        layout.addStretch()
        layout.addWidget(self.btn_budget, 0, Qt.AlignVCenter)
        layout.addWidget(self.btn_mode, 0, Qt.AlignVCenter)
        layout.addWidget(btn_menu, 0, Qt.AlignVCenter)
        layout.addWidget(self.btn_pin, 0, Qt.AlignVCenter)
        layout.addWidget(self.btn_close, 0, Qt.AlignVCenter)

        # 가운데 오버레이 상태 메시지 (전체 너비, 마우스 이벤트 통과)
        self.status_label = QLabel('', self)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont('Malgun Gothic', 9))
        self.status_label.setStyleSheet(f'color: {T.TEXT_ON_NOTE}; background: transparent;')
        self.status_label.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.status_label.setGeometry(0, 0, self.width(), TITLE_BAR_HEIGHT)

        btn_menu.clicked.connect(lambda: parent.show_menu(btn_menu))

        self.btn_pin.clicked.connect(parent.toggle_always_on_top)
        self.btn_close.clicked.connect(parent.close)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.status_label.setGeometry(0, 0, self.width(), self.height())

    # ── 모양 ────────────────────────────────────────────────────
    @staticmethod
    def _icon_color(active=False):
        from .theme import qcolor
        return qcolor('#000000', 225 if active else 165)

    def _icon_button(self, icon_name, tip):
        b = QPushButton()
        b._icon_name = icon_name
        b.setIcon(qta.icon(icon_name, color=self._icon_color()))
        b.setIconSize(QSize(px(self._ICON_PX), px(self._ICON_PX)))   # 글꼴 배율만큼 아이콘도
        b.setFixedSize(px(32), min(px(30), TITLE_BAR_HEIGHT - 4))
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        return b

    def set_icon_scale(self, scale: float):
        """「글자 크기」(기본·크게·더 크게)에 맞춰 제목 줄 아이콘도 키운다 (제목 줄 높이 안에서)."""
        size = min(px(round(self._ICON_PX * scale)), TITLE_BAR_HEIGHT - 12)
        for b in (self.btn_budget, self.btn_mode, self.btn_menu, self.btn_pin, self.btn_close):
            b.setIconSize(QSize(size, size))
            b.setFixedSize(max(px(32), size + 12), min(max(px(30), size + 10), TITLE_BAR_HEIGHT - 4))
        f = self.label.font()
        f.setPointSizeF(10 * scale)
        self.label.setFont(f)

    def set_active(self, button, on: bool, icon_name=None):
        """켜진 상태(메모 모드·항상 위)는 진한 아이콘 + 옅은 바탕."""
        if icon_name:
            button._icon_name = icon_name
        button.setIcon(qta.icon(button._icon_name, color=self._icon_color(on)))
        button.setProperty('active', on)
        button.style().unpolish(button)
        button.style().polish(button)

    def apply_style(self, bg):
        self.setStyleSheet(f"""
            QWidget {{ background: {bg}; border-radius: 0; }}
            QPushButton {{ background: transparent; border: none; border-radius: 6px; padding: 0; }}
            QPushButton:hover {{ background: {tc("rgba(0,0,0,0.10)")}; }}
            QPushButton:pressed {{ background: {tc("rgba(0,0,0,0.18)")}; }}
            QPushButton[active="true"] {{ background: {tc("rgba(0,0,0,0.13)")}; }}
            QPushButton[active="true"]:hover {{ background: {tc("rgba(0,0,0,0.18)")}; }}
        """)

    def _update_date(self):
        from datetime import date
        d = date.today()
        self.date_label.setText(f"{d.month}. {d.day}.({'월화수목금토일'[d.weekday()]})")

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPos() - self.parent.frameGeometry().topLeft()
            self.parent._snapped_pos = None
            self.parent._current_snap_edges = []
            self.parent._snap_is_adj = False

    def mouseMoveEvent(self, e):
        if not (e.buttons() == Qt.LeftButton and self._drag_pos):
            return
        raw_pos = e.globalPos() - self._drag_pos
        p = self.parent
        new_pos, snap_edges, adj_snapped, x_snap_partner = p._snap_full(raw_pos)

        if p._snapped_pos is not None:
            threshold = BREAKAWAY_DIST if p._snap_is_adj else SNAP_RESIST
            dx = raw_pos.x() - p._snapped_pos.x()
            dy = raw_pos.y() - p._snapped_pos.y()
            if (dx*dx + dy*dy) ** 0.5 < threshold:
                if adj_snapped:
                    p._snapped_pos = new_pos
                    p._current_snap_edges = snap_edges
                    p._snap_is_adj = True
                else:
                    new_pos = p._snapped_pos
                    snap_edges = p._current_snap_edges
                    adj_snapped = p._snap_is_adj
            else:
                p._snapped_pos = None
                p._current_snap_edges = []
                p._snap_is_adj = False

        if p._snapped_pos is None:
            if adj_snapped:
                p._snapped_pos = new_pos
                p._current_snap_edges = snap_edges
                p._snap_is_adj = True
            elif snap_edges:
                p._snapped_pos = new_pos
                p._current_snap_edges = snap_edges
                p._snap_is_adj = False
            else:
                p._snapped_pos = None
                p._current_snap_edges = []

        p.move(new_pos)

        # 합체 버튼 노출 / 갱신 (좌우 모서리 + 상단 또는 하단 라인 일치)
        already_linked = p._link_group is not None and x_snap_partner in p._link_group
        has_overlap = False
        if x_snap_partner is not None and not already_linked:
            overlap_top    = max(p.y(), x_snap_partner.y())
            overlap_bottom = min(p.y() + p.height(), x_snap_partner.y() + x_snap_partner.height())
            has_overlap    = (overlap_bottom - overlap_top) >= 30
        if x_snap_partner is not None and not already_linked and has_overlap:
            left_win = p if p.x() <= x_snap_partner.x() else x_snap_partner
            right_win = x_snap_partner if left_win is p else p
            if left_win._merge_btn is None or left_win._merge_btn._win_b is not right_win:
                if left_win._merge_btn is not None and not left_win._merge_btn._merged:
                    left_win._merge_btn.hide()
                btn = MergeButton()
                btn.set_windows(left_win, right_win)
                left_win._merge_btn = btn
            btn = left_win._merge_btn
            btn.set_merged(False)
            btn.update_position(btn._win_a, btn._win_b)
            btn.show()
            btn._raise_timer.start()
        elif x_snap_partner is None:
            if p._merge_btn is not None and not p._merge_btn._merged:
                p._merge_btn.hide()
                p._merge_btn = None

        # 합체 상태: 드래그 중에는 합체 버튼을 항상 표시 (오토하이드는 유지)
        if p._link_group:
            for member in p._link_group:
                btn = member._merge_btn
                if btn and btn._merged:
                    if not btn.isVisible():
                        btn.show()
                        btn._raise_timer.start()
                    btn._visible = True
                    btn._hide_delay.stop()
                    btn._opacity_anim.stop()
                    btn._opacity_eff.setOpacity(1.0)
                    btn.update_position(btn._win_a, btn._win_b)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None
        self.parent._snapped_pos = None
        self.parent._current_snap_edges = []
        self.parent._snap_is_adj = False
        self.parent.save_state()
        if self.parent._link_group:
            for member in self.parent._link_group:
                if member is not self.parent:
                    member.save_state()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.parent.toggle_shade()


class UrgentToast(QWidget):
    """마감 임박 태스크를 우하단에 10초간 표시하는 토스트 팝업."""

    def __init__(self, tasks):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_ShowWithoutActivating)

        self.setStyleSheet(f"""
            QWidget#toast {{
                background: {tc("#fff0e6")};
            }}
            QLabel {{ background: transparent; border: none; }}
        """)
        self.setObjectName('toast')

        # 팝업 이미지 로드 (assets/popup_images/ 폴더에서 랜덤 선택)
        import random as _random
        _base = _base_path()
        _popup_dir = os.path.join(_base, 'assets', 'popup_images')
        _exts = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        _candidates = sorted([
            os.path.join(_popup_dir, f)
            for f in os.listdir(_popup_dir)
            if os.path.splitext(f)[1].lower() in _exts
        ]) if os.path.isdir(_popup_dir) else []
        _yuju_path = _random.choice(_candidates) if _candidates else None
        face_pix = None
        try:
            if not _yuju_path:
                raise FileNotFoundError
            from PIL import Image, ImageEnhance
            import io
            _pil = Image.open(_yuju_path).convert('RGBA')
            face_h_px = int(_pil.height * 0.72)
            _pil = _pil.crop((0, 0, _pil.width, face_h_px))
            _pil = ImageEnhance.Brightness(_pil).enhance(1.2)
            _pil = ImageEnhance.Color(_pil).enhance(1.1)
            buf = io.BytesIO()
            _pil.save(buf, format='PNG')
            _yuju_pix = QPixmap()
            _yuju_pix.loadFromData(buf.getvalue())
            face_pix = _yuju_pix.scaledToHeight(130, Qt.SmoothTransformation)
        except Exception:
            if _yuju_path:
                _yuju_pix = QPixmap(_yuju_path)
                if not _yuju_pix.isNull():
                    face_h = int(_yuju_pix.height() * 0.72)
                    _tmp = _yuju_pix.copy(0, 0, _yuju_pix.width(), face_h)
                    face_pix = _tmp.scaledToHeight(130, Qt.SmoothTransformation)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(27, 18, 27, 22)
        layout.setSpacing(11)

        header_row = QHBoxLayout()
        header_row.setSpacing(9)
        header = QLabel('마감 임박')
        header.setFont(QFont('Malgun Gothic', 16, QFont.Bold))
        header.setStyleSheet(f'color: {tct("#7a3000")};')
        header_row.addWidget(header)
        header_row.addStretch()
        layout.addLayout(header_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f'background: {tc("#e8a070")}; max-height: 2px; border: none;')
        layout.addWidget(sep)

        for task in tasks:
            dday = task.get('notice') or calc_dday(task['deadline'])
            name = task['name']
            if len(name) > 30:
                name = name[:29] + '…'
            row = QLabel(f'• {name}  <b>{dday}</b>')
            row.setFont(QFont('Malgun Gothic', 14))
            row.setStyleSheet(f'color: {tct("#3a2000")};')
            layout.addWidget(row)

        self.setMinimumWidth(300)
        self.adjustSize()
        screen = QApplication.primaryScreen().availableGeometry()
        tx = screen.right() - self.width() - 24
        ty = screen.bottom() - self.height() - 52
        self.move(tx, ty)

        # 유자 이미지: 팝업 우상단 밖으로 튀어나오는 별도 위젯
        self._face_widget = None
        if face_pix:
            fw = QLabel(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
            fw.setAttribute(Qt.WA_TranslucentBackground)
            fw.setAttribute(Qt.WA_ShowWithoutActivating)
            fw.setPixmap(face_pix)
            fw.setStyleSheet('background: transparent; border: none;')
            fw.resize(face_pix.size())
            # 팝업 가로 3/4 지점, 이미지 절반이 위로 튀어나오도록 배치
            fw.move(tx + self.width() * 3 // 4 - face_pix.width() // 2,
                    ty - face_pix.height() // 2)
            fw.show()
            fw.setCursor(Qt.PointingHandCursor)
            fw.mousePressEvent = lambda e: self._click_fade()
            self._face_widget = fw
            # 토스트가 show()된 뒤 이미지 위젯을 맨 앞으로
            QTimer.singleShot(0, fw.raise_)

        self.setCursor(Qt.PointingHandCursor)

        # 10초 후 페이드 아웃
        self._anim = QPropertyAnimation(self, b'windowOpacity')
        self._anim.setDuration(600)
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.setEasingCurve(QEasingCurve.InQuad)
        self._anim.finished.connect(self._on_fade_done)
        self._fade_timer = QTimer(self)
        self._fade_timer.setSingleShot(True)
        self._fade_timer.timeout.connect(self._anim.start)
        self._fade_timer.start(10000)

    def _on_fade_done(self):
        if self._face_widget:
            self._face_widget.close()
        self.close()

    def _click_fade(self):
        self._fade_timer.stop()
        self._anim.stop()
        self._anim.setDuration(600)
        self._anim.setStartValue(self.windowOpacity())
        self._anim.start()

    def paintEvent(self, event):
        from PyQt5.QtGui import QPainter, QColor
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        pen = painter.pen()
        pen.setColor(QColor(tc('#d4637a')))
        pen.setWidth(5)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        # 안쪽으로 2px 들어와서 테두리가 잘리지 않게
        painter.drawRect(2, 2, self.width() - 4, self.height() - 4)
        painter.end()

    def mousePressEvent(self, event):
        self._click_fade()


class SnapGuideOverlay(QWidget):
    """드래그 중 정렬 가이드라인을 전체 화면 위에 그리는 투명 오버레이."""

    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self._lines = []
        self._fade_timer = QTimer(self)
        self._fade_timer.setSingleShot(True)
        self._fade_timer.timeout.connect(self._start_fade)
        self._anim = QPropertyAnimation(self, b'windowOpacity')
        self._anim.setDuration(350)
        self._anim.finished.connect(self.hide)
        screens = QApplication.screens()
        total = screens[0].geometry()
        for s in screens[1:]:
            total = total.united(s.geometry())
        self.setGeometry(total)

    def show_lines(self, lines):
        self._lines = lines
        if not lines:
            return
        self._fade_timer.stop()
        self._anim.stop()
        self.setWindowOpacity(1.0)
        self.update()
        self.show()
        self.raise_()

    def hide_lines(self):
        self._lines = []
        self._fade_timer.start(250)

    def _start_fade(self):
        self._anim.setStartValue(self.windowOpacity())
        self._anim.setEndValue(0.0)
        self._anim.start()

    def paintEvent(self, event):
        if not self._lines:
            return
        painter = QPainter(self)
        pen = painter.pen()
        pen.setColor(QColor(41, 121, 255, 180))
        pen.setWidth(1)
        painter.setPen(pen)
        for x1, y1, x2, y2 in self._lines:
            painter.drawLine(x1, y1, x2, y2)
        painter.end()


_snap_guide_overlay = None


def _get_snap_guide_overlay():
    global _snap_guide_overlay
    if _snap_guide_overlay is None:
        _snap_guide_overlay = SnapGuideOverlay()
    return _snap_guide_overlay
