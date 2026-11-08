"""ui/chrome.py — 메모 창 외곽 요소: 가장자리 핸들, 병합 버튼, 제목 표시줄, 스냅 가이드, 긴급 알림 토스트."""
import os
import math
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
from .theme import T, tc, tct, qcolor

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


# ── 마감 임박 알림 (유주) ─────────────────────────────────────────

def toast_level(task) -> tuple:
    """알림 줄 하나의 급한 정도와 표시 글자. ('now' | 'soon' | 'later', '2시간 후' 등)

    now   = 몇 시간 안 · 오늘 마감 / soon = D-1·D-2 / later = 그 밖 (D-3~5)
    """
    notice = task.get('notice') or ''
    dday = calc_dday(task.get('deadline', ''))
    label = notice.replace(' 마감', '') if notice else dday
    if notice or '시간' in dday or '분' in dday or dday == 'D-day':
        return 'now', label
    if dday.startswith('D-'):
        try:
            return ('soon' if int(dday[2:]) <= 2 else 'later'), label
        except ValueError:
            pass
    return 'later', label


# 급한 정도별 유주 표정 (assets/popup_images/yuju_N.png)
_YUJU_FOR_LEVEL = {'now': (4,), 'soon': (5,), 'later': (1, 2, 3, 6)}
_YUJU_LINE = {
    'now': '곧 마감이에요! 지금 한번 볼까요?',
    'soon': '마감이 얼마 남지 않았어요.',
    'later': '이번 주 마감, 미리 챙겨 둬요.',
}
_LEVEL_COLORS = {          # (글자, 바탕) — 라이트 기준, 다크는 tc()로 변환
    'now':   ('#D6452F', '#FDEBE7'),
    'soon':  ('#C2600C', '#FEF0E1'),
    'later': ('#8C6D00', '#FAF2D3'),
}
_LEVEL_RANK = {'now': 0, 'soon': 1, 'later': 2}


def _ease_out_cubic(x):
    return 1 - (1 - x) ** 3


def _ease_in_cubic(x):
    return x ** 3


def _ease_out_back(x, s=1.9):
    x -= 1
    return 1 + (s + 1) * x ** 3 + s * x ** 2


class UrgentToast(QWidget):
    """마감 임박 업무를 화면 오른쪽 아래에 잠시 보여주는 알림. 유주가 카드 위로 고개를 내민다.

    움직임: 카드가 아래에서 살며시 올라오고, 유주가 카드 뒤에서 통 튀어 오른 뒤 가볍게 숨쉬듯 흔들린다.
    마우스를 올리면 멈춰 기다리고(아래 시간 막대도 멈춤), 클릭하거나 ×를 누르면 내려가며 사라진다.
    급함 표시 설정에서 움직임을 끄면 흔들림 없이 나타나고 사라지기만 한다.
    """

    SHOW_SECONDS = 10.0
    MAX_ROWS = 4

    def __init__(self, tasks):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMouseTracking(True)
        self.setCursor(Qt.PointingHandCursor)

        from .common import pr_font
        rows = sorted(((toast_level(t), t) for t in tasks), key=lambda x: _LEVEL_RANK[x[0][0]])
        self._rows = [(lvl, label, t.get('name', '')) for (lvl, label), t in rows]
        self._level = self._rows[0][0] if self._rows else 'later'
        self._more = max(0, len(self._rows) - self.MAX_ROWS)
        self._rows = self._rows[:self.MAX_ROWS]

        try:
            from . import motion
            self._moving = motion.load()['enabled']
        except Exception:
            self._moving = True

        # ── 글꼴 ──
        self._f_title = pr_font(12); self._f_title.setWeight(QFont.Bold)
        self._f_line = pr_font(9)
        self._f_row = pr_font(10)
        self._f_chip = pr_font(9); self._f_chip.setWeight(QFont.DemiBold)
        self._f_more = pr_font(9)

        # ── 치수 (글꼴 배율 반영) ──
        self._shadow = px(18)                 # 그림자 여백
        self._radius = px(16)
        self._pad = px(16)
        self._row_h = px(30)
        card_w = px(340)
        card_h = (self._pad + px(22) + px(18) + px(10) + self._row_h * len(self._rows)
                  + (px(20) if self._more else 0) + px(14))
        card_h = max(card_h, px(112))
        s = self._shadow
        # 유주: 카드 상단 바깥쪽에 서 있으므로, 창 위쪽 여백을 유주 키만큼 잡는다
        self._yuju = self._load_yuju(self._level, max_w=px(150), max_h=px(140))
        self._overhang = int(self._yuju_h())
        self.resize(card_w + 2 * s, card_h + s * 2 + self._overhang)
        self._card = QRectF(s, s + self._overhang, card_w, card_h)

        # ── 위치: 오른쪽 아래 ──
        screen = QApplication.primaryScreen().availableGeometry()
        self._base = QPoint(screen.right() - self.width() + s - px(14),
                            screen.bottom() - self.height() + s - px(14))
        self.move(self._base.x(), self._base.y() + px(16))
        self.setWindowOpacity(0.0)

        # ── 시간 ──
        import time as _time
        self._now = _time.monotonic
        self._t0 = None
        self._leave_t = None
        self._shown_for = 0.0                 # 마우스가 올라가 있지 않은 동안 흐른 시간
        self._last = None
        self._hover = False
        self._hover_close = False
        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._tick)
        # 그리는 값
        self._yuju_dy = 0.0
        self._yuju_angle = 0.0

    # ── 준비 ──
    def _yuju_h(self):
        """유주 그림의 높이 (논리 픽셀). 그림이 없으면 0."""
        if not self._yuju:
            return 0
        return self._yuju.height() / self._yuju.devicePixelRatio()

    @staticmethod
    def _load_yuju(level, max_w, max_h):
        import random as _random
        folder = os.path.join(_base_path(), 'assets', 'popup_images')
        choices = [os.path.join(folder, f'yuju_{n}.png') for n in _YUJU_FOR_LEVEL.get(level, ())]
        choices = [p for p in choices if os.path.exists(p)]
        if not choices and os.path.isdir(folder):
            choices = [os.path.join(folder, f) for f in sorted(os.listdir(folder)) if f.lower().endswith('.png')]
        if not choices:
            return None
        pix = QPixmap(_random.choice(choices))
        if pix.isNull():
            return None
        dpr = QApplication.primaryScreen().devicePixelRatio() if QApplication.primaryScreen() else 1.0
        height = min(max_h, max_w * pix.height() / pix.width())   # 유주 자리 폭을 넘지 않게
        pix = pix.scaledToHeight(int(height * dpr), Qt.SmoothTransformation)
        pix.setDevicePixelRatio(dpr)
        return pix

    def showEvent(self, event):
        super().showEvent(event)
        if self._t0 is None:
            self._t0 = self._last = self._now()
            self._timer.start()

    # ── 움직임 ──
    def _tick(self):
        now = self._now()
        t = now - self._t0
        if not self._hover and self._leave_t is None:
            self._shown_for += now - self._last
        self._last = now

        # 들어오기: 살며시 올라오며 나타남
        e = _ease_out_cubic(min(1.0, t / 0.38))
        dy = (1 - e) * px(16)
        opacity = e
        # 유주: 카드 뒤에서 통 튀어 오름 → 가볍게 숨쉬기
        p = max(0.0, min(1.0, (t - 0.12) / 0.6))
        pop = _ease_out_back(p) if self._moving else _ease_out_cubic(p)
        yuju_h = self._yuju.height() / self._yuju.devicePixelRatio() if self._yuju else 0
        self._yuju_dy = (1 - pop) * yuju_h * 0.6
        self._yuju_angle = 0.0
        if self._moving and p >= 1.0:
            idle = t - 0.72
            self._yuju_dy += math.sin(2 * math.pi * idle / 2.8) * px(2)
            self._yuju_angle = math.sin(2 * math.pi * idle / 3.6) * 1.8
            if self._level == 'now' and idle < 0.9:          # 급할 땐 처음에 한 번 꼼지락
                self._yuju_angle += math.sin(idle * 2 * math.pi * 2.2) * 7 * (1 - idle / 0.9)

        if self._leave_t is None and self._shown_for >= self.SHOW_SECONDS:
            self.dismiss()
        if self._leave_t is not None:
            l = _ease_in_cubic(min(1.0, (now - self._leave_t) / 0.24))
            opacity = min(opacity, 1 - l)
            dy += l * px(12)
            self._yuju_dy += l * px(18)
            if l >= 1.0:
                self._timer.stop()
                self.close()
                return
        self.setWindowOpacity(opacity)
        self.move(self._base.x(), self._base.y() + int(round(dy)))
        self.update()

    def dismiss(self):
        if self._leave_t is None:
            self._leave_t = self._now()

    # ── 마우스 ──
    def _close_rect(self):
        c = self._card
        size = px(22)
        return QRectF(c.right() - self._pad + px(4) - size, c.top() + px(12), size, size)

    def enterEvent(self, event):
        self._hover = True
        self.update()

    def leaveEvent(self, event):
        self._hover = False
        self._hover_close = False
        self.update()

    def mouseMoveEvent(self, event):
        over = self._close_rect().adjusted(-4, -4, 4, 4).contains(event.pos())
        if over != self._hover_close:
            self._hover_close = over
            self.update()

    def mousePressEvent(self, event):
        self.dismiss()

    # ── 그리기 ──
    def paintEvent(self, event):
        from PyQt5.QtGui import QPainterPath, QPen
        from PyQt5.QtCore import QPointF
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        c, r = self._card, self._radius
        fg, soft = (tc(x) for x in _LEVEL_COLORS[self._level])

        # 그림자 (겹겹이 옅게)
        p.setPen(Qt.NoPen)
        steps = 9
        for i in range(steps, 0, -1):
            grow = self._shadow * i / steps
            alpha = int(16 * (1 - i / (steps + 1)) ** 1.6) + 1
            p.setBrush(qcolor('#3a2a10', alpha) if not _dark() else QColor(0, 0, 0, alpha * 3))
            p.drawRoundedRect(c.adjusted(-grow * 0.7, -grow * 0.4, grow * 0.7, grow * 1.0), r + grow, r + grow)

        # 카드
        card_path = QPainterPath()
        card_path.addRoundedRect(c, r, r)
        p.setBrush(QColor(tc('#ffffff')))
        p.setPen(QPen(QColor(tc('#ece5da') if not self._hover else tc('#e2d6c4')), 1))
        p.drawPath(card_path)



        # 글자 영역
        x0 = c.left() + self._pad + px(2)     # 유주는 카드 위에 서므로 글자는 왼쪽 끝까지
        x1 = c.right() - self._pad
        y = c.top() + self._pad

        # 제목 줄: ● 마감 임박  n건            ×
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(fg))
        p.drawEllipse(QPointF(x0 + px(4), y + px(11)), px(4), px(4))
        p.setFont(self._f_title)
        p.setPen(QColor(tct('#1f1a14')))
        title_rect = QRectF(x0 + px(14), y, x1 - x0, px(22))
        p.drawText(title_rect, Qt.AlignLeft | Qt.AlignVCenter, '마감 임박')
        tw = p.fontMetrics().horizontalAdvance('마감 임박')
        p.setFont(self._f_line)
        p.setPen(QColor(tct('#9a8f82')))
        total = len(self._rows) + self._more
        p.drawText(QRectF(title_rect.left() + tw + px(6), y, px(60), px(22)),
                   Qt.AlignLeft | Qt.AlignVCenter, f'{total}건')

        cr = self._close_rect()
        if self._hover_close:
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(tc('#f3eee7')))
            p.drawEllipse(cr)
        pen = QPen(QColor(tct('#6d6358') if self._hover_close else tct('#b3a99c')), 1.6)
        pen.setCapStyle(Qt.RoundCap)
        p.setPen(pen)
        k = cr.width() * 0.22
        ctr = cr.center()
        p.drawLine(QPointF(ctr.x() - k, ctr.y() - k), QPointF(ctr.x() + k, ctr.y() + k))
        p.drawLine(QPointF(ctr.x() + k, ctr.y() - k), QPointF(ctr.x() - k, ctr.y() + k))
        y += px(22)

        # 유주 한마디
        p.setFont(self._f_line)
        p.setPen(QColor(tct('#8a7f73')))
        p.drawText(QRectF(x0, y, x1 - x0, px(18)), Qt.AlignLeft | Qt.AlignVCenter, _YUJU_LINE[self._level])
        y += px(18) + px(10)

        # 업무 줄: 이름 ……… [D-day]
        for lvl, label, name in self._rows:
            lfg, lsoft = (tc(x) for x in _LEVEL_COLORS[lvl])
            p.setFont(self._f_chip)
            fm = p.fontMetrics()
            chip_w = fm.horizontalAdvance(label) + px(14)
            chip_h = px(20)
            chip = QRectF(x1 - chip_w, y + (self._row_h - chip_h) / 2, chip_w, chip_h)
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(lsoft))
            p.drawRoundedRect(chip, chip_h / 2, chip_h / 2)
            p.setPen(QColor(lfg))
            p.drawText(chip, Qt.AlignCenter, label)
            p.setFont(self._f_row)
            fm = p.fontMetrics()
            name_w = chip.left() - x0 - px(10)
            p.setPen(QColor(tct('#2b241c')))
            p.drawText(QRectF(x0, y, name_w, self._row_h), Qt.AlignLeft | Qt.AlignVCenter,
                       fm.elidedText(name, Qt.ElideRight, int(name_w)))
            y += self._row_h
        if self._more:
            p.setFont(self._f_more)
            p.setPen(QColor(tct('#9a8f82')))
            p.drawText(QRectF(x0, y, x1 - x0, px(20)), Qt.AlignLeft | Qt.AlignVCenter, f'외 {self._more}건 더')

        # 유주: 카드 상단 바깥쪽에 서 있음. 앞 층이라 글자 위에 그려지지만, 글자 자리(카드 안 위쪽 여백)와는 겹치지 않음
        if self._yuju:
            dpr = self._yuju.devicePixelRatio()
            w = self._yuju.width() / dpr
            h = self._yuju.height() / dpr
            p.save()
            foot = QPointF(c.center().x() + px(2), c.top() + px(4) + self._yuju_dy)
            p.translate(foot)
            p.rotate(self._yuju_angle)
            p.drawPixmap(QRectF(-w / 2, -h, w, h), self._yuju, QRectF(self._yuju.rect()))
            p.restore()

        # 남은 시간 막대 (카드 아래 가장자리, 마우스를 올리면 멈춤)
        remain = max(0.0, 1 - self._shown_for / self.SHOW_SECONDS)
        if remain > 0:
            p.save()
            p.setClipPath(card_path)
            bar = QColor(fg)
            bar.setAlpha(110 if self._hover else 70)
            p.setPen(Qt.NoPen)
            p.setBrush(bar)
            left = c.left() + self._pad
            p.drawRect(QRectF(left, c.bottom() - px(3), (c.right() - left) * remain, px(3)))
            p.restore()
        p.end()


def _dark() -> bool:
    from .theme import DARK
    return DARK


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
