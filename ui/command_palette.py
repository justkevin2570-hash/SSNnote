"""ui/command_palette.py — 빠른 찾기 (Ctrl+K).

키보드만으로 기능을 실행하고 업무를 찾는다. 초성 검색 지원 ('ㅇㅅ' → 예산 관리).
  ↑↓ 이동 · Enter 실행 · Esc 닫기

명령은 Command(제목, 실행 함수, 단축키, 검색어)로 넘겨받는다.
"""

from collections import namedtuple

from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtWidgets import (
    QApplication, QFrame, QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem,
    QVBoxLayout, QWidget,
)

from korean_search import best_score
from .common import pr_font
from .theme import T

Command = namedtuple('Command', 'title run shortcut keywords')
Command.__new__.__defaults__ = ('', '')

_MAX_TASKS = 8

_STYLE = f"""
    QFrame#Palette {{ background: {T.SURFACE}; border: 1px solid {T.SURFACE_WARM_HOVER}; border-radius: 10px; }}
    QLineEdit {{
        border: none; border-bottom: 1px solid {T.BORDER_WARM}; padding: 10px 14px;
        font-size: 12pt; color: {T.INK}; background: transparent;
    }}
    QListWidget {{ border: none; background: transparent; outline: 0; padding: 4px; }}
    QListWidget::item {{ border-radius: 6px; }}
    QListWidget::item:selected {{ background: {T.ACCENT_SOFT}; }}
    QLabel {{ background: transparent; }}
    QLabel#Title {{ color: {T.INK}; }}
    QLabel#Hint {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QLabel#Section {{ color: {T.TEXT_WARM_MUTED}; font-size: 8.5pt; font-weight: 700; padding: 6px 10px 2px; }}
    QLabel#Empty {{ color: {T.TEXT_WARM_MUTED}; padding: 14px; }}
"""


def rank(query: str, commands, tasks=()):
    """(명령 목록, 업무 목록)을 점수순으로. 검색어가 비면 명령은 원래 순서, 업무는 없음."""
    q = query.strip()
    if not q:
        return list(commands), []
    cmds = [(best_score(q, c.title, c.keywords), i, c) for i, c in enumerate(commands)]
    cmds = [c for s, i, c in sorted((x for x in cmds if x[0] > 0), key=lambda x: (-x[0], x[1]))]
    ts = [(best_score(q, t.get('name', '')), i, t) for i, t in enumerate(tasks)]
    ts = [t for s, i, t in sorted((x for x in ts if x[0] > 0), key=lambda x: (-x[0], x[1]))]
    return cmds, ts[:_MAX_TASKS]


class CommandPalette(QWidget):
    """on_task(task_dict): 업무 결과를 골랐을 때."""

    def __init__(self, commands, tasks=(), on_task=None, parent=None):
        super().__init__(parent, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setFont(pr_font(11))
        self.setStyleSheet(_STYLE)
        self._commands = list(commands)
        self._tasks = list(tasks)
        self._on_task = on_task
        self.setFixedWidth(460)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        frame = QFrame()
        frame.setObjectName('Palette')
        outer.addWidget(frame)
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.input = QLineEdit()
        self.input.setPlaceholderText('기능·업무 찾기 — 초성도 돼요 (ㅇㅅ → 예산)')
        lay.addWidget(self.input)
        self.list = QListWidget()
        self.list.setFocusPolicy(Qt.NoFocus)
        self.list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        lay.addWidget(self.list)
        foot = QLabel('↑↓ 이동 · Enter 실행 · Esc 닫기')
        foot.setObjectName('Hint')
        foot.setContentsMargins(14, 4, 14, 8)
        lay.addWidget(foot)

        self.input.textChanged.connect(self._refresh)
        self.input.installEventFilter(self)
        self.list.itemClicked.connect(self._activate)
        self._refresh()

    # ── 목록 ────────────────────────────────────────────────────
    def _add_section(self, text):
        item = QListWidgetItem()
        item.setFlags(Qt.NoItemFlags)
        lbl = QLabel(text)
        lbl.setObjectName('Section')
        item.setSizeHint(lbl.sizeHint())
        self.list.addItem(item)
        self.list.setItemWidget(item, lbl)

    def _add_row(self, title, hint, payload):
        item = QListWidgetItem()
        item.setData(Qt.UserRole, payload)
        w = QWidget()
        row = QHBoxLayout(w)
        row.setContentsMargins(10, 6, 10, 6)
        t = QLabel(title)
        t.setObjectName('Title')
        row.addWidget(t, 1)
        if hint:
            h = QLabel(hint)
            h.setObjectName('Hint')
            row.addWidget(h)
        item.setSizeHint(w.sizeHint())
        self.list.addItem(item)
        self.list.setItemWidget(item, w)

    def _refresh(self):
        self.list.clear()
        cmds, tasks = rank(self.input.text(), self._commands, self._tasks)
        if cmds:
            if tasks:
                self._add_section('기능')
            for c in cmds:
                self._add_row(c.title, c.shortcut, ('cmd', c))
        if tasks:
            self._add_section('업무')
            for t in tasks:
                hint = t.get('deadline', '')[:10].replace('-', '. ') if t.get('deadline') else ''
                self._add_row(('☑ ' if t.get('strikethrough') else '• ') + t.get('name', ''), hint, ('task', t))
        if not cmds and not tasks:
            item = QListWidgetItem()
            item.setFlags(Qt.NoItemFlags)
            lbl = QLabel('일치하는 기능이나 업무가 없어요.')
            lbl.setObjectName('Empty')
            item.setSizeHint(lbl.sizeHint())
            self.list.addItem(item)
            self.list.setItemWidget(item, lbl)
        self._select_first()
        # 최대 9줄 높이까지 늘어나고 그 뒤로는 스크롤
        h = sum(self.list.sizeHintForRow(i) for i in range(min(self.list.count(), 9)))
        self.list.setFixedHeight(max(h + 10, 48))
        self.adjustSize()

    def _selectable_rows(self):
        return [i for i in range(self.list.count()) if self.list.item(i).flags() & Qt.ItemIsSelectable]

    def _select_first(self):
        rows = self._selectable_rows()
        if rows:
            self.list.setCurrentRow(rows[0])

    def _move(self, step):
        rows = self._selectable_rows()
        if not rows:
            return
        cur = self.list.currentRow()
        idx = rows.index(cur) if cur in rows else -1
        self.list.setCurrentRow(rows[(idx + step) % len(rows)])

    def _activate(self, item=None):
        item = item or self.list.currentItem()
        if item is None or not (item.flags() & Qt.ItemIsSelectable):
            return
        kind, payload = item.data(Qt.UserRole)
        self.close()
        if kind == 'cmd':
            payload.run()
        elif self._on_task:
            self._on_task(payload)

    # ── 키보드 / 위치 ────────────────────────────────────────────
    def eventFilter(self, obj, event):
        if obj is self.input and event.type() == QEvent.KeyPress:
            k = event.key()
            if k == Qt.Key_Down:
                self._move(1)
                return True
            if k == Qt.Key_Up:
                self._move(-1)
                return True
            if k in (Qt.Key_Return, Qt.Key_Enter):
                self._activate()
                return True
            if k == Qt.Key_Escape:
                self.close()
                return True
        return super().eventFilter(obj, event)

    def changeEvent(self, event):
        # 다른 곳을 누르면 닫힌다 (팝업처럼)
        if event.type() == QEvent.ActivationChange and not self.isActiveWindow():
            self.close()
        super().changeEvent(event)

    def popup(self, anchor: QWidget = None):
        """anchor 창이 있는 화면의 위쪽 1/4 지점 가운데에 띄운다."""
        screen = (anchor.screen() if anchor is not None else None) or QApplication.primaryScreen()
        geo = screen.availableGeometry()
        self.adjustSize()
        self.move(geo.x() + (geo.width() - self.width()) // 2, geo.y() + geo.height() // 4)
        self.show()
        self.raise_()
        self.activateWindow()
        self.input.setFocus()
