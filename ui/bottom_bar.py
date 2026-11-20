"""업무창 아래 줄 버튼 순서와 Alt+1~4 단축키.

- 순서: db 설정 'bottom_order' (예: 'doc,budget,calendar,ai'). 기본은 공문 · 예산 · 달력 · AI.
- Alt+숫자: 왼쪽부터 Alt+1 · 2 · 3 · 4 — 버튼 자리를 따라간다. 앱 창이 눌려 있을 때만 (AltNumberFilter).
- BottomOrderDialog: ▲▼로 순서 바꾸기, 바꾸는 즉시 열려 있는 업무창 모두에 반영.
"""
from PyQt5.QtCore import QEvent, QObject, Qt
from PyQt5.QtWidgets import QApplication, QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

ORDER_KEY = 'bottom_order'
DEFAULT_ORDER = ('doc', 'budget', 'calendar', 'ai')
LABELS = {'doc': '공문', 'budget': '예산', 'calendar': '달력', 'ai': 'AI 대화'}
_KEYS = {Qt.Key_1: 1, Qt.Key_2: 2, Qt.Key_3: 3, Qt.Key_4: 4}


def load_order() -> tuple:
    try:
        from db import get_setting
        raw = get_setting(ORDER_KEY, '')
    except Exception:
        raw = ''
    keys = [k for k in raw.split(',') if k in DEFAULT_ORDER]
    if sorted(keys) != sorted(DEFAULT_ORDER):          # 비었거나 망가졌으면 기본 순서
        return DEFAULT_ORDER
    return tuple(keys)


def save_order(order):
    from db import set_setting
    order = tuple(order)
    if sorted(order) != sorted(DEFAULT_ORDER):
        order = DEFAULT_ORDER
    set_setting(ORDER_KEY, ','.join(order))
    apply_everywhere()


def shortcut_text(key, order=None) -> str:
    order = order or load_order()
    return f'Alt+{order.index(key) + 1}'


def memo_windows():
    from window import MemoWindow
    return [w for w in QApplication.topLevelWidgets() if isinstance(w, MemoWindow)]


def apply_everywhere():
    """열려 있는 업무창 모두의 아래 줄 순서 · 툴팁을 다시 맞춘다."""
    for w in memo_windows():
        try:
            w._apply_bottom_order()
        except (RuntimeError, AttributeError):
            pass


class AltNumberFilter(QObject):
    """앱 창이 눌려 있을 때 Alt+1~4 → 아래 줄 그 자리의 버튼.
    get_target(): 버튼을 누를 업무창 (마지막으로 쓴 업무창), enabled(): 「단축키 사용」 설정."""

    def __init__(self, get_target, enabled=lambda: True, parent=None):
        super().__init__(parent)
        self._get_target = get_target
        self._enabled = enabled

    def eventFilter(self, obj, event):
        if event.type() != QEvent.KeyPress or event.isAutoRepeat():
            return False
        n = _KEYS.get(event.key())
        mods = event.modifiers() & (Qt.AltModifier | Qt.ControlModifier | Qt.ShiftModifier | Qt.MetaModifier)
        if n is None or mods != Qt.AltModifier or not self._enabled():
            return False
        if QApplication.activeWindow() is None:        # 앱 창이 눌려 있을 때만
            return False
        target = self._get_target()
        if target is None:
            return False
        target.trigger_bottom(n)
        return True


class BottomOrderDialog(QDialog):
    """아래 줄 버튼 순서 — ▲▼로 바꾸면 바로 반영된다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        from ui.common import px
        self.setWindowTitle('아래 줄 버튼 순서')
        self.setStyleSheet("""
            QDialog { background: #F5F6F8; }
            QLabel#Title { font-size: 13pt; font-weight: 700; color: #1E293B; }
            QLabel#Hint { color: #64748B; }
            QFrame, QLabel { color: #1E293B; }
            QLabel#Key { color: #4F46E5; font-weight: 700; }
            QLabel#Name { font-weight: 600; }
            QPushButton { background: #FFFFFF; border: 1px solid #E3E6EB; border-radius: 6px; padding: 4px 10px; color: #475569; }
            QPushButton:hover { background: #F1F3F6; }
            QPushButton:disabled { color: #CBD5E1; }
            QPushButton#Primary { background: #4F46E5; border-color: #4F46E5; color: white; font-weight: 700; }
        """)
        self.setFixedWidth(px(300))
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 16)
        v.setSpacing(8)
        t = QLabel('아래 줄 버튼 순서')
        t.setObjectName('Title')
        v.addWidget(t)
        hint = QLabel('왼쪽부터 Alt+1 · 2 · 3 · 4')
        hint.setObjectName('Hint')
        v.addWidget(hint)
        self._rows = []
        for i in range(len(DEFAULT_ORDER)):
            h = QHBoxLayout()
            h.setSpacing(6)
            key = QLabel(f'Alt+{i + 1}')
            key.setObjectName('Key')
            key.setFixedWidth(px(48))
            name = QLabel()
            name.setObjectName('Name')
            up, down = QPushButton('▲'), QPushButton('▼')
            for b in (up, down):
                b.setFixedWidth(px(34))
                b.setCursor(Qt.PointingHandCursor)
            up.clicked.connect(lambda _, i=i: self.move(i, -1))
            down.clicked.connect(lambda _, i=i: self.move(i, 1))
            h.addWidget(key)
            h.addWidget(name, 1)
            h.addWidget(up)
            h.addWidget(down)
            v.addLayout(h)
            self._rows.append((name, up, down))
        v.addSpacing(6)
        bottom = QHBoxLayout()
        self.btn_reset = QPushButton('기본 순서로')
        self.btn_reset.clicked.connect(self.reset)
        close = QPushButton('닫기')
        close.setObjectName('Primary')
        close.clicked.connect(self.accept)
        bottom.addWidget(self.btn_reset)
        bottom.addStretch(1)
        bottom.addWidget(close)
        v.addLayout(bottom)
        self.order = list(load_order())
        self._show()

    def _show(self):
        last = len(self.order) - 1
        for i, (name, up, down) in enumerate(self._rows):
            name.setText(LABELS[self.order[i]])
            up.setEnabled(i > 0)
            down.setEnabled(i < last)
        self.btn_reset.setEnabled(tuple(self.order) != DEFAULT_ORDER)

    def move(self, i, d):
        j = i + d
        if 0 <= j < len(self.order):
            self.order[i], self.order[j] = self.order[j], self.order[i]
            save_order(self.order)
            self._show()

    def reset(self):
        self.order = list(DEFAULT_ORDER)
        save_order(self.order)
        self._show()


def open_order_dialog(parent=None):
    dlg = BottomOrderDialog(parent)
    dlg.exec_()
    return dlg
