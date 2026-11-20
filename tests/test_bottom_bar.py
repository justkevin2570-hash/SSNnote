"""업무창 아래 줄 버튼 순서(기본 공문 · 예산 · 달력 · AI) · Alt+1~4 · 순서 바꾸기 창."""
import sys

import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def win(tmp_path, monkeypatch):
    from unittest.mock import MagicMock
    sys.modules.setdefault('winreg', MagicMock())
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    monkeypatch.setenv('APPDATA', str(tmp_path))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = app = QApplication.instance() or QApplication([])
    import window
    w = window.MemoWindow(db.get_first_window_id())
    w.resize(420, 420)
    w.show()
    app.processEvents()
    calls = []
    for key in ('doc', 'budget', 'calendar'):
        b = w._bottom_btns[key]
        b.clicked.disconnect()
        b.clicked.connect(lambda _=False, k=key: calls.append(k))
    yield w, app, calls
    w.close()


def _names(w):
    lay = w._bottom_layout
    rev = {b: k for k, b in w._bottom_btns.items()}
    return [rev[lay.itemAt(i).widget()] for i in range(lay.count()) if lay.itemAt(i).widget() in rev]


def test_default_order_and_tooltips(win):
    w, app, _ = win
    assert _names(w) == ['doc', 'budget', 'calendar', 'ai']
    assert w._bottom_btns['doc'].toolTip().endswith('(Alt+1)')
    assert w._bottom_btns['calendar'].toolTip().endswith('(Alt+3)')


def test_alt_numbers_follow_position(win):
    w, app, calls = win
    from PyQt5.QtCore import Qt, QEvent
    from PyQt5.QtGui import QKeyEvent
    from ui import bottom_bar
    f = bottom_bar.AltNumberFilter(lambda: w)
    w.activateWindow()
    app.processEvents()
    import PyQt5.QtWidgets as qw
    orig = qw.QApplication.activeWindow
    qw.QApplication.activeWindow = staticmethod(lambda: w)
    try:
        ev = QKeyEvent(QEvent.KeyPress, Qt.Key_2, Qt.AltModifier)
        assert f.eventFilter(w, ev) is True and calls == ['budget']
        assert f.eventFilter(w, QKeyEvent(QEvent.KeyPress, Qt.Key_2, Qt.NoModifier)) is False
        assert f.eventFilter(w, QKeyEvent(QEvent.KeyPress, Qt.Key_2, Qt.AltModifier | Qt.ShiftModifier)) is False
        bottom_bar.save_order(('calendar', 'doc', 'budget', 'ai'))
        assert _names(w) == ['calendar', 'doc', 'budget', 'ai']
        f.eventFilter(w, QKeyEvent(QEvent.KeyPress, Qt.Key_1, Qt.AltModifier))
        assert calls == ['budget', 'calendar']
        assert w._bottom_btns['calendar'].toolTip().endswith('(Alt+1)')
        qw.QApplication.activeWindow = staticmethod(lambda: None)        # 앱 창이 안 눌려 있으면 무시
        assert f.eventFilter(w, QKeyEvent(QEvent.KeyPress, Qt.Key_1, Qt.AltModifier)) is False
        off = bottom_bar.AltNumberFilter(lambda: w, enabled=lambda: False)
        qw.QApplication.activeWindow = staticmethod(lambda: w)
        assert off.eventFilter(w, QKeyEvent(QEvent.KeyPress, Qt.Key_1, Qt.AltModifier)) is False
    finally:
        qw.QApplication.activeWindow = orig


def test_order_dialog_moves_and_resets(win):
    w, app, _ = win
    from ui import bottom_bar
    dlg = bottom_bar.BottomOrderDialog()
    assert [r[0].text() for r in dlg._rows] == ['공문', '예산', '달력', 'AI 대화']
    assert not dlg._rows[0][1].isEnabled() and not dlg._rows[3][2].isEnabled() and not dlg.btn_reset.isEnabled()
    dlg._rows[2][1].click()                                    # 달력 ▲
    assert bottom_bar.load_order() == ('doc', 'calendar', 'budget', 'ai')
    assert _names(w) == ['doc', 'calendar', 'budget', 'ai'] and dlg.btn_reset.isEnabled()
    dlg.btn_reset.click()
    assert bottom_bar.load_order() == bottom_bar.DEFAULT_ORDER and _names(w)[0] == 'doc'


def test_broken_setting_falls_back(win):
    import db
    from ui import bottom_bar
    db.set_setting(bottom_bar.ORDER_KEY, 'doc,doc,xx')
    assert bottom_bar.load_order() == bottom_bar.DEFAULT_ORDER
