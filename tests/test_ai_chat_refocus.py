"""다른 앱에 갔다 돌아와도 AI 대화 창이 꺼지지 않는다."""
import os
import sys
import unittest.mock as um

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.modules.setdefault('winreg', um.MagicMock())

import pytest
from PyQt5.QtGui import QHideEvent
from PyQt5.QtWidgets import QApplication


@pytest.fixture
def win(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 's.db'))
    db.init_db()
    app = QApplication.instance() or QApplication([])
    import window
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, 'auto_open', lambda h: None)
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    w._open_windows.append(w)
    w.resize(340, 600)
    w.show()
    app.processEvents()
    w._toggle_ai_chat()
    app.processEvents()
    yield w, acw.instance(), app
    acw.host_closed(w)
    w.close()


def _spontaneous_hide():
    class _E(QHideEvent):
        def spontaneous(self):
            return True
    return _E()


def test_spontaneous_hide_keeps_chat(win):
    w, c, app = win
    assert c.is_shown()
    w.hideEvent(_spontaneous_hide())          # 윈도우가 숨김 (다른 앱 · 바탕화면 보기)
    assert c.is_shown()


def test_app_hide_still_hides_chat(win):
    w, c, app = win
    w.toggle_shade()                          # 제목만 남게 접기 → 함께 숨김
    app.processEvents()
    assert not c.is_shown()
    w.toggle_shade()
    app.processEvents()
    assert c.is_shown()


def test_activation_reattaches_open_chat(win, monkeypatch):
    w, c, app = win
    c.hide_chat()                             # 어떤 까닭으로 꺼졌어도 (열림 설정은 그대로)
    assert not c.is_shown()
    from PyQt5.QtCore import QEvent
    monkeypatch.setattr(type(w), 'isActiveWindow', lambda self: True)
    w.changeEvent(QEvent(QEvent.ActivationChange))
    app.processEvents()
    assert c.is_shown()


def test_user_closed_chat_stays_closed(win, monkeypatch):
    w, c, app = win
    c.hide_by_user()
    from PyQt5.QtCore import QEvent
    monkeypatch.setattr(type(w), 'isActiveWindow', lambda self: True)
    w.changeEvent(QEvent(QEvent.ActivationChange))
    assert not c.is_shown()


def _fresh(monkeypatch, tmp_path):
    """앱을 새로 켠 것처럼: AI 창 없음 · 열림 설정은 켜짐."""
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 's.db'))
    db.init_db()
    QApplication.instance() or QApplication([])
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, '_win', None)
    return db, acw


def test_restart_opens_even_if_remembered_window_is_now_memo(tmp_path, monkeypatch):
    db, acw = _fresh(monkeypatch, tmp_path)
    import window
    first = db.get_first_window_id()
    other = db.create_window()
    db.set_setting(acw.HOST_KEY, str(other))          # 기억한 창이 다시 켜니 메모장이 됨
    wins = []
    a = window.MemoWindow(first, open_windows=wins)
    b = window.MemoWindow(other, open_windows=wins, force_memo_mode=True)
    wins += [a, b]
    a.show(), b.show()
    try:
        a._auto_open_ai_chat()
        b._auto_open_ai_chat()
        c = acw.instance(create=False)
        assert c is not None and c.host is a and c.is_shown()
    finally:
        acw.host_closed(a)
        a.close(), b.close()


def test_restart_opens_when_window_shows_late(tmp_path, monkeypatch):
    db, acw = _fresh(monkeypatch, tmp_path)
    import window
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    w._open_windows.append(w)
    try:
        w._auto_open_ai_chat()                         # 아직 안 보임 → 못 붙음
        assert acw.instance(create=False) is None
        w.show()
        QApplication.processEvents()
        w.hide(), w.show()                             # 나중에 보이면 그때 붙음
        c = acw.instance(create=False)
        assert c is not None and c.host is w and c.is_shown()
    finally:
        acw.host_closed(w)
        w.close()


def test_restart_respects_user_closed(tmp_path, monkeypatch):
    db, acw = _fresh(monkeypatch, tmp_path)
    import window
    db.set_setting(acw.OPEN_KEY, '0')
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    w._open_windows.append(w)
    w.show()
    try:
        w._auto_open_ai_chat()
        assert acw.instance(create=False) is None
    finally:
        w.close()


def test_close_with_x_then_reopen_from_tray_keeps_chat(tmp_path, monkeypatch):
    """제목줄 ✕로 업무창을 닫고(앱은 트레이에 남음) 다시 열면 AI 대화도 다시 열린다."""
    db, acw = _fresh(monkeypatch, tmp_path)
    import window
    wid = db.get_first_window_id()
    a = window.MemoWindow(wid, open_windows=[])
    a._open_windows.append(a)
    a.show()
    a._auto_open_ai_chat()
    c = acw.instance(create=False)
    assert c is not None and c.is_shown()
    a.close()                                          # ✕: AI 창은 떼어 내 숨김 (열림 설정은 그대로)
    assert c.host is None and not c.is_shown()
    b = window.MemoWindow(wid, open_windows=[])        # 트레이에서 다시 열기
    b._open_windows.append(b)
    b.show()
    try:
        b._auto_open_ai_chat()
        assert c.host is b and c.is_shown()
    finally:
        acw.host_closed(b)
        b.close()


def test_outside_chat_and_host_come_forward_together(win, monkeypatch):
    """옆에 붙은 AI 창 ↔ 업무창: 한쪽을 누르면 다른 쪽도 함께 앞으로."""
    w, c, app = win
    from PyQt5.QtCore import QEvent
    c.set_side('right')
    raised = []
    monkeypatch.setattr(type(c), 'raise_', lambda self: raised.append('chat'))
    monkeypatch.setattr(type(w), 'raise_', lambda self: raised.append('host'))
    monkeypatch.setattr(type(w), 'isActiveWindow', lambda self: True)
    w.changeEvent(QEvent(QEvent.ActivationChange))            # 업무창을 누름
    assert 'chat' in raised
    raised.clear()
    monkeypatch.setattr(type(c), 'isActiveWindow', lambda self: True)
    c.changeEvent(QEvent(QEvent.ActivationChange))            # AI 창을 누름
    assert raised[:1] == ['host']
    c.set_side('inside')                                       # 안에 넣었으면 할 일 없음
    raised.clear()
    w.changeEvent(QEvent(QEvent.ActivationChange))
    assert raised == []
