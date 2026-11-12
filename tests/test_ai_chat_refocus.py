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
