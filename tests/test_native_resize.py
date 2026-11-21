"""업무창 크기 조절을 Windows에 맡기기: WM_SIZING 스냅 변환 · 끝나면 저장 · Windows가 아니면 예전 방식."""
import os
import sys
import unittest.mock as um

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
_APP = None


@pytest.fixture
def win(tmp_path, monkeypatch):
    import db
    monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])
    import window
    w = window.MemoWindow(db.get_first_window_id())
    w.setGeometry(100, 100, 400, 420)
    w.show()
    _APP.processEvents()
    yield w
    w.close()


def test_not_windows_uses_old_drag(win):
    h = win._handles['left']
    assert h._start_native(win) is False and getattr(win, '_native_resize', None) is None


def test_sizing_rect_conversion_and_snap(win, monkeypatch):
    from ctypes import wintypes
    h = win._handles['right']
    win._native_resize = h
    win._native_origin = (200, 200, 100, 100, 2.0)            # 물리 (200,200) = 논리 (100,100), 배율 2
    seen = []

    def fake_snap(nx, ny, nw, nh, edge):
        seen.append((nx, ny, nw, nh, edge))
        return nx, ny, nw + 5, nh                                # 스냅으로 폭 5 늘어난 셈
    monkeypatch.setattr(win, '_snap_resize', fake_snap)
    r = wintypes.RECT(200, 200, 200 + 900, 200 + 840)            # 물리 900x840 = 논리 450x420
    win._native_sizing(r)
    assert seen == [(100, 100, 450, 420, 'right')]
    assert (r.left, r.top, r.right - r.left, r.bottom - r.top) == (200, 200, 910, 840)


def test_end_restores_min_and_saves(win, monkeypatch):
    saved = []
    monkeypatch.setattr(win, 'save_state', lambda: saved.append(1))
    h = win._handles['bottom']
    monkeypatch.setattr(win, 'winId', lambda: 0)
    win._begin_native_resize(h)
    assert win._native_resize is h and win.minimumWidth() >= 320
    win._end_native_resize()
    assert win._native_resize is None and saved == [1]
