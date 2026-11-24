"""캡처할 때 브라우저 잠깐 확대: 브라우저일 때만 · 끝나면(취소해도) 되돌림 · 설정으로 끄기."""
import sys

import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def bz(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    import browser_zoom
    sent, state = [], {'fg': 100, 'proc': {100: 'msedge.exe', 200: 'notepad.exe'}}
    monkeypatch.setattr(sys, 'platform', 'win32')
    monkeypatch.setattr(browser_zoom, '_foreground_window', lambda: state['fg'])
    monkeypatch.setattr(browser_zoom, '_top_other_window', lambda: state['fg'])
    monkeypatch.setattr(browser_zoom, '_process_name', lambda h: state['proc'].get(h, ''))
    monkeypatch.setattr(browser_zoom, '_set_foreground', lambda h: state.__setitem__('fg', h) or True)
    monkeypatch.setattr(browser_zoom, '_key_down', lambda vk: False)
    monkeypatch.setattr(browser_zoom, '_send_ctrl', lambda vk, n: sent.append((vk, n)))
    monkeypatch.setattr(browser_zoom, '_zoomed', None)
    return browser_zoom, sent, state


def test_zoom_and_restore_once(bz):
    b, sent, state = bz
    assert b.zoom_in() is True and sent == [(b._VK_OEM_PLUS, b.STEPS)]
    assert b.zoom_in() is False                               # 이미 키운 상태면 또 키우지 않음
    state['fg'] = 300                                         # 그 사이 캡처 화면이 앞에 있어도
    assert b.restore() is True and state['fg'] == 100         # 브라우저를 앞으로 가져와 되돌림
    assert sent == [(b._VK_OEM_PLUS, b.STEPS), (b._VK_OEM_MINUS, b.STEPS)]
    assert b.restore() is False and len(sent) == 2            # 한 번만


def test_not_a_browser_or_disabled(bz):
    b, sent, state = bz
    state['fg'] = 200
    assert b.zoom_in() is False and sent == []
    state['fg'] = 100
    b.set_enabled(False)
    assert not b.enabled() and b.zoom_in() is False and sent == []
    b.set_enabled(True)
    assert b.zoom_in() is True


def test_keys_still_held_means_no_zoom(bz, monkeypatch):
    b, sent, state = bz
    monkeypatch.setattr(b, '_key_down', lambda vk: True)
    monkeypatch.setattr(b, '_wait_keys_released', lambda timeout=0.8: False)
    assert b.zoom_in() is False and sent == []


def test_zoom_then_runs_callback(bz, monkeypatch):
    b, sent, state = bz
    global _APP
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    monkeypatch.setattr(b, 'RENDER_WAIT_MS', 0)
    done = []
    b.zoom_then(lambda: done.append(1))
    assert done == []                                         # 다시 그려질 시간을 둠
    _APP.processEvents(); _APP.processEvents()
    assert done == [1]
    b.restore()
    state['fg'] = 200
    b.zoom_then(lambda: done.append(2))                       # 브라우저가 아니면 바로
    assert done == [1, 2]


def test_overlay_restores_on_cancel(bz):
    b, sent, state = bz
    global _APP
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtGui import QPixmap
    _APP = QApplication.instance() or QApplication([])
    from capture import ScreenCaptureOverlay
    b.zoom_in()
    ov = ScreenCaptureOverlay(QPixmap(200, 200))
    got = []
    ov.cancelled.connect(lambda: got.append('cancel'))
    ov.close()
    assert got == ['cancel'] and sent[-1] == (b._VK_OEM_MINUS, b.STEPS) and b._zoomed is None
