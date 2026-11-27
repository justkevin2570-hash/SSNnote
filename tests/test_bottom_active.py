"""아래 줄 공문 · 예산 · 달력 단추: 그 창이 열려 있으면 눌린 모양, 달력을 고정하면 보라 핀 · 📌 단추는 벡터 아이콘."""
import sys

import pytest

pytest.importorskip('PyQt5')


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
    app = QApplication.instance() or QApplication([])
    import window
    w = window.MemoWindow(db.get_first_window_id())
    w._open_windows = [w]
    w.resize(420, 460)
    w.show()
    app.processEvents()
    yield w, app
    dlg = getattr(w, '_calendar_dlg', None)
    if dlg is not None:
        try:
            dlg.close()
        except RuntimeError:
            pass
    w.close()


def _pump(app, n=5):
    for _ in range(n):
        app.processEvents()


def test_calendar_button_active_and_pin_badge(win):
    w, app = win
    b = w._bottom_btns['calendar']
    w._refresh_bottom_active()
    assert not b.property('active') and not w._cal_pin_badge.isVisible()
    w._show_monthly_calendar()
    _pump(app)
    dlg = w._calendar_dlg
    w._refresh_bottom_active()
    assert b.property('active')                                      # 열려 있으면 눌린 모양
    assert not w._cal_pin_badge.isVisible()
    dlg._apply_pin('top')                                            # 고정 → 보라 핀 바로
    _pump(app)
    assert w._cal_pin_badge.isVisible()
    assert w._cal_pin_badge.geometry().right() <= b.width()          # 단추 안 오른쪽 위
    dlg.showMinimized()
    _pump(app)
    w._refresh_bottom_active()
    assert not b.property('active') and not w._cal_pin_badge.isVisible()   # 내려 두면 원래 모양
    dlg.showNormal()
    _pump(app)
    w._refresh_bottom_active()
    assert b.property('active')
    dlg._apply_pin('none')
    _pump(app)
    assert not w._cal_pin_badge.isVisible()
    dlg.close()
    _pump(app)
    w._refresh_bottom_active()
    assert not b.property('active')


def test_doc_and_budget_buttons_follow_window(win, monkeypatch):
    w, app = win
    from PyQt5.QtWidgets import QWidget
    fake = QWidget()
    monkeypatch.setattr(w, '_bottom_window', lambda k: fake if k in ('doc', 'budget') else None)
    w._refresh_bottom_active()
    assert not w._bottom_btns['doc'].property('active')
    fake.show()
    _pump(app)
    w._refresh_bottom_active()
    assert w._bottom_btns['doc'].property('active') and w._bottom_btns['budget'].property('active')
    fake.hide()
    w._refresh_bottom_active()
    assert not w._bottom_btns['doc'].property('active')
    fake.deleteLater()


def test_refresh_survives_deleted_window(win, monkeypatch):
    w, app = win
    from PyQt5.QtCore import QEvent
    from PyQt5.QtWidgets import QWidget
    gone = QWidget()
    monkeypatch.setattr(w, '_bottom_window', lambda k: gone)
    gone.deleteLater()
    app.sendPostedEvents(None, QEvent.DeferredDelete)
    w._refresh_bottom_active()                                       # 지워진 창이어도 오류 없음


def test_pin_button_vector_icons(win):
    w, app = win
    w._show_monthly_calendar()
    _pump(app)
    dlg = w._calendar_dlg
    btn = dlg.btn_pin
    assert btn.text() == ''                                          # 이모지 글자 없음 (직접 그림)
    assert btn.label == '고정' and not btn.on
    dlg._apply_pin('desktop')
    assert btn.label == '바탕화면' and btn.on
    btn.grab()                                                       # 그려도 오류 없음
    dlg._apply_pin('none')
