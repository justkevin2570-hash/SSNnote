"""첫 사용 안내(코치마크): 처음에만 · 업무를 넣어야 다음 · 파일 놓으면 문구 바뀜 · 건너뛰기/끝."""
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
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, 'auto_open', lambda h: None)
    import window
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    w.resize(420, 460)
    w.show()
    app.processEvents()
    yield w, app
    if getattr(w, '_tour', None) is not None:
        w._tour.finish(mark=False)
    w.close()


def _pump(app, n=5):
    for _ in range(n):
        app.processEvents()


def test_full_tour(win):
    w, app = win
    from ui import onboarding
    t = onboarding.maybe_start(w)
    assert t is not None and w._tour is t and t.step == 0
    assert t.bubble.btn_next.isHidden()                      # 업무를 넣어야 다음 ([다음] 없음)
    w.input_name.setText('체험학습 계획서 제출')
    w._add_task()
    _pump(app)
    assert t.step == 1 and t._row().task['name'] == '체험학습 계획서 제출'
    t.next()
    assert t.step == 2 and t.bubble.lbl_body.text() == onboarding.TEXT['file'][1]
    row = t._row()
    row._attachments.append('C:/문서/계획서.hwp')
    t._tick()
    assert t.bubble.lbl_body.text() == onboarding.TEXT['file_done'][1]
    t.next(); t.next()
    assert t.step == 4 and t.bubble.btn_next.text() == '시작하기'
    assert not onboarding.is_done()
    t.next()
    _pump(app)
    assert onboarding.is_done() and w._tour is None
    assert onboarding.maybe_start(w) is None                 # 한 번만


def test_hole_lets_clicks_through(win):
    w, app = win
    from PyQt5.QtCore import QPoint
    from ui import onboarding
    t = onboarding.start(w)
    t._tick()
    c = t._hole.center()
    assert not t.mask().contains(c)                          # 뚫린 곳은 아래 칸이 받음
    assert t.mask().contains(QPoint(2, t.height() - 2))


def test_skip_marks_done(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.bubble.btn_skip.click()
    _pump(app)
    assert onboarding.is_done() and w._tour is None


def test_existing_user_not_shown(win):
    w, app = win
    import db
    from ui import onboarding
    db.add_task(w.window_id, '쓰던 업무', '')
    assert onboarding.maybe_start(w) is None and onboarding.is_done()


def test_replay_from_menu_still_needs_new_task(win):
    w, app = win
    import db
    from ui import onboarding
    db.add_task(w.window_id, '쓰던 업무', '')
    w._refresh_tasks()
    _pump(app)
    w.start_tour()
    t = w._tour
    assert t.step == 0 and t.bubble.btn_next.isHidden()      # 다시 볼 때도 직접 하나 넣어야
    w.input_name.setText('새로 넣은 업무')
    w._add_task()
    _pump(app)
    assert t.step == 1 and t._row().task['name'] == '새로 넣은 업무' and not t.bubble.btn_next.isHidden()
