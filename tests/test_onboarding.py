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


def _finish_step(t):
    """퀘스트를 다 했으면 완료 연출을 건너뛰고 다음 단계로."""
    t._check()
    assert t._burst is not None, f'완료되지 않음: {t._key()} {t.done}'
    t._burst -= 10_000
    t._tick()


def test_full_tour(win, monkeypatch):
    w, app = win
    import PyQt5.QtWidgets as qw
    from ui import onboarding
    t = onboarding.maybe_start(w)
    assert t is not None and w._tour is t and t.intro              # 먼저 유주 인사
    t.begin()
    assert t.step == 0 and not t.intro
    assert t.btn_action.isHidden()                                 # [다음] 같은 단추 없음 — 직접 해야
    # 1 업무 추가: 쓰기 → Tab(마감일 칸) → Enter
    w.input_name.setText('체험학습 계획서 제출')
    t._check()
    assert t.done == [False, False, False]                          # 치자마자는 아직
    t._t += onboarding.TYPE_PAUSE_MS
    t._check()
    assert t.done == [True, False, False]                           # 1.5초 멈추면 ✓
    focus = {'w': w.input_date}
    monkeypatch.setattr(qw.QApplication, 'focusWidget', staticmethod(lambda: focus['w']))
    t._check()
    assert t.done == [True, False, False] and '마감일' in t._line   # 마감일 칸: 아직
    focus['w'] = w.input_time
    t._check()
    assert not t.done[1] and '시간' in t._line
    focus['w'] = w.input_recurrence
    t._check()
    assert t.done == [True, True, False]                            # 반복 칸까지 와야 ✓
    w._add_task()
    _pump(app)
    _finish_step(t)
    assert t._key() == 'row' and t._row().task['name'] == '체험학습 계획서 제출'
    # 2 업무 다루기: 별 → 메모 → ⋮ 메뉴
    import db
    db.update_task(t.task_id, '체험학습 계획서 제출', '', 0, 1, '')
    t._check()
    assert t.done[0] and not t.done[1]
    row = t._row()
    row._expanded = True
    t._check()
    assert t.done[1]
    t._menu_seen = True
    _finish_step(t)
    # 3 파일
    assert t._key() == 'file'
    t._row()._attachments.append('C:/문서/계획서.hwp')
    _finish_step(t)
    # 4 아래 줄 버튼: 열기 → 닫기
    assert t._key() == 'bottom'
    state = {'open': True}
    monkeypatch.setattr(t, '_bottom_open', lambda k: state['open'])
    t._check()
    assert t.done == [True, False]
    state['open'] = False
    _finish_step(t)
    # 5 단축키: Ctrl+Shift+X → Ctrl+Shift+E → Ctrl+K (순서대로)
    assert t._key() == 'keys'
    onboarding.notify('restore')                                    # 순서가 앞서면 아직
    t._check()
    assert t.done == [False, False, False]
    onboarding.notify('capture')
    t._check()
    assert t.done == [True, True, False]                            # 앞서 누른 E도 이제 인정
    onboarding.notify('palette')
    _finish_step(t)
    # 5 AI: 이미 연결돼 있으면 저절로
    import ai_client
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'claude_code')
    assert t._key() == 'ai'
    _finish_step(t)
    # 6 메뉴 → 클리어
    assert t._key() == 'menu'
    t._menu_seen = True
    _finish_step(t)
    assert t.clear and t.btn_action.text() == '시작하기' and len(t.completed) == 7
    assert not onboarding.is_done()
    t.btn_action.click()
    _pump(app)
    assert onboarding.is_done() and w._tour is None
    assert onboarding.maybe_start(w) is None                       # 한 번만


def test_enter_without_tab_needs_another(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    w.input_name.setText('하나')
    w._add_task()
    _pump(app)
    t._check()
    assert t.step == 0 and t.done == [True, False, False] and t._burst is None


def test_skip_step_and_ai_later(win, monkeypatch):
    w, app = win
    import ai_client
    from ui import onboarding
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    t = onboarding.start(w)
    t.begin()
    for _ in range(5):
        t.btn_skip.click()
    assert t._key() == 'ai'
    t._layout_buttons()
    assert not t.btn_action.isHidden() and t.btn_later.text() == '나중에' and t.btn_skip.isHidden()
    t.btn_later.click()
    assert t._key() == 'menu'


def test_hole_lets_clicks_through(win):
    w, app = win
    from PyQt5.QtCore import QPoint
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    t._tick()
    c = t._hole.center()
    assert not t.mask().contains(c)                          # 뚫린 곳은 아래 칸이 받음
    assert t.mask().contains(QPoint(2, t.height() - 2))


def test_skip_marks_done(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    t.btn_quit.click()
    _pump(app)
    assert onboarding.is_done() and w._tour is None


def test_existing_user_not_shown(win):
    w, app = win
    import db
    from ui import onboarding
    db.add_task(w.window_id, '쓰던 업무', '')
    assert onboarding.maybe_start(w) is None and onboarding.is_done()


def test_replay_from_menu(win):
    w, app = win
    import db
    db.add_task(w.window_id, '쓰던 업무', '')
    w._refresh_tasks()
    _pump(app)
    w.start_tour()
    t = w._tour
    t.begin()
    assert t.step == 0 and t.done == [False, False, False]        # 다시 볼 때도 직접 해야


def test_typing_again_restarts_pause(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    w.input_name.setText('체')
    t._check()
    t._t += 1000
    w.input_name.setText('체험')                                   # 다시 치면 처음부터
    t._check()
    t._t += 1000
    t._check()
    assert not t.done[0]
    t._t += 600
    t._check()
    assert t.done[0]


def test_card_goes_beside_window_when_room(win, monkeypatch):
    w, app = win
    from PyQt5.QtCore import QPoint
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    monkeypatch.setattr(t, '_side_spot', lambda pw, ph: QPoint(900, 100))
    t._tick()
    assert t._outside and t.btn_quit.parent() is t.panel and t._card.left() < 20
    monkeypatch.setattr(t, '_side_spot', lambda pw, ph: None)      # 자리가 없으면 창 안쪽
    t._tick()
    assert not t._outside and t.btn_quit.parent() is t and t.panel.isHidden()


def test_intro_waits_for_start_button(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    assert t.intro and t.btn_action.text() == '시작하기' and t.btn_quit.isHidden() and t.btn_skip.isHidden()
    assert t._typed == len(t._line)                                 # 글은 한 번에
    t._t += 60_000
    t._tick()
    assert t.intro                                                  # 기다려도 저절로 넘어가지 않음


def test_intro_start_button(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.btn_action.click()
    assert not t.intro and t._key() == 'add'
