"""업무창 AI 대화: 제안 읽기·확인(날짜·없는 업무·완료), 넣기·되돌리기, 붙는 자리 계산, 창 열고 숨기기."""
import json
import os
import sys
import unittest.mock as um
from datetime import date, datetime

import pytest

import task_ai

TODAY = date(2026, 11, 12)          # 목요일
TASKS = [
    {'id': 1, 'window_id': 1, 'name': '교육청 공문 회신', 'deadline': '2026-11-14', 'priority': 0,
     'recurrence': '', 'strikethrough': 0},
    {'id': 2, 'window_id': 1, 'name': '협의회 자료 준비', 'deadline': '2026-11-18 15:00', 'priority': 1,
     'recurrence': 'weekly', 'strikethrough': 0},
    {'id': 3, 'window_id': 2, 'name': '끝난 일', 'deadline': '', 'priority': 0, 'recurrence': '', 'strikethrough': 1},
]
BY_ID = {t['id']: t for t in TASKS}


def _reply(**kw):
    return json.dumps({'answer': '확인', **kw}, ensure_ascii=False)


def test_context_and_prompt_have_ids_dates_and_today():
    ctx = task_ai.build_context(TASKS, {1: '업무창 1', 2: '업무창 2'}, current_window_id=1)
    assert 'id=2 | 업무창 1 (지금 창) | 협의회 자료 준비 | 마감 2026-11-18 15:00 (11월 18일(수) 15:00) | 중요 | 반복 매주' in ctx
    assert 'id=3 | 업무창 2 | 끝난 일 | 마감 없음 | 완료' in ctx
    p = task_ai.build_prompt('내일 할 일?', ctx, now=datetime(2026, 11, 12, 9, 30))
    assert '오늘: 2026-11-12 (목요일), 지금 시각 09:30' in p


def test_add_and_change_proposals_are_checked():
    raw = _reply(add=[{'name': ' 수행평가  채점 ', 'deadline': '2026-11-20', 'important': True},
                      {'name': '시각 있는 일', 'deadline': '2026-11-21T09:05'},
                      {'name': '날짜 틀림', 'deadline': '2026-02-30'},
                      {'name': '', 'deadline': ''},
                      {'name': '010-****-1234 연락', 'deadline': ''}],
                 set_deadline=[{'task_id': 1, 'deadline': '2026-11-19 15:00'},
                               {'task_id': 2, 'deadline': '2026-11-18 15:00'},      # 그대로: 조용히 뺌
                               {'task_id': 99, 'deadline': '2026-11-19'}],
                 done=[{'task_id': 1}, {'task_id': 3}],
                 important=[{'task_id': 1, 'on': True}, {'task_id': 2, 'on': True}])
    r = task_ai.parse_reply(raw, BY_ID, TODAY)
    kinds = [(c['kind'], c.get('name') or c['task']['name']) for c in r['changes']]
    assert kinds == [('add', '수행평가 채점'), ('add', '시각 있는 일'), ('deadline', '교육청 공문 회신'),
                     ('done', '교육청 공문 회신'), ('important', '교육청 공문 회신')]
    assert r['changes'][0]['important'] is True and r['changes'][1]['deadline'] == '2026-11-21 09:05'
    reasons = {(x['label'], x['reason']) for x in r['rejected']}
    assert ('날짜 틀림', '없는 날짜예요') in reasons
    assert ('(이름 없음)', '업무 이름이 비어 있어요') in reasons
    assert ('업무 99', '없는 업무예요') in reasons
    assert ('끝난 일', '이미 완료된 업무예요') in reasons
    assert any('가린 개인정보' in x['reason'] for x in r['rejected'])


def test_not_json_is_plain_answer_and_deadline_can_be_cleared():
    assert task_ai.parse_reply('그냥 답', BY_ID, TODAY) == {'answer': '그냥 답', 'changes': [], 'rejected': []}
    r = task_ai.parse_reply(_reply(set_deadline=[{'task_id': 1, 'deadline': ''}]), BY_ID, TODAY)
    assert r['changes'][0]['new'] == '' and '마감 없음' in task_ai.describe(r['changes'][0])


@pytest.mark.parametrize('value, out, ok', [
    ('2026-11-20', '2026-11-20', True), ('2026-1-5 9:00', '2026-01-05 09:00', True),
    ('2026-11-20 25:00', None, False), ('다음 주', None, False), ('2099-01-01', None, False), ('', '', True),
])
def test_norm_deadline(value, out, ok):
    got, reason = task_ai.norm_deadline(value, TODAY)
    assert got == out and (reason is None) == ok


# ── DB에 넣기 · 되돌리기 ─────────────────────────────────────────
@pytest.fixture
def tdb(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    return db


def test_apply_then_undo(tdb):
    wid = tdb.get_first_window_id() or tdb.create_window()
    t1 = tdb.add_task(wid, '공문 회신', '2026-11-14')
    task = tdb.get_task(t1)
    changes = [{'kind': 'add', 'task': None, 'name': '새 일', 'deadline': '2026-11-20', 'important': True},
               {'kind': 'deadline', 'task': task, 'old': '2026-11-14', 'new': '2026-11-19 15:00'},
               {'kind': 'done', 'task': task},
               {'kind': 'important', 'task': task, 'on': True}]
    undo = task_ai.apply_changes(changes, wid)
    after = tdb.get_task(t1)
    assert (after['deadline'], after['strikethrough'], after['priority']) == ('2026-11-19 15:00', 1, 1)
    new = [t for t in tdb.get_tasks(wid) if t['name'] == '새 일']
    assert len(new) == 1 and new[0]['priority'] == 1 and new[0]['window_id'] == wid
    task_ai.undo_changes(undo)
    back = tdb.get_task(t1)
    assert (back['deadline'], back['strikethrough'], back['priority']) == ('2026-11-14', 0, 0)
    assert [t['name'] for t in tdb.get_tasks(wid)] == ['공문 회신']


def test_task_windows_tasks_skips_memo_windows(tdb):
    wid = tdb.get_first_window_id() or tdb.create_window()
    tdb.add_task(wid, '업무', '')
    tasks, labels = task_ai.task_windows_tasks()
    assert [t['name'] for t in tasks] == ['업무'] and labels[wid] == '업무창 1'
    tdb.set_window_memo(wid, 1, '')
    assert task_ai.task_windows_tasks()[0] == []


# ── 붙는 자리 · 창 ───────────────────────────────────────────────
pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


def test_dock_rect_flips_at_screen_edge():
    from PyQt5.QtCore import QRect
    from ui.ai_chat_window import dock_rect
    screen = QRect(0, 0, 1920, 1040)
    r, where = dock_rect(QRect(100, 100, 300, 400), 'right', 320, screen)
    assert where == 'right' and (r.x(), r.y(), r.width(), r.height()) == (400, 100, 320, 400)
    r, where = dock_rect(QRect(1700, 100, 200, 400), 'right', 320, screen)      # 오른쪽 자리 없음 → 왼쪽
    assert where == 'left' and r.x() == 1700 - 320
    r, where = dock_rect(QRect(100, 100, 300, 400), 'bottom', 240, screen)
    assert where == 'bottom' and (r.x(), r.y(), r.width(), r.height()) == (100, 500, 300, 240)
    r, where = dock_rect(QRect(100, 700, 300, 300), 'bottom', 240, screen)      # 아래 자리 없음 → 위
    assert where == 'top' and r.y() == 700 - 240


_APP = None                         # QApplication을 붙잡아 둬야 다음 테스트에서도 창을 만들 수 있음


@pytest.fixture
def win(tdb, tmp_path, monkeypatch):
    global _APP
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    from PyQt5.QtWidgets import QApplication, QMenu
    app = _APP = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMenu, 'exec_', lambda self, *a: None)
    import window
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, '_win', None)
    w = window.MemoWindow(tdb.get_first_window_id())
    w.move(100, 100)
    w.resize(320, 400)
    w.show()
    app.processEvents()
    yield w, acw, tdb, app
    if acw._win is not None:
        acw._win.close()
        acw._win.deleteLater()
    w.close()


def test_toggle_attach_follow_and_remember(win):
    w, acw, db, app = win
    w._toggle_ai_chat()
    c = acw.instance()
    assert c.isVisible() and c.host is w and db.get_setting(acw.OPEN_KEY) == '1'
    assert c.geometry().x() == w.frameGeometry().right() + 1          # 오른쪽에 붙음
    w.move(150, 120)
    app.processEvents()
    assert c.geometry().x() == w.frameGeometry().right() + 1 and c.geometry().y() == w.frameGeometry().y()
    c.set_side('bottom')
    assert db.get_setting(acw.SIDE_KEY) == 'bottom' and c.geometry().y() == w.frameGeometry().bottom() + 1
    w.toggle_shade()                                                   # 접으면 함께 숨고
    assert not c.isVisible()
    w.toggle_shade()                                                   # 펼치면 다시 붙음
    assert c.isVisible()
    w._toggle_ai_chat()                                                # 💬 다시 누르면 숨기기
    assert not c.isVisible() and db.get_setting(acw.OPEN_KEY) == '0'


def test_chat_proposal_apply_refreshes_window_and_undo(win, monkeypatch):
    w, acw, db, app = win
    wid = w.window_id
    tid = db.add_task(wid, '공문 회신', '2026-11-14')
    w._refresh_tasks()
    w._toggle_ai_chat()
    c = acw.instance()
    tasks = {t['id']: t for t in db.get_tasks(wid)}
    reply = task_ai.parse_reply(_reply(add=[{'name': '새 일', 'deadline': ''}], done=[{'task_id': tid}]), tasks)
    c._on_done(reply, {})
    assert c.proposal.isVisibleTo(c) and c.btn_apply.isVisibleTo(c)
    c._apply()
    assert db.get_task(tid)['strikethrough'] == 1
    assert '새 일' in [r.task['name'] for r in w._task_rows()]
    c._undo_last()
    assert db.get_task(tid)['strikethrough'] == 0 and '새 일' not in [t['name'] for t in db.get_tasks(wid)]


def test_ai_off_disables_input(win, monkeypatch, tmp_path):
    w, acw, db, app = win
    import ai_client
    monkeypatch.setattr(ai_client, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai_client, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    from ui import ai_connect
    ai_connect.turn_off()
    w._toggle_ai_chat()
    c = acw.instance()
    assert not c.input.isEnabled() and 'AI 꺼짐' in c.btn_status.text()
    ai_client.save_ai_mode('openai')
    ai_connect.notify_changed()
    assert c.input.isEnabled()


def test_title_bar_ai_button_shows_open_state(win):
    w, acw, db, app = win
    btn = w.title_bar.btn_ai
    assert not btn.icon().isNull() and btn.toolTip() == 'AI 대화 열기'
    w._toggle_ai_chat()
    assert btn.property('active') is True and btn.toolTip() == 'AI 대화 숨기기'
    w._toggle_ai_chat()
    assert btn.property('active') is False and btn.toolTip() == 'AI 대화 열기'


def test_ai_chat_remembers_host_window(win):
    w, acw, db, app = win
    import window
    wid2 = db.create_window()
    w2 = window.MemoWindow(wid2)
    w2.show()
    app.processEvents()
    acw.instance().attach(w2)
    assert db.get_setting(acw.HOST_KEY) == str(wid2)
    acw._win.close()
    acw._win = None
    acw.auto_open(w)                    # 다음에 켤 때: 먼저 뜬 창(w)이 아니라
    assert acw.instance(create=False) is None
    acw.auto_open(w2)                   # 기억한 창(w2) 옆에 붙음
    assert acw.instance(create=False).host is w2
    acw.reset_layout()
    assert db.get_setting(acw.SIDE_KEY) == 'right'
    w2.close()
