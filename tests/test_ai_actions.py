"""AI가 앱을 다루는 기능 목록 (ai_actions): 검사 · 위험 정도별 분류 · 실행 · 되돌리기."""
import json
import os
import sys
import unittest.mock as um

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

import ai_actions  # noqa: E402

_APP = None


@pytest.fixture
def env(tmp_path, monkeypatch):
    import db
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication, QMenu
    global _APP
    _APP = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMenu, 'exec_', lambda self, *a: None)
    import window
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, '_win', None)
    monkeypatch.setattr(acw, 'saved_open', lambda: False)
    opened = []
    windows = []

    def on_new(offset_from=None, memo_mode=False, **kw):            # main.new_window 대신
        wid = db.create_window()
        if memo_mode:
            db.set_window_memo(wid, 1, '')
        nw = window.MemoWindow(wid, open_windows=windows, force_memo_mode=memo_mode)
        windows.append(nw)
    w = window.MemoWindow(db.get_first_window_id() or db.create_window(), on_new=on_new, open_windows=windows)
    windows.append(w)
    w.show()
    for name in ('_show_monthly_calendar', '_open_dashboard', '_open_budget_window', '_open_document_editor'):
        monkeypatch.setattr(w, name, lambda n=name: opened.append(n))
    monkeypatch.setattr(w, '_show_history', lambda tab=None: opened.append(('history', tab)))
    w._toggle_ai_chat()
    chat = acw.instance()
    from ui.ai_chat_window import ChatApp
    app = ChatApp(chat)
    yield {'db': db, 'w': w, 'chat': chat, 'app': app, 'opened': opened, 'windows': windows}
    for x in windows:
        x.close()
    if acw._win is not None:
        acw._win.close()


def _reply(*actions, answer='확인'):
    return json.dumps({'answer': answer, 'actions': list(actions)}, ensure_ascii=False)


def _run(env, *actions):
    """검사 → (바꾸기·지우기 모두) 실행 → 되돌리기 기록."""
    ctx = ai_actions.build_state(env['app'])
    r = ai_actions.parse_reply(_reply(*actions), ctx)
    assert not r['rejected'], r['rejected']
    return r, ai_actions.execute(r['now'] + r['changes'] + r['deletes'], env['app'])


def test_buckets_by_risk_and_rejections(env):
    db, w = env['db'], env['w']
    tid = db.add_task(w.window_id, '공문 회신', '2026-11-20')
    ctx = ai_actions.build_state(env['app'])
    r = ai_actions.parse_reply(_reply(
        {'do': 'open', 'what': 'calendar'},
        {'do': 'task.rename', 'task_id': tid, 'name': '공문 회신하기'},
        {'do': 'task.delete', 'task_id': tid},
        {'do': 'app.quit'},
        {'do': 'task.delete', 'task_id': 9999},
        {'do': 'task.rename', 'task_id': tid, 'name': '공문 회신'},          # 그대로 → 조용히 뺌
    ), ctx)
    assert [i['do'] for i in r['now']] == ['open']
    assert [i['do'] for i in r['changes']] == ['task.rename']
    assert [i['do'] for i in r['deletes']] == ['task.delete']
    assert {x['reason'] for x in r['rejected']} == {'앱에 없는 기능이에요', '없는 업무예요'}


def test_max_actions(env):
    ctx = ai_actions.build_state(env['app'])
    acts = [{'do': 'task.add', 'name': f'일 {i}'} for i in range(ai_actions.MAX_ACTIONS + 3)]
    r = ai_actions.parse_reply(_reply(*acts), ctx)
    assert len(r['changes']) == ai_actions.MAX_ACTIONS and len(r['rejected']) == 3


def test_task_actions_and_undo(env):
    db, w = env['db'], env['w']
    tid = db.add_task(w.window_id, '공문 회신', '2026-11-20')
    w2 = env['app'].new_window('task')
    r, j = _run(env, {'do': 'task.rename', 'task_id': tid, 'name': '교육청 공문 회신'},
                {'do': 'task.repeat', 'task_id': tid, 'repeat': 'weekly'},
                {'do': 'task.note', 'task_id': tid, 'text': '담당 확인'},
                {'do': 'task.related', 'task_id': tid, 'text': '교육과-123'},
                {'do': 'task.important', 'task_id': tid, 'on': True},
                {'do': 'task.move', 'task_id': tid, 'window_id': w2.window_id})
    t = db.get_task(tid)
    assert (t['name'], t['recurrence'], t['notes'], t['related_no'], t['priority'], t['window_id']) == \
        ('교육청 공문 회신', 'weekly', '담당 확인', '교육과-123', 1, w2.window_id)
    ai_actions.undo(j, env['app'])
    t = db.get_task(tid)
    assert (t['name'], t['recurrence'], t['notes'], t['related_no'], t['priority'], t['window_id']) == \
        ('공문 회신', '', '', '', 0, w.window_id)


def test_task_delete_goes_to_history_and_undo_restores(env):
    db, w = env['db'], env['w']
    tid = db.add_task(w.window_id, '지울 일', '2026-11-20', priority=1)
    r, j = _run(env, {'do': 'task.delete', 'task_id': tid})
    assert db.get_task(tid) is None
    assert [h['name'] for h in db.get_task_history(w.window_id)] == ['지울 일']
    ai_actions.undo(j, env['app'])
    names = [(t['name'], t['priority']) for t in db.get_tasks(w.window_id)]
    assert names == [('지울 일', 1)] and db.get_task_history(w.window_id) == []


def test_repeat_needs_deadline_and_mask_rejected(env):
    db, w = env['db'], env['w']
    tid = db.add_task(w.window_id, '마감 없음', '')
    ctx = ai_actions.build_state(env['app'])
    r = ai_actions.parse_reply(_reply({'do': 'task.repeat', 'task_id': tid, 'repeat': 'monthly'},
                                      {'do': 'task.add', 'name': '010-****-1234 연락'}), ctx)
    reasons = [x['reason'] for x in r['rejected']]
    assert '마감이 있어야 반복할 수 있어요' in reasons and any('가린 개인정보' in x for x in reasons)


def test_new_memo_write_style_and_undo(env):
    db, w = env['db'], env['w']
    r, j = _run(env, {'do': 'window.new', 'kind': 'memo', 'text': '회의 메모'})
    memo = env['windows'][-1]
    assert memo._memo_mode and memo.memo_editor.toPlainText() == '회의 메모'
    r, j2 = _run(env, {'do': 'memo.write', 'window_id': memo.window_id, 'text': '두 번째 줄', 'mode': 'append'},
                 {'do': 'window.style', 'window_id': memo.window_id, 'color': '분홍', 'size': '크게', 'gap': 10})
    assert memo.memo_editor.toPlainText() == '회의 메모\n두 번째 줄'
    assert memo._bg_color.upper() == '#FBD5E3' and memo._scale == 1.1 and memo._row_gap == 10
    ai_actions.undo(j2, env['app'])
    assert memo.memo_editor.toPlainText() == '회의 메모' and memo._scale == 1.0 and memo._row_gap == 0
    wid = memo.window_id
    ai_actions.undo(j, env['app'])                                        # 새 메모장 되돌리기 = 지움
    assert all(x['id'] != wid for x in db.get_all_windows())


def test_open_is_immediate_and_cannot_delete_windows(env):
    ctx = ai_actions.build_state(env['app'])
    r = ai_actions.parse_reply(_reply({'do': 'open', 'what': 'history'}, {'do': 'window.delete', 'window_id': 1},
                                      {'do': 'budget.clear_all'}), ctx)
    ai_actions.execute(r['now'], env['app'])
    assert env['opened'] == [('history', 0)]
    assert len(r['rejected']) == 2                                         # 창 지우기 · 예산 전체 지우기는 없음


def test_budget_actions_and_undo(env):
    db = env['db']
    cid = db.add_budget_category('교육과정 운영')
    iid = db.add_budget_item(cid, '운영수당', '강사 수당', 1000000, 200000)
    r, j = _run(env, {'do': 'budget.spend', 'item_id': iid, 'amount': 50000, 'purpose': '특강'},
                {'do': 'budget.item_add', 'category_id': cid, 'cost_type': '교육운영비', 'description': '간식',
                 'budget': 300000},
                {'do': 'budget.cat_group', 'category_id': cid, 'group': '본예산'},
                {'do': 'budget.cat_exclude', 'category_id': cid, 'on': True})
    items = {i['description']: i for i in db.get_budget_items()}
    assert items['강사 수당']['spent'] == 250000 and '간식' in items
    cat = db.get_budget_categories()[0]
    assert cat['excluded'] == 1 and db.get_budget_groups()[0]['name'] == '본예산'
    ai_actions.undo(j, env['app'])
    items = {i['description']: i for i in db.get_budget_items()}
    assert list(items) == ['강사 수당'] and items['강사 수당']['spent'] == 200000
    assert db.get_budget_expenses() == [] and db.get_budget_groups() == []
    assert db.get_budget_categories()[0]['excluded'] == 0


def test_budget_delete_and_restore(env):
    db = env['db']
    cid = db.add_budget_category('방과후')
    iid = db.add_budget_item(cid, '교육운영비', '간식', 500000, 0)
    db.add_budget_expense(iid, 30000, '1차 간식', '2026-11-01')
    r, j = _run(env, {'do': 'budget.item_delete', 'item_id': iid})
    assert r['deletes'] and db.get_budget_items() == []
    ai_actions.undo(j, env['app'])
    it = db.get_budget_items()[0]
    assert (it['description'], it['spent']) == ('간식', 30000)
    assert [e['purpose'] for e in db.get_budget_expenses()] == ['1차 간식']
    r, j = _run(env, {'do': 'budget.cat_delete', 'category_id': cid})
    assert db.get_budget_categories() == []
    ai_actions.undo(j, env['app'])
    assert [c['name'] for c in db.get_budget_categories()] == ['방과후'] and len(db.get_budget_items()) == 1


def test_budget_checks(env):
    db = env['db']
    cid = db.add_budget_category('A')
    iid = db.add_budget_item(cid, 'x', '항목', 1000, 0)
    ctx = ai_actions.build_state(env['app'])
    r = ai_actions.parse_reply(_reply({'do': 'budget.set_remaining', 'item_id': iid, 'remaining': 5000},
                                      {'do': 'budget.cat_add', 'name': 'A'},
                                      {'do': 'budget.item_add', 'category_id': 999, 'description': 'y', 'budget': 1}), ctx)
    reasons = {x['reason'] for x in r['rejected']}
    assert '잔액 목표가 예산보다 커요' in reasons and '같은 이름의 세부항목이 이미 있어요' in reasons
    assert '없는 세부항목이에요' in reasons


def test_settings_and_chat_actions(env, monkeypatch):
    app = env['app']
    alarm = {'m': 180}
    monkeypatch.setattr(type(app), 'get_alarm', lambda self: alarm['m'])
    monkeypatch.setattr(type(app), 'set_alarm', lambda self, m: alarm.update(m=m))
    import updater
    monkeypatch.setattr(updater, 'auto_enabled', lambda: True)
    calls = []
    monkeypatch.setattr(updater, 'set_auto_enabled', lambda on: calls.append(on))
    r, j = _run(env, {'do': 'settings.alarm', 'minutes': 60}, {'do': 'settings.auto_update', 'on': False},
                {'do': 'chat.side', 'side': 'right'}, {'do': 'chat.style', 'color': '하늘', 'size': '크게'})
    assert alarm['m'] == 60 and calls == [False]
    assert env['chat'].side == 'right' and env['chat'].color == '#D2E4FA' and env['chat'].scale == 1.1
    ai_actions.undo(j, app)
    assert alarm['m'] == 180 and calls == [False, True] and env['chat'].side == 'inside' and env['chat'].color == ''


def test_prompt_has_catalog_and_state(env):
    db, w = env['db'], env['w']
    db.add_task(w.window_id, '공문 회신', '2026-11-20')
    ctx = ai_actions.build_state(env['app'])
    p = ai_actions.build_prompt('지워줘', ctx)
    assert 'task.delete' in p and 'budget.spend' in p and 'window.new' in p
    assert 'window_id=' in p and '공문 회신' in p and '[창]' in p


def test_chat_shows_delete_box_and_runs_after_confirm(env):
    db, w, chat = env['db'], env['w'], env['chat']
    tid = db.add_task(w.window_id, '지울 일', '')
    ctx = ai_actions.build_state(env['app'])
    reply = ai_actions.parse_reply(_reply({'do': 'task.delete', 'task_id': tid}, answer='지울게요'), ctx)
    chat._on_done(reply, {})
    assert chat.proposal_del.isVisibleTo(chat) and not chat.proposal.isVisibleTo(chat)
    assert db.get_task(tid) is not None                                    # 확인 전에는 그대로
    chat._apply_deletes()
    assert db.get_task(tid) is None
    chat._undo_last()
    assert [t['name'] for t in db.get_tasks(w.window_id)] == ['지울 일']
