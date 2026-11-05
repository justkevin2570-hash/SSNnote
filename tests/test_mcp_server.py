import importlib
from datetime import date, timedelta

import pytest

pytest.importorskip('mcp')


@pytest.fixture
def srv(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    import mcp_server
    return importlib.reload(mcp_server), db


def d(n):
    return (date.today() + timedelta(days=n)).isoformat()


def test_add_list_update_flow(srv):
    m, db = srv
    sig0 = db.get_setting(m.EXTERNAL_CHANGE_KEY, '')
    t = m.add_task('교과협의회 계획서 제출', deadline=d(5), time='17:00', notes='☐ 협의록', important=True)['added']
    assert t['dday'] == 'D-5' and t['time'] == '17:00' and t['important'] and not t['done']
    assert db.get_setting(m.EXTERNAL_CHANGE_KEY, '') != sig0          # 앱 새로고침 신호
    assert db.get_task(t['task_id'])['window_id'] == db.get_first_window_id()

    m.update_task(t['task_id'], time='15:00')                           # 시각만 → 날짜 유지
    assert db.get_task(t['task_id'])['deadline'] == d(5) + ' 15:00'
    m.update_task(t['task_id'], deadline=d(7))                          # 날짜만 → 시각 유지
    assert db.get_task(t['task_id'])['deadline'] == d(7) + ' 15:00'
    m.update_task(t['task_id'], done=True)
    row = db.get_task(t['task_id'])
    assert row['strikethrough'] == 1 and row['priority'] == 1 and row['notes'] == '☐ 협의록'
    assert [x['name'] for x in m.list_tasks('done')['tasks']] == ['교과협의회 계획서 제출']
    assert m.list_tasks('open')['tasks'] == []
    m.update_task(t['task_id'], deadline='')
    assert db.get_task(t['task_id'])['deadline'] == ''


def test_overview_and_search(srv):
    m, db = srv
    m.add_task('학교운영위원회 회의록 제출', deadline=d(-1))
    m.add_task('방과후 강사 계약 공문 회신', deadline=d(3), related_doc='교육과정과-12345')
    ov = m.today_overview()
    assert ov['today'] == date.today().isoformat()
    assert [t['name'] for t in ov['overdue']] == ['학교운영위원회 회의록 제출']
    assert ov['open_official_doc_tasks'] == 1
    assert m.search_tasks('ㅂㄱㅎ')['tasks'][0]['name'] == '방과후 강사 계약 공문 회신'
    assert m.search_tasks('12345')['tasks'][0]['related_doc'] == '교육과정과-12345'


@pytest.mark.parametrize('kw,msg', [
    ({'name': ''}, '비어'),
    ({'name': 'x', 'deadline': '11월 14일'}, '날짜 형식'),
    ({'name': 'x', 'deadline': '2026-11-14', 'time': '5pm'}, '시각 형식'),
    ({'name': 'x', 'time': '17:00'}, '시각만'),
])
def test_add_validation_errors_reach_claude(srv, kw, msg):
    m, _ = srv
    from mcp.server.mcpserver.exceptions import ToolError
    with pytest.raises(ToolError, match=msg):
        m.add_task(**kw)


def test_update_missing_task(srv):
    m, _ = srv
    from mcp.server.mcpserver.exceptions import ToolError
    with pytest.raises(ToolError, match='찾을 수 없습니다'):
        m.update_task(99999, done=True)


def test_document_guide_includes_profile(srv):
    m, _ = srv
    import school_profile
    school_profile.save_profile({'school': '서서니고', 'name': '홍길동'})
    g = m.official_document_guide()
    assert '서서니고' in g['school_profile'] and any('끝.' in r for r in g['format_rules'])
