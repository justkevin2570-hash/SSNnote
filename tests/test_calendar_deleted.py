"""달력: 업무 목록에서 지운 업무도 마감일 칸에 회색(지난 업무)으로 남고, 지난 기록에서 지우면 사라진다."""
import os
import sys
import unittest.mock as um
from datetime import date

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
_APP = None


@pytest.fixture
def env(tmp_path, monkeypatch):
    global _APP
    import db
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    from ui import ai_chat_window
    monkeypatch.setattr(ai_chat_window, 'saved_open', lambda: False)
    import window
    w = window.MemoWindow(db.get_first_window_id())
    w._open_windows = [w]
    yield db, w
    w.close()


def _cell(cal, ds):
    from PyQt5.QtCore import QEvent
    from ui.calendar_view import DayCellWidget
    _APP.sendPostedEvents(None, QEvent.DeferredDelete)                  # 다시 그리기 전 칸은 치움
    return next(c for c in cal.findChildren(DayCellWidget) if c.date_str == ds and not c._other)


def test_deleted_task_stays_gray_in_calendar(env):
    db, w = env
    from ui.calendar_view import NativeCalendarDialog, TaskButton, DetailTaskRow
    ds = date.today().replace(day=15).isoformat()
    keep = db.add_task(w.window_id, '공문 회신', ds)
    gone = db.add_task(w.window_id, '협의회 자료', ds + ' 15:00')
    w._refresh_tasks()
    w._delete_task(next(t for t in db.get_tasks(w.window_id) if t['id'] == gone))   # 업무 목록에서 지우기
    cal = NativeCalendarDialog([w], None)
    cal.resize(900, 700)
    cal.show()
    _APP.processEvents()
    cell = _cell(cal, ds)
    assert [(t['name'], t['_source']) for t in cell._tasks] == [('공문 회신', 'tasks'), ('협의회 자료', 'history')]
    btns = {b.full_text: b for b in cell.findChildren(TaskButton)}
    assert btns['협의회 자료']._past and not btns['공문 회신']._past

    cal._on_cell_clicked(ds)                                            # 그날 목록: 지운 업무는 보기 전용
    rows = cal._detail_dlg.findChildren(DetailTaskRow)
    assert len(rows) == 2
    past_row = rows[1]
    assert past_row.findChildren(type(cal._detail_dlg._panel._count_label)) and \
        any(l.text() == '지난 업무' for l in past_row.findChildren(type(cal._detail_dlg._panel._count_label)))
    assert not [b for b in past_row.findChildren(__import__('PyQt5.QtWidgets', fromlist=['QPushButton']).QPushButton)]
    assert cal._detail_dlg._panel._count_label.text() == '할 일 1개 · 지난 업무 1개'

    hist = next(h for h in db.get_all_task_history() if h['name'] == '협의회 자료')
    db.delete_task_history(hist['id'])                                  # 지난 기록에서 지우면 달력에서도 사라짐
    cal._on_task_changed()
    assert [t['name'] for t in _cell(cal, ds)._tasks] == ['공문 회신']
    assert keep in [t['id'] for t in db.get_tasks(w.window_id)]
    cal.close()


def test_calendar_delete_also_keeps_history(env):
    db, w = env
    from ui.calendar_view import NativeCalendarDialog
    ds = date.today().replace(day=20).isoformat()
    tid = db.add_task(w.window_id, '수행평가 채점', ds)
    cal = NativeCalendarDialog([w], None)
    cal.show()
    _APP.processEvents()
    cal._on_cell_clicked(ds)
    task = next(t for t in cal._task_data[ds] if t['id'] == tid)
    cal._detail_dlg._panel._on_delete_task(task)                         # 달력에서 지워도
    assert tid not in [t['id'] for t in db.get_tasks(w.window_id)]
    assert [(t['name'], t['_source']) for t in db.get_calendar_items(ds)] == [('수행평가 채점', 'history')]
    assert [(t['name'], t['_source']) for t in _cell(cal, ds)._tasks] == [('수행평가 채점', 'history')]
    cal.close()
