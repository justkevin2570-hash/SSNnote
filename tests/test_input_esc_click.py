"""업무 추가 칸: Esc로 빠져나오기 · 칸을 쓰다가 업무 줄의 별을 한 번에 누르기 (마감일 줄이 접히며 클릭이 엇나가던 문제)."""
import sys
import time
from unittest import mock

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
    wid = db.get_first_window_id()
    db.add_task(wid, '첫 업무', '')
    db.add_task(wid, '둘째 업무', '')
    import window
    w = window.MemoWindow(wid)
    w.resize(420, 480)
    w.show()
    w.activateWindow()
    QApplication.setActiveWindow(w)
    app.processEvents()
    yield w, app, db
    try:
        w.close()
    except RuntimeError:          # 테스트에서 이미 지운 창
        pass


def _pump(app, ms=0):
    end = time.monotonic() + ms / 1000.0
    app.processEvents()
    while time.monotonic() < end:
        app.processEvents()
        time.sleep(0.01)
    app.processEvents()


def _focus_name(w, app):
    from PyQt5.QtWidgets import QApplication
    QApplication.setActiveWindow(w)
    w.input_name.setFocus()
    _pump(app)
    assert w.date_row_widget.isVisible()


def _key(w, target, key):
    from PyQt5.QtTest import QTest
    QTest.keyClick(target, key)


def _rows(w):
    from ui.task_row import TaskRow
    return [r for r in w.task_list_widget.findChildren(TaskRow) if r.isVisible()]


# ── Esc ──────────────────────────────────────────────────────────
def test_esc_leaves_empty_input_and_folds_row(win):
    from PyQt5.QtCore import Qt
    w, app, db = win
    _focus_name(w, app)
    _key(w, w.input_name, Qt.Key_Escape)
    _pump(app)
    assert not w.input_name.hasFocus()
    assert not w.date_row_widget.isVisible()            # 아무것도 안 썼으면 마감일 줄도 접힘


def test_esc_keeps_typed_text(win):
    from PyQt5.QtCore import Qt
    w, app, db = win
    _focus_name(w, app)
    w.input_name.setText('쓰던 업무')
    _key(w, w.input_name, Qt.Key_Escape)
    _pump(app)
    assert not w.input_name.hasFocus() and w.input_name.text() == '쓰던 업무'
    assert w.date_row_widget.isVisible()                # 글자가 남아 있으면 마감일 줄은 그대로


def test_esc_in_time_and_repeat_fields(win):
    from PyQt5.QtCore import Qt
    w, app, db = win
    for field in (w.input_recurrence, w.input_time):
        _focus_name(w, app)
        w.input_time.setEnabled(True)
        field.setFocus()
        _pump(app)
        target = app.focusWidget() or field
        _key(w, target, Qt.Key_Escape)
        _pump(app)
        assert not field.hasFocus() and not (app.focusWidget() and field.isAncestorOf(app.focusWidget()))


def test_esc_in_date_clears_then_leaves(win):
    from PyQt5.QtCore import Qt, QDate
    w, app, db = win
    _focus_name(w, app)
    w.input_date.setDate(QDate.currentDate())
    w.input_date.setFocus()
    _pump(app)
    _key(w, w.input_date, Qt.Key_Escape)                # 첫 Esc: 날짜 지우고 업무 칸으로 (예전 그대로)
    _pump(app)
    assert w.input_date.date() == QDate(1900, 1, 1) and w.input_name.hasFocus()
    w.input_date.setFocus()
    _pump(app)
    _key(w, w.input_date, Qt.Key_Escape)                # 비어 있으면: 빠져나오기
    _pump(app)
    assert not w.input_name.hasFocus() and not w.input_date.hasFocus()


def test_esc_then_alt_n_returns(win):
    """Esc로 빠져나온 뒤에도 업무창 단축키(Alt+N 업무 추가 칸)로 다시 들어갈 수 있다."""
    from PyQt5.QtCore import Qt
    w, app, db = win
    _focus_name(w, app)
    _key(w, w.input_name, Qt.Key_Escape)
    _pump(app)
    w.focus_add_task()
    _pump(app)
    assert w.input_name.hasFocus()


# ── 칸을 쓰다가 별 한 번에 누르기 ──────────────────────────────
def _click_star_real(w, app, row):
    """실제 마우스처럼: 누른 자리 그대로 뗀다 (누르는 동안 화면이 움직이면 뗄 때 별이 거기 없음)."""
    from PyQt5.QtCore import Qt, QPoint
    from PyQt5.QtTest import QTest
    star = row.btn_star
    wp = star.mapTo(w, star.rect().center())
    target = w.childAt(wp)
    lp = target.mapFrom(w, wp)
    with mock.patch('PyQt5.QtWidgets.QApplication.mouseButtons', return_value=Qt.LeftButton):
        QTest.mousePress(target, Qt.LeftButton, Qt.NoModifier, lp)
        _pump(app, 30)                                   # 누르는 동안 (포커스 이동 → 예전엔 여기서 마감일 줄이 접힘)
        assert w.date_row_widget.isVisible()            # 이제는 손을 뗄 때까지 접히지 않음
    target2 = w.childAt(wp)
    QTest.mouseRelease(target, Qt.LeftButton, Qt.NoModifier, target.mapFrom(w, wp))
    _pump(app)
    return target2


def test_star_click_works_while_typing(win):
    w, app, db = win
    _focus_name(w, app)
    row = _rows(w)[0]
    tid = row.task['id']
    before = next(t for t in db.get_tasks(w.window_id) if t['id'] == tid)['priority']
    _click_star_real(w, app, row)
    _pump(app, 200)                                      # 손을 뗀 뒤 → 마감일 줄 접힘
    after = next(t for t in db.get_tasks(w.window_id) if t['id'] == tid)['priority']
    assert after != before                               # 한 번에 눌림
    assert not w.date_row_widget.isVisible()


def test_star_click_repeated_no_crash(win):
    """칸 쓰기 → 별 누르기를 여러 번 반복해도 오류 없이 매번 눌림."""
    w, app, db = win
    for i in range(4):
        _focus_name(w, app)
        row = _rows(w)[i % 2]
        tid = row.task['id']
        before = next(t for t in db.get_tasks(w.window_id) if t['id'] == tid)['priority']
        _click_star_real(w, app, row)
        _pump(app, 150)
        after = next(t for t in db.get_tasks(w.window_id) if t['id'] == tid)['priority']
        assert after != before


def test_fold_waits_for_popup_menu(win):
    """메뉴 · 목록이 열린 동안에는 접지 않고, 닫힌 뒤에 접는다."""
    w, app, db = win
    _focus_name(w, app)
    sentinel = object()
    with mock.patch('PyQt5.QtWidgets.QApplication.activePopupWidget', return_value=sentinel):
        w.input_name.clearFocus()
        w._sync_input_details()
        _pump(app, 150)
        assert w.date_row_widget.isVisible()
    _pump(app, 200)
    assert not w.date_row_widget.isVisible()


def test_pending_fold_after_close_no_crash(win):
    """접기를 기다리는 중에 창이 닫히고 지워져도 오류 · 튕김 없음."""
    from PyQt5.QtCore import Qt, QEvent
    w, app, db = win
    _focus_name(w, app)
    with mock.patch('PyQt5.QtWidgets.QApplication.mouseButtons', return_value=Qt.LeftButton):
        w.input_name.clearFocus()
        _pump(app)
    assert w._details_timer.isActive()
    w.close()
    w.deleteLater()
    app.sendPostedEvents(None, QEvent.DeferredDelete)
    _pump(app, 200)                                      # 타이머가 창과 함께 사라져 불리지 않음
