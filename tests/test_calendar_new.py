"""새 달력 (크림 디자인 · 옆 패널 · 📌 고정 · 켤 때 열기)."""
import os
import sys
import unittest.mock as um
from datetime import date, timedelta

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


@pytest.fixture
def env(tmp_path, monkeypatch):
    import db
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    yield db, app


def _cal(app):
    from ui.calendar_view import NativeCalendarDialog
    cal = NativeCalendarDialog([], None)
    cal.resize(1100, 720)
    cal.show()
    app.processEvents()
    return cal


def test_chip_kind():
    from ui.calendar_view import chip_kind
    t = date(2026, 11, 26)
    d = lambda n: (t + timedelta(days=n)).isoformat()
    assert chip_kind({'deadline': d(-1)}, t) == 'over'
    assert chip_kind({'deadline': d(-1), 'strikethrough': 1}, t) == 'done'
    assert chip_kind({'deadline': d(5), '_source': 'history'}, t) == 'done'
    assert chip_kind({'deadline': d(0)}, t) == 'soon'
    assert chip_kind({'deadline': d(3)}, t) == 'soon'
    assert chip_kind({'deadline': d(4)}, t) == 'norm'
    assert chip_kind({'deadline': d(9), 'priority': 1}, t) == 'star'
    assert chip_kind({'deadline': d(-2), 'priority': 1}, t) == 'over'     # 지난 건 별이어도 빨강


def test_header_layout_and_month_nav(env):
    db, app = env
    cal = _cal(app)
    t = date.today()
    assert cal._year_label.text() == str(t.year) and cal._month_title.text() == f'{t.month}월'
    c = cal._header._center
    mid = c.x() + c.width() / 2
    assert abs(mid - cal._header.width() / 2) <= 2                      # ‹ 11월 › 은 머리줄 가운데
    assert cal._header._left.geometry().right() < c.x()                 # 왼쪽(2026 · 오늘)과 안 겹침
    cal._next_month()
    nxt = (t.replace(day=1) + timedelta(days=32))
    assert cal._month_title.text() == f'{nxt.month}월' and cal._year_label.text() == str(nxt.year)
    cal._go_today()
    assert cal._month_title.text() == f'{t.month}월'
    cal.close()


def test_side_panel_and_info(env):
    db, app = env
    wid = db.get_first_window_id()
    today = date.today()
    db.add_task(wid, '오늘 할 일', today.isoformat(), priority=1)
    cal = _cal(app)
    assert cal._panel._kicker.text() == '오늘'
    assert cal._panel._count_label.text() == '할 일 1개'
    assert '업무 1' in cal._info_label.text()
    other = (today.replace(day=1) - timedelta(days=1))                   # 지난달 마지막 날
    cal._on_cell_clicked(other.isoformat())                              # 다른 달 칸 → 그 달로
    assert cal._month_title.text() == f'{other.month}월'
    assert cal._panel.current_date_str == other.isoformat()
    cal.close()


def test_pin_modes_saved_and_flags(env):
    db, app = env
    from PyQt5.QtCore import Qt
    from ui import calendar_view as cv
    cal = _cal(app)
    assert cal._pin == 'none' and cal.btn_pin.label == '고정'
    cal._apply_pin('top')
    assert cal.windowFlags() & Qt.WindowStaysOnTopHint and db.get_setting(cv.PIN_KEY) == 'top'
    assert cal.isVisible() and cal.btn_pin.label == '항상 위에'
    cal._apply_pin('desktop')
    assert cal.windowFlags() & Qt.WindowStaysOnBottomHint and cal._lock_frame
    assert not cal.btn_win_min.isVisible()                               # 바탕화면 위젯은 내리기 없음
    from ui.native_frame import hit_test
    from PyQt5.QtCore import QPoint
    assert hit_test(cal, QPoint(2, 200)) is None                         # 위치 · 크기 잠금
    cal.close()
    cal2 = _cal(app)                                                     # 다음에 열어도 그대로
    assert cal2._pin == 'desktop'
    cal2._apply_pin('none')
    assert not (cal2.windowFlags() & (Qt.WindowStaysOnTopHint | Qt.WindowStaysOnBottomHint)) and not cal2._lock_frame
    cal2.close()


def test_startup_setting(env):
    db, app = env
    from ui import calendar_view as cv
    assert not cv.startup_enabled()
    cv.set_startup(True)
    assert cv.startup_enabled()
    cv.set_startup(False)
    assert not cv.startup_enabled()


def test_text_size_split_task_and_other(env):
    """Aa: 「업무 글자」(칸 안 알약 · 패널 할 일 줄)와 「날짜 · 요일 글자」를 따로."""
    db, app = env
    from ui import calendar_view as cv
    from ui.calendar_view import TaskButton, DetailTaskRow
    from PyQt5.QtCore import QEvent
    wid = db.get_first_window_id()
    db.add_task(wid, '글자 크기 확인', date.today().isoformat())
    cal = _cal(app)
    assert cal.btn_size.geometry().left() > cal.btn_pin.geometry().left()   # 📌 고정 오른쪽

    def chip_pt():
        app.sendPostedEvents(None, QEvent.DeferredDelete)
        return next(b for b in cal.findChildren(TaskButton) if b.full_text == '글자 크기 확인').font().pointSizeF()

    def row_name_pt():
        app.sendPostedEvents(None, QEvent.DeferredDelete)
        app.processEvents()
        return next(r for r in cal._panel.findChildren(DetailTaskRow) if not r.isHidden())._name_lbl.font().pointSizeF()

    chip0, row0 = chip_pt(), row_name_pt()
    date0, dow0 = cal._panel._date_label.font().pointSizeF(), cal._dow_labels[0].font().pointSizeF()
    month_pt = cal._month_title.font().pointSizeF()
    cal.set_task_scale(1.5)                                                   # 업무 글자만
    app.processEvents()
    assert abs(chip_pt() - chip0 * 1.5) < 0.6 and row_name_pt() > row0
    assert cal._panel._date_label.font().pointSizeF() == date0 and cal._dow_labels[0].font().pointSizeF() == dow0
    cal.set_text_scale(1.35)                                                  # 날짜 · 요일 글자만
    app.processEvents()
    assert cal._panel._date_label.font().pointSizeF() > date0 and cal._dow_labels[0].font().pointSizeF() > dow0
    assert abs(chip_pt() - chip0 * 1.5) < 0.6                                 # 업무 글자는 그대로
    assert cal._month_title.font().pointSizeF() == month_pt                  # 머리줄은 그대로
    assert db.get_setting(cv.TASK_SCALE_KEY) == '1.5' and db.get_setting(cv.TEXT_SCALE_KEY) == '1.35'
    cal.close()
    cal2 = _cal(app)                                                          # 다음에 열어도 그대로
    assert abs(cv._TASK_SCALE[0] - 1.5) < 0.01 and abs(cv._TEXT_SCALE[0] - 1.35) < 0.01
    cal2.set_task_scale(9.9)                                                  # 없는 크기는 무시
    assert abs(cv._TASK_SCALE[0] - 1.5) < 0.01
    for v in (1.1, 1.2, 1.0):                                                 # 여러 번 바꿔도 오류 없음
        cal2.set_task_scale(v)
        cal2.set_text_scale(v)
        app.processEvents()
    cal2.close()


def test_task_scale_inherits_old_setting(env):
    """예전 버전에서 고른 달력 글자 크기는 업무 글자에도 그대로 이어진다."""
    db, app = env
    from ui import calendar_view as cv
    db.set_setting(cv.TEXT_SCALE_KEY, '1.2')
    assert cv.load_task_scale() == 1.2 and cv.load_text_scale() == 1.2
