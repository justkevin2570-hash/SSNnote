"""오늘 한눈에 보기 창: 상단 숫자 칸은 글자가 가로 가운데, 업무 줄은 이름이 같은 자리에서 시작."""
import os
from datetime import date, datetime, timedelta

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


_APP = None      # 테스트가 끝나도 QApplication을 붙잡아 둔다 (움직임 시계(ui/motion)가 앱과 함께 지워지지 않게)


@pytest.fixture
def app(tmp_path, monkeypatch):
    global _APP
    import db
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    return _APP


NOW = datetime(2026, 11, 9, 10, 0)


def dl(n, t=''):
    return (date(2026, 11, 9) + timedelta(days=n)).isoformat() + t


TASKS = [
    {'name': '교과협의회 계획서 제출', 'deadline': dl(-2), 'priority': 1},
    {'name': '수학경시대회 결과 보고', 'deadline': dl(0, ' 17:00'), 'related_no': '교무-12'},
    {'name': '교육과정 박람회 운영 계획', 'deadline': dl(2), 'priority': 1},
    {'name': '학생부 점검', 'deadline': '', 'priority': 1},
    {'name': '기한 없는 일반 업무 (목록에 안 나옴)', 'deadline': ''},
]


def make(app, tasks=TASKS, budget=None):
    import dashboard
    from ui.dashboard_dialog import DashboardDialog
    s = dashboard.build_summary(tasks, budget if budget is not None else [{'budget': 1000, 'spent': 560}], now=NOW)
    dlg = DashboardDialog(None, s)
    dlg.show()
    app.processEvents()
    return dlg


def test_stat_tiles_text_is_centered_in_each_tile(app):
    from PyQt5.QtCore import Qt
    dlg = make(app)
    assert [t.name_label.text() for t in dlg.tiles] == ['지난 마감', '오늘 마감', '이번 주']   # 공문 업무 칸은 없음
    for tile in dlg.tiles:
        for lb in (tile.value_label, tile.name_label):
            assert lb.alignment() & Qt.AlignHCenter                      # 가운데 정렬
            mid = lb.geometry().x() + lb.geometry().width() / 2
            assert abs(mid - tile.width() / 2) <= 1                       # 칸의 가로 가운데에 놓임
    dlg.close()


def test_rows_names_start_at_the_same_x(app):
    dlg = make(app)
    xs = {r.name_label.mapTo(dlg, r.name_label.rect().topLeft()).x() for r in dlg.rows}
    assert len(dlg.rows) == 4 and len(xs) == 1        # D-day·★이 있든 없든 이름 줄이 맞음
    dlg.close()


def test_budget_card_centered_and_level_color(app):
    from PyQt5.QtCore import Qt
    from ui.dashboard_dialog import _level_color
    from ui.calendar_view import P
    dlg = make(app)
    for lb in dlg.budget_labels:
        assert lb.alignment() & Qt.AlignHCenter
    dlg.close()
    assert _level_color(50) == P.PRIMARY and _level_color(80) == '#D97706' and _level_color(100) == P.DANGER


def test_empty_day_and_no_budget(app):
    dlg = make(app, tasks=[], budget=[])
    assert [t.value_label.text() for t in dlg.tiles] == ['0', '0', '0']
    assert dlg.rows == [] and not hasattr(dlg, 'budget_labels')
    dlg.close()
