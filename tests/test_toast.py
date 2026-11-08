"""마감 임박 알림(유주) — 급한 정도 판별과 나타났다 사라지기."""
import os
from datetime import date, timedelta

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


def d(n):
    return (date.today() + timedelta(days=n)).isoformat()


@pytest.mark.parametrize('task, level, label', [
    ({'deadline': d(0) + ' 23:59', 'notice': '2시간 후 마감'}, 'now', '2시간 후'),
    ({'deadline': d(0)}, 'now', 'D-day'),
    ({'deadline': d(1)}, 'soon', 'D-1'),
    ({'deadline': d(2)}, 'soon', 'D-2'),
    ({'deadline': d(4)}, 'later', 'D-4'),
])
def test_toast_level(task, level, label):
    from ui.chrome import toast_level
    assert toast_level(task) == (level, label)


@pytest.fixture(scope='module')
def app():
    from PyQt5.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def test_toast_sorts_limits_and_closes(app):
    from ui.chrome import UrgentToast
    tasks = [{'name': f'업무 {i}', 'deadline': d(4)} for i in range(5)] + [{'name': '급한 일', 'deadline': d(0)}]
    t = UrgentToast(tasks)
    assert t._level == 'now' and t._rows[0][2] == '급한 일'          # 가장 급한 것이 맨 위
    assert len(t._rows) == UrgentToast.MAX_ROWS and t._more == 2
    clock = [0.0]
    t._now = lambda: clock[0]
    t.show()
    t._t0 = t._last = 0.0
    closed = []
    t.destroyed.connect(lambda *_: closed.append(1))
    t._hover = True                                                  # 마우스가 올라가 있으면 기다림
    for _ in range(int(12 / 0.05)):
        clock[0] += 0.05
        t._tick()
    assert t._leave_t is None
    t._hover = False
    for _ in range(int(11 / 0.05)):
        clock[0] += 0.05
        if t._leave_t is not None and clock[0] - t._leave_t > 0.3:
            break
        t._tick()
    assert t._leave_t is not None
    t._tick()
    assert not t.isVisible()
