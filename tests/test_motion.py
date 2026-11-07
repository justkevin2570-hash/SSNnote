"""급함 표시 규칙 (ui/motion.py)."""
from datetime import datetime

import pytest

pytest.importorskip('PyQt5')
from ui import motion  # noqa: E402

NOW = datetime(2026, 11, 7, 14, 0)
DEFAULT = {'preset': 'default', 'normal_days': 0, 'star_days': 3, 'countdown_hours': 3, 'enabled': True}
EARLY = {'preset': 'early', 'normal_days': 3, 'star_days': 7, 'countdown_hours': 6, 'enabled': True}
OFF = {**DEFAULT, 'preset': 'off', 'enabled': False}


def u(deadline, priority=0, cfg=DEFAULT, done=0):
    return motion.urgency({'deadline': deadline, 'priority': priority, 'strikethrough': done}, NOW, cfg)


def test_no_deadline_or_far_away_is_still():
    assert u('')['breathe'] == 0 and not u('')['overdue']
    assert u('2026-11-20')['breathe'] == 0
    assert u('2026-11-20', priority=1)['breathe'] == 0                 # 날짜 없는/먼 ★도 움직이지 않음
    assert u('', priority=1)['breathe'] == 0


def test_today_breathes_and_starred_starts_earlier():
    assert u('2026-11-07')['breathe'] == 4.0
    assert u('2026-11-09')['breathe'] == 0                             # 보통 업무는 당일부터
    s = u('2026-11-09', priority=1)                                    # ★는 3일 전부터, 멀수록 느리게
    assert s['breathe'] == 6.0 and s['star_glow']
    assert u('2026-11-10', priority=1)['breathe'] == 7.0
    assert u('2026-11-11', priority=1)['breathe'] == 0


def test_countdown_progress_and_last_hour():
    a = u('2026-11-07 16:30')                                          # 2시간 30분 남음
    assert a['countdown'] and abs(a['progress'] - (1 - 2.5 / 3)) < 1e-6 and a['side_line'] == ''
    assert motion.remaining_text(a['remaining']) == '2시간 30분'
    b = u('2026-11-07 14:40')                                          # 40분 남음
    assert b['breathe'] == 2.0 and b['side_line'] == 'orange'
    assert motion.remaining_text(b['remaining']) == '40분'
    c = u('2026-11-07 16:30', priority=1)
    assert c['side_line'] == 'gold'                                    # ★는 남은 시간 표시부터 금색 선
    assert not u('2026-11-07 18:00')['countdown']                      # 4시간 전엔 아직


def test_overdue_and_done():
    assert u('2026-11-06')['overdue']
    assert u('2026-11-07 13:30')['overdue']                            # 오늘인데 시각이 지남
    assert u('2026-11-06', done=1) == u('', done=1)                    # 완료 표시하면 조용


def test_presets_and_off():
    assert u('2026-11-10', cfg=EARLY)['breathe'] == 7.0                # 미리미리: 보통도 3일 전부터
    off = u('2026-11-07 14:30', cfg=OFF)
    assert off['breathe'] == 0 and not off['countdown']
    assert u('2026-11-06', cfg=OFF)['overdue']                         # 기한 지남 판정은 그대로(박동만 꺼짐)


def test_wave_shapes():
    assert motion.breath(0, 4) == 0 and abs(motion.breath(2, 4) - 1) < 1e-9
    assert motion.breath(1, 0) == 0
    assert 0 <= motion.heartbeat(0.5) <= 1 and motion.heartbeat(0.128) > 0.9


def test_describe():
    assert motion.describe(DEFAULT) == '보통 당일 · ★ 3일 전 · 남은 시간 3시간 전'
    assert motion.describe(OFF) == '움직임 끔'
