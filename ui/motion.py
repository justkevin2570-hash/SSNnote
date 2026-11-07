"""ui/motion.py — 마감이 급한 업무를 움직임으로 알리기.

원칙
  · 급할수록 더 움직인다. 평소엔 가만히, 마감이 다가오면 D-day 표시가 숨쉬듯 빛나고 점점 빨라진다.
  · 마감 몇 시간 전부터는 남은 시간이 실시간으로 줄고, 줄 아래 막대가 차오른다.
  · 기한이 지나면 빨간 점이 심장 박동처럼 뛴다.
  · ★ 중요 업무는 한 단계 먼저 움직이기 시작하고, 왼쪽 세로선이 금색이다.
  · 번쩍임 없이 은은하게. 창이 가려지거나 접혀 있으면 그리지 않는다.

설정(메뉴 → 설정 → 급함 표시)
  느긋하게 / 기본 / 미리미리 / 직접 설정 / 움직임 끄기  — app_settings 'motion_profile'(JSON)

    u = urgency(task)            # 이 업무를 지금 어떻게 움직일지
    clock().register(widget, fn)  # fn(t초)를 초당 30번 — 등록한 위젯이 보일 때만
"""

import json
import math
import time
import weakref
from datetime import date, datetime

from PyQt5.QtCore import QObject, QTimer

SETTING_KEY = 'motion_profile'

# 보통 업무 숨쉬기 시작(마감 며칠 전), ★ 중요 업무 숨쉬기 시작, 남은 시간 실시간 표시(몇 시간 전, 0=끔)
PRESETS = {
    'relaxed': {'label': '느긋하게', 'normal_days': 0, 'star_days': 1, 'countdown_hours': 1},
    'default': {'label': '기본', 'normal_days': 0, 'star_days': 3, 'countdown_hours': 3},
    'early':   {'label': '미리미리', 'normal_days': 3, 'star_days': 7, 'countdown_hours': 6},
}
PRESET_ORDER = ('default', 'early', 'relaxed')
DAY_CHOICES = (0, 1, 3, 7, 14)          # 직접 설정에서 고를 수 있는 '며칠 전'
HOUR_CHOICES = (0, 1, 3, 6)              # 남은 시간 표시 시작 (0 = 끄기)

_cache = None


def load() -> dict:
    """{'preset', 'normal_days', 'star_days', 'countdown_hours', 'enabled'}"""
    global _cache
    if _cache is not None:
        return dict(_cache)
    cfg = {'preset': 'default', **{k: v for k, v in PRESETS['default'].items() if k != 'label'}}
    try:
        from db import get_setting
        raw = json.loads(get_setting(SETTING_KEY, '') or '{}')
        preset = raw.get('preset', 'default')
        if preset in PRESETS:
            cfg.update({k: v for k, v in PRESETS[preset].items() if k != 'label'}, preset=preset)
        elif preset in ('custom', 'off'):
            cfg['preset'] = preset
            for k in ('normal_days', 'star_days', 'countdown_hours'):
                if isinstance(raw.get(k), int) and raw[k] >= 0:
                    cfg[k] = raw[k]
    except Exception:
        pass
    cfg['enabled'] = cfg['preset'] != 'off'
    _cache = cfg
    return dict(cfg)


def save(preset: str, normal_days=None, star_days=None, countdown_hours=None):
    """프리셋 이름(default/early/relaxed/off) 또는 'custom' + 값 세 개."""
    global _cache
    data = {'preset': preset}
    if preset in ('custom', 'off'):
        cur = load()
        data.update(normal_days=cur['normal_days'] if normal_days is None else int(normal_days),
                    star_days=cur['star_days'] if star_days is None else int(star_days),
                    countdown_hours=cur['countdown_hours'] if countdown_hours is None else int(countdown_hours))
    from db import set_setting
    set_setting(SETTING_KEY, json.dumps(data))
    _cache = None


def describe(cfg: dict = None) -> str:
    cfg = cfg or load()
    if not cfg['enabled']:
        return '움직임 끔'

    def d(n):
        return '당일' if n == 0 else f'{n}일 전'
    h = cfg['countdown_hours']
    return (f"보통 {d(cfg['normal_days'])} · ★ {d(cfg['star_days'])} · "
            f"남은 시간 {str(h) + '시간 전' if h else '끔'}")


# ── 급한 정도 ─────────────────────────────────────────────────────

def _parse(deadline: str):
    deadline = (deadline or '').strip()
    if not deadline:
        return None, None
    try:
        if len(deadline) > 10:
            dt = datetime.strptime(deadline[:16], '%Y-%m-%d %H:%M')
            return dt.date(), dt
        return date.fromisoformat(deadline[:10]), None
    except ValueError:
        return None, None


def urgency(task: dict, now: datetime = None, cfg: dict = None) -> dict:
    """이 업무를 지금 어떻게 보여줄지.

    breathe:   숨쉬기 주기(초, 0 = 안 움직임) — 가까울수록 짧다
    overdue:   기한 지남 (빨간 점 박동)
    countdown: 남은 시간을 실시간으로 보여줄지, remaining: 남은 초, progress: 0~1 (막대)
    side_line: 왼쪽 세로선 ('' / 'orange' / 'gold')
    star_glow: ★ 별도 함께 반짝일지
    """
    now = now or datetime.now()
    cfg = cfg or load()
    out = {'breathe': 0.0, 'overdue': False, 'countdown': False, 'remaining': None,
           'progress': 0.0, 'side_line': '', 'star_glow': False, 'days_left': None}
    if task.get('strikethrough'):
        return out
    d, dt = _parse(task.get('deadline', ''))
    if d is None:
        return out
    starred = bool(task.get('priority'))
    days_left = (d - now.date()).days
    out['days_left'] = days_left
    if days_left < 0 or (dt is not None and dt <= now):
        out['overdue'] = True
        return out
    if not cfg['enabled']:
        return out

    start = cfg['star_days'] if starred else cfg['normal_days']
    if days_left <= start:
        # 당일 4초, 하루 전 5초 … 멀수록 느리게(최대 8초)
        out['breathe'] = 4.0 if days_left == 0 else min(8.0, 4.0 + days_left)
        out['star_glow'] = starred
    if dt is not None and days_left == 0:
        remaining = (dt - now).total_seconds()
        out['remaining'] = remaining
        window = cfg['countdown_hours'] * 3600
        if window and remaining <= window:
            out['countdown'] = True
            out['progress'] = max(0.0, min(1.0, 1 - remaining / window))
            if starred:
                out['side_line'] = 'gold'
        if remaining <= 3600:          # 마지막 한 시간: 빠르게 숨쉬고 왼쪽 세로선
            out['breathe'] = 2.0
            out['side_line'] = 'gold' if starred else 'orange'
    return out


def remaining_text(seconds: float) -> str:
    """남은 시간: 2시간 47분 / 47분 / 1분 미만."""
    m = int(seconds // 60)
    if m <= 0:
        return '1분 미만'
    h, m = divmod(m, 60)
    return f'{h}시간 {m}분' if h else f'{m}분'


# ── 움직임 값 ────────────────────────────────────────────────────

def breath(t: float, period: float) -> float:
    """0~1 부드러운 숨쉬기 (사인 제곱)."""
    if period <= 0:
        return 0.0
    return math.sin(math.pi * ((t % period) / period)) ** 2


def heartbeat(t: float, period: float = 1.6) -> float:
    """두 번 뛰고 쉬는 박동 0~1."""
    p = (t % period) / period
    def bump(c, w):
        x = (p - c) / w
        return math.exp(-x * x)
    return min(1.0, bump(0.08, 0.05) + 0.7 * bump(0.26, 0.05))


# ── 공용 시계 ────────────────────────────────────────────────────

class _Clock(QObject):
    """움직이는 위젯이 하나라도 보일 때만 초당 30번 깨어나 등록된 함수를 부른다."""

    def __init__(self):
        super().__init__()
        self._subs = {}                  # id → (weakref(owner), fn)
        self._t0 = time.monotonic()
        self._timer = QTimer(self)
        self._timer.setInterval(33)
        self._timer.timeout.connect(self._tick)

    def now(self) -> float:
        return time.monotonic() - self._t0

    def register(self, owner, fn):
        self._subs[id(owner)] = (weakref.ref(owner), fn)
        if not self._timer.isActive():
            self._timer.start()

    def unregister(self, owner):
        self._subs.pop(id(owner), None)
        if not self._subs:
            self._timer.stop()

    def _tick(self):
        t = self.now()
        shown = 0
        for key, (ref, fn) in list(self._subs.items()):
            w = ref()
            if w is None:
                self._subs.pop(key, None)
                continue
            try:
                if not w.isVisible() or w.window().isMinimized():
                    continue
                shown += 1
                fn(t)
            except RuntimeError:          # 이미 정리된 위젯
                self._subs.pop(key, None)
        if not self._subs:
            self._timer.stop()
            return
        # 보이는 게 없으면(트레이로 숨김 등) 0.5초에 한 번만 깨어나 다시 보이는지 확인
        want = 33 if shown else 500
        if self._timer.interval() != want:
            self._timer.setInterval(want)


_CLOCK = None


def clock() -> _Clock:
    global _CLOCK
    if _CLOCK is None:
        _CLOCK = _Clock()
    return _CLOCK
