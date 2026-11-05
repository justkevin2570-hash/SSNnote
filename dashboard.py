"""dashboard.py — 「오늘 한눈에 보기」 집계 로직 (UI·DB 의존 없음).

    summary = build_summary(db.get_open_tasks(), db.get_budget_items())
"""

from datetime import date, datetime, timedelta

_WEEKDAYS = '월화수목금토일'


def _parse_deadline(deadline: str):
    """'YYYY-MM-DD' → (date, None), 'YYYY-MM-DD HH:MM' → (date, datetime). 못 읽으면 (None, None)."""
    if not deadline:
        return None, None
    try:
        if len(deadline) > 10:
            dt = datetime.strptime(deadline, '%Y-%m-%d %H:%M')
            return dt.date(), dt
        return date.fromisoformat(deadline), None
    except ValueError:
        return None, None


def dday_label(deadline: str, today: date) -> str:
    d, _ = _parse_deadline(deadline)
    if d is None:
        return ''
    diff = (d - today).days
    return 'D-day' if diff == 0 else (f'D-{diff}' if diff > 0 else f'D+{-diff}')


def korean_date(d: date) -> str:
    """date(2026, 11, 5) → '11월 5일 (목)'."""
    return f'{d.month}월 {d.day}일 ({_WEEKDAYS[d.weekday()]})'


def build_summary(tasks, budget_items=(), now: datetime = None, week_days: int = 7) -> dict:
    """열린 업무를 지난 마감 / 오늘 / 이번 주(내일~7일)로 나누고 예산 집행률을 계산한다.

    시각이 있는 업무는 그 시각이 지나면 '지난 마감', 날짜만 있는 업무는 다음 날부터 '지난 마감'.
    """
    now = now or datetime.now()
    today = now.date()
    overdue, due_today, week, important = [], [], [], []
    doc_open = 0
    for t in tasks:
        if t.get('strikethrough'):
            continue
        if (t.get('related_no') or '').strip():
            doc_open += 1
        d, dt = _parse_deadline(t.get('deadline') or '')
        item = dict(t)
        item['dday'] = dday_label(t.get('deadline') or '', today)
        item['time'] = dt.strftime('%H:%M') if dt else ''
        if d is None:
            if t.get('priority'):
                important.append(item)
            continue
        if (dt is not None and dt < now) or d < today:
            overdue.append(item)
        elif d == today:
            due_today.append(item)
        elif d <= today + timedelta(days=week_days):
            week.append(item)

    def key(x):
        return (x.get('deadline') or '', -int(x.get('priority') or 0), x.get('name') or '')
    overdue.sort(key=key)
    due_today.sort(key=key)
    week.sort(key=key)
    important.sort(key=lambda x: x.get('name') or '')

    budget = None
    items = list(budget_items or ())
    if items:
        allocated = sum(int(i.get('budget') or 0) for i in items)
        spent = sum(int(i.get('spent') or 0) for i in items)
        budget = {
            'allocated': allocated,
            'spent': spent,
            'remaining': allocated - spent,
            'pct': (spent / allocated * 100) if allocated > 0 else 0.0,
        }

    return {
        'date_label': korean_date(today),
        'overdue': overdue,
        'today': due_today,
        'week': week,
        'important': important,
        'doc_open': doc_open,
        'budget': budget,
        'headline': headline(len(overdue), len(due_today), len(week)),
    }


def headline(n_overdue: int, n_today: int, n_week: int) -> str:
    """한 줄 요약 — 가장 급한 것부터."""
    if n_overdue:
        return f'마감이 지난 업무가 {n_overdue}건 있어요. 먼저 확인하세요.'
    if n_today:
        return f'오늘 마감 {n_today}건' + (f', 이번 주 {n_week}건이 남았어요.' if n_week else '이 있어요.')
    if n_week:
        return f'오늘 마감은 없어요. 이번 주 마감 {n_week}건이 있어요.'
    return '이번 주 마감 업무가 없어요. 여유로운 한 주!'


def format_won(n: int) -> str:
    """1234567 → '123만 원', 12345678 → '1,235만 원', 9000 → '9,000원'."""
    n = int(n)
    sign = '-' if n < 0 else ''
    n = abs(n)
    if n >= 10000:
        return f'{sign}{round(n / 10000):,}만 원'
    return f'{sign}{n:,}원'
