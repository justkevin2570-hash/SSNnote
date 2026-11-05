from datetime import datetime

from dashboard import build_summary, dday_label, headline, format_won, korean_date

NOW = datetime(2026, 11, 5, 14, 0)   # 목요일 오후 2시


def T(name, deadline='', **kw):
    return {'id': hash(name) % 1000, 'window_id': 1, 'name': name, 'deadline': deadline,
            'strikethrough': 0, 'priority': 0, 'related_no': '', **kw}


def test_buckets_and_time_rules():
    tasks = [
        T('어제 마감', '2026-11-04'),
        T('오늘 오전 9시(지남)', '2026-11-05 09:00'),
        T('오늘 날짜만', '2026-11-05'),
        T('오늘 오후 5시', '2026-11-05 17:00'),
        T('내일', '2026-11-06'),
        T('7일 뒤', '2026-11-12'),
        T('8일 뒤', '2026-11-13'),
        T('완료됨', '2026-11-05', strikethrough=1),
        T('중요·기한없음', '', priority=1),
        T('기한없음', ''),
        T('공문 업무', '2026-11-20', related_no='교육과-123'),
        T('날짜 오류', '2026-13-40'),
    ]
    s = build_summary(tasks, now=NOW)
    names = lambda k: [t['name'] for t in s[k]]  # noqa: E731
    assert names('overdue') == ['어제 마감', '오늘 오전 9시(지남)']
    assert names('today') == ['오늘 날짜만', '오늘 오후 5시']
    assert names('week') == ['내일', '7일 뒤']
    assert names('important') == ['중요·기한없음']
    assert s['doc_open'] == 1
    assert s['today'][1]['time'] == '17:00' and s['week'][0]['dday'] == 'D-1'
    assert s['date_label'] == '11월 5일 (목)'
    assert s['budget'] is None


def test_budget():
    s = build_summary([], [{'budget': 3_000_000, 'spent': 1_200_000}, {'budget': 1_000_000, 'spent': 0}], now=NOW)
    assert s['budget'] == {'allocated': 4_000_000, 'spent': 1_200_000, 'remaining': 2_800_000, 'pct': 30.0}
    assert build_summary([], [{'budget': 0, 'spent': 5}], now=NOW)['budget']['pct'] == 0.0


def test_labels():
    assert dday_label('2026-11-05 09:00', NOW.date()) == 'D-day'
    assert dday_label('2026-11-01', NOW.date()) == 'D+4'
    assert dday_label('bad', NOW.date()) == ''
    assert headline(2, 1, 1).startswith('마감이 지난 업무가 2건')
    assert headline(0, 2, 3) == '오늘 마감 2건, 이번 주 3건이 남았어요.'
    assert headline(0, 1, 0) == '오늘 마감 1건이 있어요.'
    assert headline(0, 0, 0).startswith('이번 주 마감 업무가 없어요')
    assert format_won(12_345_678) == '1,235만 원' and format_won(9000) == '9,000원'
    assert format_won(-50_000) == '-5만 원'
    assert korean_date(NOW.date().replace(day=8)) == '11월 8일 (일)'
