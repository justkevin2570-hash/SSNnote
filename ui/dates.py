"""ui/dates.py — 날짜 계산 로직 (D-day, 반복 일정 다음 마감일). UI 의존 없음."""
from datetime import date, datetime, timedelta

def calc_dday(deadline_str):
    try:
        if len(deadline_str) > 10:  # 'yyyy-MM-dd HH:MM'
            deadline_dt = datetime.strptime(deadline_str, '%Y-%m-%d %H:%M')
            diff_days = (deadline_dt.date() - date.today()).days
            diff_seconds = (deadline_dt - datetime.now()).total_seconds()
        else:  # 'yyyy-MM-dd'
            diff_days = (date.fromisoformat(deadline_str) - date.today()).days
            diff_seconds = None
    except ValueError:
        return '날짜오류'

    if diff_days > 0:
        return f'D-{diff_days}'
    elif diff_days < 0:
        return f'D+{abs(diff_days)}'
    else:
        return 'D-day'


def _next_recurrence_deadline(deadline_str, recurrence):
    import calendar as _cal
    try:
        base = date.fromisoformat(deadline_str[:10])
    except ValueError:
        return ''

    next_date = None
    if recurrence == 'weekly':
        next_date = base + timedelta(weeks=1)
    elif recurrence == 'biweekly':
        next_date = base + timedelta(weeks=2)
    elif recurrence == 'monthly':
        year = base.year + (base.month // 12)
        month = (base.month % 12) + 1
        day = min(base.day, _cal.monthrange(year, month)[1])
        next_date = date(year, month, day)
    elif recurrence == 'yearly':
        year = base.year + 1
        day = min(base.day, _cal.monthrange(year, base.month)[1])
        next_date = date(year, base.month, day)
    elif recurrence.startswith('custom:'):
        # custom:month|week|weekday (month 0=매월, 1~12 / week 1~4 / weekday 0=월~6=일)
        try:
            _, parts = recurrence.split(':')
            m_target, w_target, wd_target = map(int, parts.split('|'))
            
            # 기준일 다음날부터 탐색
            curr = base + timedelta(days=1)
            found = False
            # 최대 12년치 탐색 (특정 월 반복의 경우 고려)
            for _ in range(365 * 12):
                # 월 체크
                if m_target != 0 and curr.month != m_target:
                    # 월이 다르면 다음 달 1일로 점프 (최적화)
                    if curr.month == 12:
                        curr = date(curr.year + 1, 1, 1)
                    else:
                        curr = date(curr.year, curr.month + 1, 1)
                    continue
                
                # 요일 체크
                if curr.weekday() == wd_target:
                    # 몇째주인지 계산 (해당 월의 n번째 요일)
                    # 1일의 요일을 구해서 계산
                    first_day_of_month = date(curr.year, curr.month, 1)
                    first_wd = first_day_of_month.weekday()
                    # curr.day가 해당 월의 몇 번째 wd_target 인지 계산
                    # 예: 1일이 월(0)이고 8일도 월(0)이면, (8-1)//7 + 1 = 2번째
                    nth = (curr.day - 1) // 7 + 1
                    
                    if nth == w_target:
                        next_date = curr
                        found = True
                        break
                
                curr += timedelta(days=1)
            if not found:
                return ''
        except:
            return ''
    else:
        return ''
    
    if not next_date:
        return ''
        
    time_suffix = deadline_str[10:] if len(deadline_str) > 10 else ''
    return next_date.isoformat() + time_suffix
