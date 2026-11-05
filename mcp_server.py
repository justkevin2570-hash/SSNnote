"""mcp_server.py — 서서니 노트 MCP 서버 (Claude Code · Claude 데스크톱 앱 연결용).

Claude가 서서니 노트의 업무 목록을 보고, 업무를 추가·수정·완료 표시하고,
학교 정보와 공문 서식을 참고해 공문 초안을 쓸 수 있게 한다.
사용자는 자기 Claude 구독으로 Claude 앱을 쓰고, 이 서버는 PC 안의 서서니 노트 DB만 다룬다
(API 키 불필요, 데이터는 PC 밖으로 나가지 않음 — Claude에게 보여주는 내용만 대화에 들어감).

등록 (Claude Code):
    claude mcp add ssnnote -- python "C:\\경로\\SSNnote\\mcp_server.py"
실행 중인 메모 창에는 3초 안에 반영된다 (app_settings 'external_change' 신호).

주의: stdio로 통신하므로 이 파일에서 print()로 표준출력을 쓰면 안 된다.
"""

from datetime import date, datetime
from typing import Literal, Optional

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

import dashboard
import db
from korean_search import best_score
from school_profile import prompt_block

EXTERNAL_CHANGE_KEY = 'external_change'

server = MCPServer(
    name='ssnnote',
    title='서서니 노트',
    instructions=(
        '서서니 노트는 한국 고등학교 교사의 업무·마감 관리 메모 앱입니다. '
        '업무 질문에는 today_overview / list_tasks / search_tasks로 실제 목록을 먼저 확인하세요. '
        '업무를 추가·수정하기 전에 이름과 마감일을 사용자에게 한 줄로 확인받고, '
        '날짜는 today_overview의 오늘 날짜를 기준으로 YYYY-MM-DD로 바꿔 넣으세요. '
        '공문·계획서 초안을 쓸 때는 official_document_guide의 서식과 학교 정보를 따르세요.'
    ),
)

_READ = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
_WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)


# ── 내부 도우미 ────────────────────────────────────────────────

def _deadline(day: str, time: str) -> str:
    """('2026-11-14', '17:00') → '2026-11-14 17:00'. 형식이 틀리면 ToolError (Claude에게 보이는 한국어 안내)."""
    day, time = (day or '').strip(), (time or '').strip()
    if not day:
        if time:
            raise ToolError('시각만 넣을 수는 없습니다. 날짜(YYYY-MM-DD)도 함께 주세요.')
        return ''
    try:
        date.fromisoformat(day)
    except ValueError:
        raise ToolError(f"날짜 형식이 올바르지 않습니다: '{day}' (예: 2026-11-14)")
    if time:
        try:
            datetime.strptime(time, '%H:%M')
        except ValueError:
            raise ToolError(f"시각 형식이 올바르지 않습니다: '{time}' (예: 17:00)")
        return f'{day} {time}'
    return day


def _view(t: dict, today: date) -> dict:
    """Claude에게 보여줄 업무 요약 (내부 열 이름 대신 의미 있는 이름)."""
    dl = t.get('deadline') or ''
    notes = (t.get('notes') or '').strip()
    return {
        'task_id': t['id'],
        'name': t.get('name', ''),
        'deadline': dl[:10],
        'time': dl[11:16] if len(dl) > 10 else '',
        'dday': dashboard.dday_label(dl, today),
        'important': bool(t.get('priority')),
        'done': bool(t.get('strikethrough')),
        'related_doc': t.get('related_no', ''),
        'recurrence': t.get('recurrence', ''),
        'notes': notes if len(notes) <= 400 else notes[:400] + '…',
    }


def _all_tasks():
    with db._connect() as conn:
        return [dict(r) for r in conn.execute('SELECT * FROM tasks').fetchall()]


def _signal_app():
    """실행 중인 서서니 노트 창이 다시 그리도록 알린다."""
    db.set_setting(EXTERNAL_CHANGE_KEY, datetime.now().isoformat(timespec='microseconds'))


def _task_list_window_id() -> int:
    """업무 목록을 보여주는 창(첫 번째 창)."""
    wid = db.get_first_window_id()
    if wid is None:
        wid = db.create_window()
    return wid


# ── 읽기 ──────────────────────────────────────────────────────

@server.tool(title='오늘 한눈에 보기', annotations=_READ)
def today_overview() -> dict:
    """오늘 날짜, 한 줄 요약, 지난 마감·오늘 마감·이번 주(7일) 업무, 중요 표시 업무, 공문 업무 수, 예산 집행률.

    업무 상황을 물으면 먼저 이걸 부르세요. 날짜 계산의 기준(오늘)도 여기서 확인합니다.
    """
    now = datetime.now()
    s = dashboard.build_summary(db.get_open_tasks(), db.get_budget_items(), now=now)
    today = now.date()
    return {
        'today': today.isoformat(),
        'today_label': s['date_label'],
        'headline': s['headline'],
        'overdue': [_view(t, today) for t in s['overdue']],
        'due_today': [_view(t, today) for t in s['today']],
        'this_week': [_view(t, today) for t in s['week']],
        'important_no_deadline': [_view(t, today) for t in s['important']],
        'open_official_doc_tasks': s['doc_open'],
        'budget': s['budget'],
    }


@server.tool(title='업무 목록', annotations=_READ)
def list_tasks(status: Literal['open', 'done', 'all'] = 'open', limit: int = 50) -> dict:
    """업무 목록을 마감일 순으로 돌려줍니다. status: open(진행 중, 기본) / done(완료 표시) / all."""
    today = date.today()
    tasks = _all_tasks()
    if status == 'open':
        tasks = [t for t in tasks if not t.get('strikethrough')]
    elif status == 'done':
        tasks = [t for t in tasks if t.get('strikethrough')]
    tasks.sort(key=lambda t: (t.get('deadline') or '9999', -int(t.get('priority') or 0), t.get('name') or ''))
    limit = max(1, min(int(limit), 200))
    return {'today': today.isoformat(), 'count': len(tasks),
            'tasks': [_view(t, today) for t in tasks[:limit]]}


@server.tool(title='업무 찾기', annotations=_READ)
def search_tasks(query: str, include_done: bool = False) -> dict:
    """업무 이름·메모·관련 공문번호로 찾습니다. 한글 초성 검색도 됩니다 (예: 'ㄱㅁㅎㅇ' → 교무회의)."""
    today = date.today()
    scored = []
    for t in _all_tasks():
        if t.get('strikethrough') and not include_done:
            continue
        sc = best_score(query, t.get('name', ''), t.get('related_no', ''), (t.get('notes') or '')[:200])
        if sc > 0:
            scored.append((sc, t))
    scored.sort(key=lambda x: -x[0])
    return {'query': query, 'tasks': [_view(t, today) for _, t in scored[:20]]}


@server.tool(title='공문 작성 안내', annotations=_READ)
def official_document_guide() -> dict:
    """공문·계획서 초안을 쓰기 전에 부르세요. 학교 공문 서식 규칙과 저장된 학교 정보(학교명·부서·담당자 등)."""
    return {
        'format_rules': [
            "본문 구조: '1. 관련: [관련 문서번호(날짜)]' → '2. [제목]을/를 다음과 같이 [운영/보고/제출]하고자 합니다.' → 가. 나. 다. 항목 → '붙임  [파일명] 1부.  끝.'",
            "날짜는 '2026. 3. 12.' 형식 (숫자 뒤 마침표, 요일은 '(목)').",
            '가.나.다. 항목은 앞에 공백 2칸, 1) 2) 3) 항목은 공백 4칸.',
            "문체는 개조식, '~함', '~바람', '~하고자 합니다'.",
            "마지막은 반드시 '끝.'으로 맺음. 붙임이 없으면 본문 끝에 '끝.'.",
            '마크다운(**, ##) 쓰지 않음. 공문에 없는 수치·명단은 지어내지 말고 ○○로 비워 둠.',
        ],
        'school_profile': prompt_block() or '(학교 정보 없음 — 서서니 노트 메뉴 → 🏫 학교 정보 설정에서 입력)',
    }


# ── 쓰기 ──────────────────────────────────────────────────────

@server.tool(title='업무 추가', annotations=_WRITE)
def add_task(name: str, deadline: str = '', time: str = '', notes: str = '',
             related_doc: str = '', important: bool = False) -> dict:
    """서서니 노트 업무 목록에 업무를 추가합니다.

    deadline: 'YYYY-MM-DD' (없으면 빈 값), time: 'HH:MM' (선택), related_doc: 관련 공문번호 (선택),
    notes: 업무 메모 (요약·할 일 등, 선택). 추가 전에 이름과 마감을 사용자에게 확인받으세요.
    """
    name = (name or '').strip()
    if not name:
        raise ToolError('업무 이름이 비어 있습니다.')
    dl = _deadline(deadline, time)
    tid = db.add_task(_task_list_window_id(), name, dl, priority=1 if important else 0,
                      notes=notes.strip(), related_no=related_doc.strip())
    _signal_app()
    return {'added': _view(db.get_task(tid), date.today())}


@server.tool(title='업무 수정', annotations=_WRITE)
def update_task(task_id: int, name: Optional[str] = None, deadline: Optional[str] = None,
                time: Optional[str] = None, important: Optional[bool] = None,
                done: Optional[bool] = None, notes: Optional[str] = None,
                related_doc: Optional[str] = None) -> dict:
    """업무를 고칩니다. 바꿀 항목만 넣으세요 (나머지는 그대로).

    deadline='' 이면 마감 없음으로, done=true 면 완료 표시(취소선), done=false 면 다시 진행 중.
    task_id는 list_tasks / search_tasks / today_overview 결과에서 확인합니다. 삭제는 지원하지 않습니다.
    """
    t = db.get_task(task_id)
    if t is None:
        raise ToolError(f'업무 번호 {task_id}를 찾을 수 없습니다. list_tasks로 다시 확인하세요.')
    new_name = t['name'] if name is None else name.strip()
    if not new_name:
        raise ToolError('업무 이름을 비울 수 없습니다.')
    old = t.get('deadline') or ''
    if deadline is None and time is None:
        new_dl = old
    else:
        day = old[:10] if deadline is None else deadline
        tm = (old[11:16] if len(old) > 10 else '') if time is None else time
        new_dl = _deadline(day, tm if day else '')
    db.update_task(task_id, new_name, new_dl,
                   strikethrough=int(t.get('strikethrough', 0) if done is None else bool(done)),
                   priority=int(t.get('priority', 0) if important is None else bool(important)),
                   recurrence=t.get('recurrence', ''))
    if notes is not None:
        db.set_task_notes(task_id, notes.strip())
    if related_doc is not None:
        db.set_task_related_no(task_id, related_doc.strip())
    _signal_app()
    return {'updated': _view(db.get_task(task_id), date.today())}


def main():
    db.init_db()   # 앱을 한 번도 안 켠 PC에서도 동작하도록
    server.run('stdio')


if __name__ == '__main__':
    main()
