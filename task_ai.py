"""task_ai.py — 업무창 옆 AI 대화 창에서 AI에게 묻고 답 듣기 (UI 없음 + 뒤에서 실행할 일꾼).

  · 모든 업무창의 업무(이름·마감·중요·반복·완료)를 함께 보내 질문에 답한다.
  · "다음 주 금요일까지 공문 회신 추가해줘", "협의회 자료 마감을 수요일 3시로", "OO 완료 표시해줘"처럼 말하면
    AI는 무엇을 바꿀지만 정한다. 날짜·업무 확인은 앱이 한다 (parse_reply). 실제 반영은 사용자가 [넣기]를 눌러야 한다.
  · 삭제는 하지 않는다 (잘못 지우면 되돌리기 어렵다).
  · 개인정보 보호 모드면 보내는 글의 개인정보를 가린다 (school_profile.mask_for_ai).

응답 형식 (JSON):
  {"answer": "답",
   "add": [{"name": "업무 이름", "deadline": "2026-11-21 15:00", "important": false}],
   "set_deadline": [{"task_id": 12, "deadline": "2026-11-18"}],
   "done": [{"task_id": 12}],
   "important": [{"task_id": 12, "on": true}]}
"""

import json
import re
from datetime import date, datetime

_WEEKDAYS = '월화수목금토일'
NAME_MAX = 80                    # 새 업무 이름 글자 수 한도
ADD_MAX = 10                     # 한 번에 추가할 수 있는 업무 수

SYSTEM = (
    "당신은 한국 고등학교 선생님의 업무 메모(할 일 목록)를 돕는 도우미입니다. 아래 업무 목록만 근거로 답합니다.\n"
    "- 답은 짧고 분명하게, 목록이 필요하면 '- '로 시작하는 줄로. 마크다운(**, ##, 표)은 쓰지 않습니다.\n"
    "- 날짜는 '11월 21일(금)', 시각은 '오후 3시'처럼 씁니다. 목록에 없는 업무는 지어내지 않습니다.\n"
    "- '내일', '다음 주 금요일' 같은 말은 위에 적힌 오늘 날짜를 기준으로 계산합니다.\n"
    "- 새 업무를 넣어 달라고 하면 add에 이름(name), 마감(deadline: 'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM', 없으면 빈 글),\n"
    "  중요(important: 참/거짓)를 넣습니다. 시각을 말하지 않았으면 날짜만 씁니다.\n"
    "- 마감을 바꾸라고 하면 set_deadline에 업무 id와 새 마감을, 마감을 없애라고 하면 deadline을 빈 글로 넣습니다.\n"
    "- 끝냈다거나 완료 표시를 하라고 하면 done에 업무 id를 넣습니다. 이미 완료된 업무에는 넣지 않습니다.\n"
    "- 중요 표시를 하거나 빼라고 하면 important에 업무 id와 on(참: 표시, 거짓: 빼기)을 넣습니다.\n"
    "- 업무를 지우는 일은 하지 않습니다. 지워 달라고 하면 업무 줄의 ⋮ 메뉴에서 지울 수 있다고 안내합니다.\n"
    "- answer에는 무엇을 바꿀지 한 줄로 확인하듯 적습니다. 실제 반영은 사용자가 확인한 뒤에 합니다.\n"
    "- 어느 업무인지 애매하면 add·set_deadline·done·important를 모두 비우고 answer에서 후보를 보여 주며 되묻습니다.\n"
    "- 업무 id는 set_deadline·done·important에만 쓰고, answer에는 id를 절대 쓰지 마세요. 업무는 이름으로 말합니다.\n"
    "반드시 JSON만 출력합니다."
)

_FORMAT = ('{"answer": "답", "add": [{"name": "업무 이름", "deadline": "2026-11-21 15:00", "important": false}], '
           '"set_deadline": [{"task_id": 12, "deadline": "2026-11-18"}], "done": [{"task_id": 12}], '
           '"important": [{"task_id": 12, "on": true}]}')

# 자주 묻는 질문 (버튼 글자, 실제 질문)
QUICK = (
    ('오늘 할 일', '오늘 마감이거나 기한이 지난, 아직 끝내지 않은 업무를 알려줘. 없으면 없다고 해줘.'),
    ('이번 주 마감', '오늘부터 7일 안에 마감인, 아직 끝내지 않은 업무를 날짜순으로 알려줘.'),
    ('급한 순서', '아직 끝내지 않은 업무를 급한 순서(기한 지남 → 마감 가까운 순, 중요 표시 먼저)로 5개 알려줘.'),
)

_MASK_MARK = re.compile(r'\*{3,}|○')       # 개인정보를 가린 표시 (privacy.mask_pii)

_RECUR = {'weekly': '매주', 'biweekly': '격주', 'monthly': '매월', 'yearly': '매년'}


def fmt_deadline(deadline: str) -> str:
    """'2026-11-21 15:00' → '11월 21일(금) 15:00'. 빈 글이면 '마감 없음'."""
    if not deadline:
        return '마감 없음'
    try:
        d = date.fromisoformat(deadline[:10])
    except ValueError:
        return deadline
    text = f'{d.month}월 {d.day}일({_WEEKDAYS[d.weekday()]})'
    return text + (f' {deadline[11:16]}' if len(deadline) > 10 else '')


def window_labels(windows) -> dict:
    """업무창 id → '업무창 1' 같은 이름 (만든 순서대로)."""
    return {w['id']: f'업무창 {i}' for i, w in enumerate(sorted(windows, key=lambda w: w['id']), 1)}


def build_context(tasks, labels=None, current_window_id=None) -> str:
    """업무 목록 → AI에 함께 보낼 글. 업무마다 id를 붙여 바꾸기 제안 때 가리킬 수 있게."""
    labels = labels or {}
    if not tasks:
        return '(업무가 없습니다)'
    lines = []
    for t in tasks:
        parts = [f"id={t['id']}"]
        if labels:
            here = ' (지금 창)' if t.get('window_id') == current_window_id else ''
            parts.append(labels.get(t.get('window_id'), '업무창') + here)
        parts.append(str(t.get('name') or '').replace('\n', ' '))
        parts.append('마감 ' + (t.get('deadline') or '없음'))
        if t.get('deadline'):
            parts[-1] += f" ({fmt_deadline(t['deadline'])})"
        if t.get('priority'):
            parts.append('중요')
        rec = t.get('recurrence') or ''
        if rec:
            parts.append('반복 ' + _RECUR.get(rec, '사용자 지정'))
        if t.get('strikethrough'):
            parts.append('완료')
        lines.append(' | '.join(parts))
    return '\n'.join(lines)


def build_prompt(question: str, context: str, history=(), now: datetime = None) -> str:
    now = now or datetime.now()
    today = now.date()
    past = ''.join(f"\n이전 질문: {q}\n이전 답: {a}" for q, a in history[-2:])
    return (f"오늘: {today.isoformat()} ({_WEEKDAYS[today.weekday()]}요일), 지금 시각 {now:%H:%M}\n\n"
            f"업무 목록\n{context}\n{past}\n\n질문: {question}\n\n다음 형식의 JSON만 출력하세요:\n{_FORMAT}")


def _extract_json(raw: str):
    raw = (raw or '').strip()
    m = re.search(r'```(?:json)?\s*(.*?)```', raw, re.S)
    if m:
        raw = m.group(1).strip()
    try:
        return json.loads(raw)
    except Exception:
        a, b = raw.find('{'), raw.rfind('}')
        if a != -1 and b > a:
            try:
                return json.loads(raw[a:b + 1])
            except Exception:
                pass
    return None


def _to_int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    digits = re.sub(r'[^\d]', '', str(v or ''))
    return int(digits) if digits else None


def _to_bool(v, default=False):
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    return str(v).strip().lower() in ('true', '1', 'yes', 'on', '참', '예')


def norm_deadline(value, today: date = None):
    """AI가 준 마감 → ('YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM' 또는 '', 이유). 이유가 있으면 쓸 수 없는 값."""
    text = str(value or '').strip().replace('T', ' ')
    if not text:
        return '', None
    m = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::\d{2})?)?', text)
    if not m:
        return None, '날짜 형식을 읽지 못했어요'
    y, mo, d, hh, mm = m.groups()
    try:
        day = date(int(y), int(mo), int(d))
    except ValueError:
        return None, '없는 날짜예요'
    today = today or date.today()
    if abs((day - today).days) > 366 * 2:
        return None, '날짜가 너무 멀어요'
    if hh is None:
        return day.isoformat(), None
    if not (0 <= int(hh) <= 23 and 0 <= int(mm) <= 59):
        return None, '없는 시각이에요'
    return f'{day.isoformat()} {int(hh):02d}:{int(mm):02d}', None


def parse_reply(raw: str, tasks_by_id: dict, today: date = None) -> dict:
    """AI 응답 → {'answer': 글, 'changes': [...], 'rejected': [...]}.

    changes: 앱이 확인을 통과한 변경 {'kind': add|deadline|done|important, 'task': 업무(dict) 또는 None, ...}.
    rejected: 쓸 수 없는 제안 {'kind', 'task', 'label', 'reason'}. JSON이 아니면 글 전체를 답으로 쓴다.
    """
    today = today or date.today()
    data = _extract_json(raw)
    if not isinstance(data, dict):
        return {'answer': (raw or '').strip(), 'changes': [], 'rejected': []}
    answer = str(data.get('answer') or '').strip()
    changes, rejected = [], []

    seen_names = set()
    for s in (data.get('add') or [])[:ADD_MAX]:
        if not isinstance(s, dict):
            continue
        name = ' '.join(str(s.get('name') or '').split())[:NAME_MAX]
        if not name:
            rejected.append({'kind': 'add', 'task': None, 'label': '(이름 없음)', 'reason': '업무 이름이 비어 있어요'})
            continue
        deadline, reason = norm_deadline(s.get('deadline'), today)
        if reason:
            rejected.append({'kind': 'add', 'task': None, 'label': name, 'reason': reason})
            continue
        if _MASK_MARK.search(name):
            rejected.append({'kind': 'add', 'task': None, 'label': name,
                             'reason': '가린 개인정보가 들어 있어요 — 직접 입력해 주세요'})
            continue
        if name in seen_names:
            continue
        seen_names.add(name)
        changes.append({'kind': 'add', 'task': None, 'name': name, 'deadline': deadline,
                        'important': _to_bool(s.get('important'))})

    touched = set()                     # (업무 id, 종류) — 같은 업무에 같은 종류는 한 번만

    def task_for(s, kind):
        tid = _to_int(s.get('task_id'))
        task = tasks_by_id.get(tid)
        if task is None:
            rejected.append({'kind': kind, 'task': None, 'label': f'업무 {tid}', 'reason': '없는 업무예요'})
            return None
        if (tid, kind) in touched:
            return None
        return task

    for s in data.get('set_deadline') or []:
        if not isinstance(s, dict):
            continue
        task = task_for(s, 'deadline')
        if task is None:
            continue
        deadline, reason = norm_deadline(s.get('deadline'), today)
        if reason:
            rejected.append({'kind': 'deadline', 'task': task, 'label': task['name'], 'reason': reason})
            continue
        if deadline == (task.get('deadline') or ''):
            continue                                    # 변화 없음: 조용히 뺌
        touched.add((task['id'], 'deadline'))
        changes.append({'kind': 'deadline', 'task': task, 'old': task.get('deadline') or '', 'new': deadline})

    for s in data.get('done') or []:
        if not isinstance(s, dict):
            continue
        task = task_for(s, 'done')
        if task is None:
            continue
        if task.get('strikethrough'):
            rejected.append({'kind': 'done', 'task': task, 'label': task['name'], 'reason': '이미 완료된 업무예요'})
            continue
        touched.add((task['id'], 'done'))
        changes.append({'kind': 'done', 'task': task})

    for s in data.get('important') or []:
        if not isinstance(s, dict):
            continue
        task = task_for(s, 'important')
        if task is None:
            continue
        on = _to_bool(s.get('on'), True)
        if bool(task.get('priority')) == on:
            continue
        touched.add((task['id'], 'important'))
        changes.append({'kind': 'important', 'task': task, 'on': on})

    return {'answer': answer, 'changes': changes, 'rejected': rejected}


def describe(change: dict) -> str:
    """제안 한 줄 설명 (확인 화면에 보여 줄 글)."""
    k = change['kind']
    if k == 'add':
        star = ' · ★ 중요' if change.get('important') else ''
        return f"추가: {change['name']} · {fmt_deadline(change['deadline'])}{star}"
    name = change['task']['name']
    if k == 'deadline':
        return f"마감 바꾸기: {name} · {fmt_deadline(change['old'])} → {fmt_deadline(change['new'])}"
    if k == 'done':
        return f'완료 표시: {name}'
    return f"중요 표시 {'하기' if change['on'] else '빼기'}: {name}"


# ── 반영 · 되돌리기 (DB) ──────────────────────────────────────

def apply_changes(changes, window_id) -> dict:
    """제안을 DB에 넣는다. 되돌리기용 기록 {'added': [새 업무 id], 'before': [바꾸기 전 업무]}을 돌려준다.
    새 업무는 window_id 업무창에 넣는다."""
    import db
    undo = {'added': [], 'before': []}
    saved = set()
    for c in changes:
        if c['kind'] == 'add':
            tid = db.add_task(window_id, c['name'], c['deadline'], priority=1 if c.get('important') else 0)
            undo['added'].append(tid)
            continue
        task = _current_task(c['task']['id'])
        if task is None:
            continue
        if task['id'] not in saved:
            undo['before'].append(dict(task))
            saved.add(task['id'])
        deadline, strike, prio = task['deadline'], task.get('strikethrough', 0), task.get('priority', 0)
        if c['kind'] == 'deadline':
            deadline = c['new']
        elif c['kind'] == 'done':
            strike = 1
        elif c['kind'] == 'important':
            prio = 1 if c['on'] else 0
        db.update_task(task['id'], task['name'], deadline, strikethrough=strike, priority=prio,
                       recurrence=task.get('recurrence', ''))
    return undo


def undo_changes(undo: dict):
    """apply_changes 전으로 되돌린다 (새로 넣은 업무는 지우고, 바꾼 업무는 원래 값으로)."""
    import db
    for tid in undo.get('added', []):
        db.delete_task(tid)
    for t in undo.get('before', []):
        if _current_task(t['id']) is None:
            continue
        db.update_task(t['id'], t['name'], t.get('deadline', ''), strikethrough=t.get('strikethrough', 0),
                       priority=t.get('priority', 0), recurrence=t.get('recurrence', ''))


def _current_task(task_id):
    import db
    return db.get_task(task_id)


def task_windows_tasks():
    """업무 모드 창(메모장 아님)의 모든 업무와 창 이름표 → (업무 목록, {창 id: '업무창 n'})."""
    import db
    windows = [w for w in db.get_all_windows() if not w.get('memo_mode')]
    tasks = []
    for w in sorted(windows, key=lambda w: w['id']):
        tasks.extend(db.get_tasks(w['id']))
    return tasks, window_labels(windows)


# ── 실행 (뒤에서) ────────────────────────────────────────────────

try:
    from PyQt5.QtCore import QThread, pyqtSignal
except ImportError:          # 테스트 환경 등
    QThread = object

    def pyqtSignal(*_a):     # noqa: N802
        return None


class TaskAskWorker(QThread):
    """질문 → AI 답.  done(reply: dict, masked: dict) / failed(message: str)"""
    done = pyqtSignal(dict, dict)
    failed = pyqtSignal(str)

    def __init__(self, question, tasks, labels=None, current_window_id=None, history=(), parent=None):
        super().__init__(parent)
        self._q = question
        self._tasks = [dict(t) for t in tasks]
        self._labels = dict(labels or {})
        self._wid = current_window_id
        self._history = list(history)

    def run(self):
        try:
            from ai_client import get_adapter
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있어요. 위의 AI 연결 버튼에서 OpenCode·Claude·ChatGPT·Claude 구독 중 하나를 연결하세요.')
                return
            if not adapter.get_api_key():
                self.failed.emit('AI API 키가 없어요. 위의 AI 연결 버튼에서 키를 넣어 주세요.')
                return
            context = build_context(self._tasks, self._labels, self._wid)
            prompt = build_prompt(self._q, context, self._history)
            masked = {}
            try:
                from ai_client import masking_required
                from school_profile import mask_for_ai
                if masking_required():
                    prompt, masked = mask_for_ai(prompt)
            except Exception:
                pass
            raw = adapter.complete(SYSTEM, prompt, max_tokens=1500)
            reply = parse_reply(raw, {t['id']: t for t in self._tasks})
            if not reply['answer'] and not reply['changes'] and not reply['rejected']:
                self.failed.emit('AI가 빈 답을 보냈어요. 다시 물어봐 주세요.')
                return
            self.done.emit(reply, masked or {})
        except Exception as e:
            try:
                from ai_client import friendly_error
                msg = friendly_error(e) if getattr(e, 'status_code', None) else str(e)
            except Exception:
                msg = str(e)
            self.failed.emit(msg)
