"""ai_actions.py — AI가 앱을 다루는 기능 목록 (업무 · 업무창·메모장 · 예산 · 설정 · AI 대화 창).

AI에게 기능 목록과 지금 앱 상태를 보내면, AI는 {"answer": 답, "actions": [{"do": 기능, ...값}]}만 돌려준다.
앱이 기능마다 값을 검사(check)하고, 위험 정도에 따라 처리한다.
  · view   (보기·열기: 창 열기, 대화 지우기)            → 바로 실행
  · change (추가·수정·꾸미기·설정)                      → 확인 칸 [넣기] → 실행, 되돌리기 1회
  · delete (업무·예산 항목·세부항목 지우기)              → 빨간 확인 칸 [지우기] → 실행, 되돌리기 1회
AI가 하지 않는 일(안내만): 예산 전체 지우기, 업무창·메모장 자체 삭제, AI 연결·API 키·개인정보 보호 설정, 업데이트, 앱 종료.
한 번에 MAX_ACTIONS개까지.

실행(run)·되돌리기(undo)는 화면 쪽 어댑터(app)를 받는다 — ui/ai_chat_window.ChatApp.
검사(check)는 화면 없이 ctx(build_state가 만든 지금 상태)만 보므로 뒤 스레드에서도, 테스트에서도 돈다.
"""
import json
from datetime import date, datetime

from task_ai import (_MASK_MARK, _WEEKDAYS, NAME_MAX, _extract_json, _to_bool, _to_int, fmt_deadline,
                     norm_deadline)

MAX_ACTIONS = 20
VIEW, CHANGE, DELETE = 'view', 'change', 'delete'
TEXT_MAX = 2000                      # 메모·업무 메모 글 한도

from ui.theme import PALETTE_NAMES as COLOR_NAMES          # hex → 이름
COLOR_ALIASES = {'노란색': '노랑', '진노랑': '노랑', '주황': '노랑', '귤': '노랑', '살구': '분홍', '복숭아': '분홍',
                 '보라': '라벤더', '파랑': '하늘', '초록': '연두', '민트': '연두', '베이지': '연회색', '갈색': '연회색',
                 '모래': '연회색', '회색': '연회색', '하얀색': '흰색'}
FONT_KEYS = {'pretendard': 'Pretendard', 'reatic': '리틱체', 'reatic_narrow': '리틱체 Narrow', 'malgun': '맑은 고딕'}
SIZE_VALUES = {'기본': 1.0, '크게': 1.1, '더 크게': 1.2}
GAP_VALUES = (0, 5, 10, 15)
REPEATS = {'none': '반복 없음', 'weekly': '매주', 'biweekly': '격주', 'monthly': '매월', 'yearly': '매년'}
OPEN_TARGETS = {'calendar': '달력', 'dashboard': '오늘 한눈에 보기', 'history': '지난 기록(지난 업무)',
                'documents': '지난 기록(작성 공문)', 'document': '공문 작성', 'budget': '예산 관리'}
ALARM_MINUTES = (0, 30, 60, 180)
CHAT_SIDES = {'inside': '업무창 안', 'left': '왼쪽', 'right': '오른쪽', 'bottom': '아래쪽'}


class Reject(Exception):
    """쓸 수 없는 제안 (이유를 사용자에게 보여 줌)."""


class Skip(Exception):
    """이미 그 상태라 바꿀 것이 없음 (조용히 뺌)."""


_ACTIONS = {}


def action(name, risk, help_text):
    """기능 등록. cls에 check(p, ctx) → (값, 설명), run(값, app) → 되돌리기 기록, undo(기록, app)."""
    def deco(cls):
        _ACTIONS[name] = {'name': name, 'risk': risk, 'help': help_text, 'check': cls.check,
                          'run': getattr(cls, 'run', None), 'undo': getattr(cls, 'undo', None)}
        return cls
    return deco


def names():
    return list(_ACTIONS)


# ── 값 읽기 도우미 ───────────────────────────────────────────────
def _text(p, key, required=True, limit=NAME_MAX, label='글'):
    v = p.get(key)
    t = ' '.join(str(v).split()) if limit <= NAME_MAX else str(v if v is not None else '').strip()
    t = t[:limit]
    if required and not t:
        raise Reject(f'{label}이(가) 비어 있어요')
    if t and _MASK_MARK.search(t):
        raise Reject('가린 개인정보가 들어 있어요 — 직접 입력해 주세요')
    return t


def _task(p, ctx):
    tid = _to_int(p.get('task_id'))
    t = ctx['tasks_by_id'].get(tid)
    if t is None:
        raise Reject('없는 업무예요')
    return t


def _window(p, ctx, kind=None, key='window_id', default_current=False):
    wid = _to_int(p.get(key))
    if wid is None and default_current:
        wid = ctx.get('current_window')
        if kind == 'task' and ctx['windows_by_id'].get(wid, {}).get('kind') != 'task':
            wid = next((w['id'] for w in ctx['windows'] if w['kind'] == 'task'), None)
    w = ctx['windows_by_id'].get(wid)
    if w is None:
        raise Reject('없는 창이에요')
    if kind and w['kind'] != kind:
        raise Reject('업무창이 아니에요' if kind == 'task' else '메모장이 아니에요')
    return w


def _color(v):
    s = str(v or '').strip()
    if not s:
        return None
    up = s.upper()
    if up in COLOR_NAMES:
        return up
    from ui.theme import OLD_COLOR_MAP
    if up in OLD_COLOR_MAP:
        return OLD_COLOR_MAP[up]
    name = s.replace(' ', '')
    name = COLOR_ALIASES.get(name, name)
    for hx, nm in COLOR_NAMES.items():
        if name == nm:
            return hx
    raise Reject(f'고를 수 없는 색이에요: {s}')


def _font(v):
    s = str(v or '').strip()
    if not s:
        return None
    if s in FONT_KEYS:
        return s
    for k, nm in FONT_KEYS.items():
        if s.replace(' ', '').lower() == nm.replace(' ', '').lower():
            return k
    raise Reject(f'고를 수 없는 글꼴이에요: {s}')


def _size(v):
    if v in (None, ''):
        return None
    if str(v).strip() in SIZE_VALUES:
        return SIZE_VALUES[str(v).strip()]
    try:
        f = float(v)
    except (TypeError, ValueError):
        raise Reject('글자 크기는 기본·크게·더 크게 중 하나예요')
    if f not in SIZE_VALUES.values():
        raise Reject('글자 크기는 기본·크게·더 크게 중 하나예요')
    return f


def _gap(v):
    if v in (None, ''):
        return None
    g = _to_int(v)
    if g not in GAP_VALUES:
        raise Reject('줄 간격은 0·5·10·15 중 하나예요')
    return g


def _money(p, key, required=True):
    v = _to_int(p.get(key))
    if v is None:
        if required:
            raise Reject('금액을 읽지 못했어요')
        return None
    if v < 0:
        raise Reject('금액이 0원보다 작아요')
    return v


def _wname(ctx, wid):
    return ctx['windows_by_id'].get(wid, {}).get('label', '창')


# ════════════════════════════════════════════════════════════════
# 업무
# ════════════════════════════════════════════════════════════════
def _task_row(app, tid):
    import db
    return db.get_task(tid)


def _save_task(t):
    import db
    db.update_task(t['id'], t['name'], t.get('deadline', ''), strikethrough=t.get('strikethrough', 0),
                   priority=t.get('priority', 0), recurrence=t.get('recurrence', ''))


def _restore_fields(rec, app):
    """업무 한 건을 되돌리기 전 값으로 (지워졌으면 그냥 넘어감)."""
    import db
    cur = db.get_task(rec['id'])
    if cur is None:
        return
    _save_task(rec)
    if 'notes' in rec:
        db.set_task_notes(rec['id'], rec['notes'])
    if 'related_no' in rec:
        db.set_task_related_no(rec['id'], rec['related_no'])
    if cur['window_id'] != rec['window_id']:
        db.set_task_window(rec['id'], rec['window_id'])


class _TaskChange:
    """업무 한 건 바꾸기 공통: run 전 값을 보관 → undo로 되돌림."""
    field = None

    @classmethod
    def run(cls, v, app):
        before = _task_row(app, v['task_id'])
        if before is None:
            return None
        cls.apply(before, v)
        app.refresh_tasks()
        return dict(before)

    @staticmethod
    def undo(rec, app):
        if rec:
            _restore_fields(rec, app)
            app.refresh_tasks()


@action('task.add', CHANGE, 'task.add {name, deadline?: "YYYY-MM-DD" 또는 "YYYY-MM-DD HH:MM", important?: true/false, '
        'window_id?: 넣을 업무창(없으면 지금 창)} — 업무 추가')
class TaskAdd:
    @staticmethod
    def check(p, ctx):
        name = _text(p, 'name', label='업무 이름')
        deadline, reason = norm_deadline(p.get('deadline'), ctx['today'])
        if reason:
            raise Reject(reason)
        w = _window(p, ctx, kind='task', default_current=True)
        imp = _to_bool(p.get('important'))
        v = {'name': name, 'deadline': deadline or '', 'important': imp, 'window_id': w['id']}
        star = ' · ★ 중요' if imp else ''
        where = f' · {w["label"]}' if len([x for x in ctx['windows'] if x['kind'] == 'task']) > 1 else ''
        return v, f'업무 추가: {name} · {fmt_deadline(v["deadline"])}{star}{where}'

    @staticmethod
    def run(v, app):
        import db
        tid = db.add_task(v['window_id'], v['name'], v['deadline'], priority=1 if v['important'] else 0)
        app.refresh_tasks()
        return tid

    @staticmethod
    def undo(tid, app):
        import db
        db.delete_task(tid)
        app.refresh_tasks()


@action('task.rename', CHANGE, 'task.rename {task_id, name} — 업무 이름 바꾸기')
class TaskRename(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        name = _text(p, 'name', label='업무 이름')
        if name == t['name']:
            raise Skip()
        return {'task_id': t['id'], 'name': name}, f'이름 바꾸기: {t["name"]} → {name}'

    @staticmethod
    def apply(t, v):
        _save_task({**t, 'name': v['name']})


@action('task.deadline', CHANGE, 'task.deadline {task_id, deadline: 날짜 또는 "" (마감 없애기)} — 마감 바꾸기')
class TaskDeadline(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        deadline, reason = norm_deadline(p.get('deadline'), ctx['today'])
        if reason:
            raise Reject(reason)
        if (deadline or '') == (t.get('deadline') or ''):
            raise Skip()
        return ({'task_id': t['id'], 'deadline': deadline or ''},
                f'마감 바꾸기: {t["name"]} · {fmt_deadline(t.get("deadline") or "")} → {fmt_deadline(deadline or "")}')

    @staticmethod
    def apply(t, v):
        _save_task({**t, 'deadline': v['deadline']})


@action('task.done', CHANGE, 'task.done {task_id, on?: true(완료 표시)/false(완료 해제)} — 완료 표시')
class TaskDone(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        on = _to_bool(p.get('on'), True)
        if bool(t.get('strikethrough')) == on:
            raise Skip()
        return {'task_id': t['id'], 'on': on}, f'완료 표시{"" if on else " 해제"}: {t["name"]}'

    @staticmethod
    def apply(t, v):
        _save_task({**t, 'strikethrough': 1 if v['on'] else 0})


@action('task.important', CHANGE, 'task.important {task_id, on: true/false} — 중요 표시(★)')
class TaskImportant(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        on = _to_bool(p.get('on'), True)
        if bool(t.get('priority')) == on:
            raise Skip()
        return {'task_id': t['id'], 'on': on}, f'중요 표시 {"하기" if on else "빼기"}: {t["name"]}'

    @staticmethod
    def apply(t, v):
        _save_task({**t, 'priority': 1 if v['on'] else 0})


@action('task.repeat', CHANGE, 'task.repeat {task_id, repeat: none|weekly|biweekly|monthly|yearly} — 반복 (마감이 있어야 함)')
class TaskRepeat(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        r = str(p.get('repeat') or 'none').strip().lower()
        if r not in REPEATS:
            raise Reject('반복은 없음·매주·격주·매월·매년 중 하나예요')
        r = '' if r == 'none' else r
        if r and not t.get('deadline'):
            raise Reject('마감이 있어야 반복할 수 있어요')
        if r == (t.get('recurrence') or ''):
            raise Skip()
        return {'task_id': t['id'], 'repeat': r}, f'반복: {t["name"]} → {REPEATS[r or "none"]}'

    @staticmethod
    def apply(t, v):
        _save_task({**t, 'recurrence': v['repeat']})


@action('task.note', CHANGE, 'task.note {task_id, text, mode?: replace|append} — 업무 메모 쓰기(바꾸기/덧붙이기)')
class TaskNote(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        text = _text(p, 'text', required=False, limit=TEXT_MAX, label='메모')
        mode = 'append' if str(p.get('mode') or '').lower() == 'append' else 'replace'
        if mode == 'append' and not text:
            raise Skip()
        short = text if len(text) <= 30 else text[:30] + '…'
        what = '덧붙이기' if mode == 'append' else ('지우기' if not text else '쓰기')
        return {'task_id': t['id'], 'text': text, 'mode': mode}, f'업무 메모 {what}: {t["name"]} · {short}'

    @staticmethod
    def apply(t, v):
        import db
        old = t.get('notes') or ''
        db.set_task_notes(t['id'], (old + ('\n' if old else '') + v['text']) if v['mode'] == 'append' else v['text'])


@action('task.related', CHANGE, 'task.related {task_id, text} — 업무의 관련 공문번호')
class TaskRelated(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        text = _text(p, 'text', required=False, limit=80, label='관련 번호')
        if text == (t.get('related_no') or ''):
            raise Skip()
        return {'task_id': t['id'], 'text': text}, f'관련 번호: {t["name"]} → {text or "(지움)"}'

    @staticmethod
    def apply(t, v):
        import db
        db.set_task_related_no(t['id'], v['text'])


@action('task.move', CHANGE, 'task.move {task_id, window_id: 옮길 업무창} — 다른 업무창으로 옮기기')
class TaskMove(_TaskChange):
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        w = _window(p, ctx, kind='task')
        if w['id'] == t.get('window_id'):
            raise Skip()
        return {'task_id': t['id'], 'window_id': w['id']}, f'옮기기: {t["name"]} → {w["label"]}'

    @staticmethod
    def apply(t, v):
        import db
        db.set_task_window(t['id'], v['window_id'])


@action('task.delete', DELETE, 'task.delete {task_id} — 업무 지우기 (지난 기록에 남음)')
class TaskDelete:
    @staticmethod
    def check(p, ctx):
        t = _task(p, ctx)
        return {'task_id': t['id']}, f'지우기: {t["name"]}'

    @staticmethod
    def run(v, app):
        import db
        t = db.get_task(v['task_id'])
        if t is None:
            return None
        hid = db.add_task_history(t['window_id'], t['name'], t.get('deadline', ''),
                                  strikethrough=t.get('strikethrough', 0), priority=t.get('priority', 0),
                                  recurrence=t.get('recurrence', ''), notes=t.get('notes', ''),
                                  related_no=t.get('related_no', ''))
        db.delete_task(t['id'])
        app.refresh_tasks()
        return {'task': dict(t), 'history_id': hid}

    @staticmethod
    def undo(rec, app):
        import db
        if not rec:
            return
        db.restore_task(rec['task'])
        if rec.get('history_id'):
            db.delete_task_history(rec['history_id'])
        app.refresh_tasks()


# ════════════════════════════════════════════════════════════════
# 업무창 · 메모장
# ════════════════════════════════════════════════════════════════
@action('window.new', CHANGE, 'window.new {kind: task(업무창)|memo(새 메모장), text?: 메모장에 쓸 글} — 새 창 만들기')
class WindowNew:
    @staticmethod
    def check(p, ctx):
        kind = 'memo' if str(p.get('kind') or '').lower() in ('memo', '메모장', 'note') else 'task'
        text = _text(p, 'text', required=False, limit=TEXT_MAX, label='메모') if kind == 'memo' else ''
        label = '새 메모장' if kind == 'memo' else '새 업무창'
        return {'kind': kind, 'text': text}, label + (f' · {text[:30]}' if text else '')

    @staticmethod
    def run(v, app):
        w = app.new_window(v['kind'])
        if w is None:
            return None
        if v['text']:
            app.set_memo_text(w, v['text'])
        return w.window_id

    @staticmethod
    def undo(wid, app):
        if wid:
            app.delete_window(wid)


@action('memo.write', CHANGE, 'memo.write {window_id: 메모장, text, mode?: replace|append} — 메모장 글 쓰기(바꾸기/덧붙이기)')
class MemoWrite:
    @staticmethod
    def check(p, ctx):
        w = _window(p, ctx, kind='memo')
        text = _text(p, 'text', required=False, limit=TEXT_MAX, label='메모')
        mode = 'append' if str(p.get('mode') or '').lower() == 'append' else 'replace'
        if mode == 'append' and not text:
            raise Skip()
        return ({'window_id': w['id'], 'text': text, 'mode': mode},
                f'{w["label"]}에 글 {"덧붙이기" if mode == "append" else "쓰기"}: {text[:30]}{"…" if len(text) > 30 else ""}')

    @staticmethod
    def run(v, app):
        w = app.window(v['window_id'])
        if w is None:
            return None
        old = w.memo_editor.toPlainText()
        new = (old + ('\n' if old and not old.endswith('\n') else '') + v['text']) if v['mode'] == 'append' else v['text']
        app.set_memo_text(w, new)
        return {'window_id': v['window_id'], 'text': old}

    @staticmethod
    def undo(rec, app):
        w = app.window(rec['window_id']) if rec else None
        if w is not None:
            app.set_memo_text(w, rec['text'])


@action('window.style', CHANGE, 'window.style {window_id, color?: 색 이름, font?: pretendard|reatic|reatic_narrow|malgun, '
        'size?: 기본|크게|더 크게, gap?: 0|5|10|15} — 업무창·메모장 배경색·글자·줄간격')
class WindowStyle:
    @staticmethod
    def check(p, ctx):
        w = _window(p, ctx, default_current=True)
        v = {'window_id': w['id'], 'color': _color(p.get('color')), 'font': _font(p.get('font')),
             'size': _size(p.get('size')), 'gap': _gap(p.get('gap'))}
        if v['color'] and v['color'] == (w.get('color') or '').upper():
            v['color'] = None
        if v['font'] == w.get('font'):
            v['font'] = None
        if v['size'] is not None and abs(v['size'] - (w.get('scale') or 1.0)) < 0.01:
            v['size'] = None
        if v['gap'] is not None and v['gap'] == w.get('gap'):
            v['gap'] = None
        parts = []
        if v['color']:
            parts.append(f'배경색 {COLOR_NAMES[v["color"]]}')
        if v['font']:
            parts.append(f'글꼴 {FONT_KEYS[v["font"]]}')
        if v['size'] is not None:
            parts.append('글자 ' + next(k for k, x in SIZE_VALUES.items() if x == v['size']))
        if v['gap'] is not None:
            parts.append(f'줄 간격 {v["gap"]}')
        if not parts:
            raise Skip()
        return v, f'{w["label"]} 꾸미기: ' + ' · '.join(parts)

    @staticmethod
    def run(v, app):
        w = app.window(v['window_id'])
        if w is None:
            return None
        rec = {'window_id': v['window_id'], 'color': w._bg_color, 'font': w._font_key, 'size': w._scale,
               'gap': w._row_gap}
        app.style_window(w, v['color'], v['font'], v['size'], v['gap'])
        return rec

    @staticmethod
    def undo(rec, app):
        w = app.window(rec['window_id']) if rec else None
        if w is not None:
            app.style_window(w, rec['color'], rec['font'], rec['size'], rec['gap'])


@action('window.fold', CHANGE, 'window.fold {window_id, on: true(제목만 남게 접기)/false(펼치기)} — 접기·펼치기')
class WindowFold:
    @staticmethod
    def check(p, ctx):
        w = _window(p, ctx, default_current=True)
        on = _to_bool(p.get('on'), True)
        if bool(w.get('collapsed')) == on:
            raise Skip()
        return {'window_id': w['id'], 'on': on}, f'{w["label"]} {"접기" if on else "펼치기"}'

    @staticmethod
    def run(v, app):
        w = app.window(v['window_id'])
        if w is not None and bool(w.collapsed) != v['on']:
            w.toggle_shade()
            return v['window_id']
        return None

    @staticmethod
    def undo(wid, app):
        w = app.window(wid) if wid else None
        if w is not None:
            w.toggle_shade()


@action('window.pin', CHANGE, 'window.pin {window_id, on: true/false} — 항상 위에 고정')
class WindowPin:
    @staticmethod
    def check(p, ctx):
        w = _window(p, ctx, default_current=True)
        on = _to_bool(p.get('on'), True)
        if bool(w.get('pin')) == on:
            raise Skip()
        return {'window_id': w['id'], 'on': on}, f'{w["label"]} 항상 위 {"켜기" if on else "끄기"}'

    @staticmethod
    def run(v, app):
        w = app.window(v['window_id'])
        if w is not None and bool(w.pin_active) != v['on']:
            w.toggle_always_on_top()
            return v['window_id']
        return None

    @staticmethod
    def undo(wid, app):
        w = app.window(wid) if wid else None
        if w is not None:
            w.toggle_always_on_top()


@action('open', VIEW, 'open {what: calendar|dashboard|history|documents|document|budget, title?: 공문 작성 제목, '
        'category_id?: 예산 창에서 보여 줄 세부항목} — 창 열기')
class Open:
    @staticmethod
    def check(p, ctx):
        what = str(p.get('what') or '').strip().lower()
        if what not in OPEN_TARGETS:
            raise Reject('열 수 없는 창이에요')
        v = {'what': what, 'title': '', 'category_id': None}
        if what == 'document':
            v['title'] = _text(p, 'title', required=False, label='제목')
        if what == 'budget' and p.get('category_id') not in (None, ''):
            cid = _to_int(p.get('category_id'))
            if cid not in ctx['cats_by_id']:
                raise Reject('없는 세부항목이에요')
            v['category_id'] = cid
        extra = f' · {v["title"]}' if v['title'] else (
            f' · {ctx["cats_by_id"][v["category_id"]]["name"]}' if v['category_id'] else '')
        return v, f'{OPEN_TARGETS[what]} 열기{extra}'

    @staticmethod
    def run(v, app):
        app.open(v['what'], title=v['title'], category_id=v['category_id'])


# ════════════════════════════════════════════════════════════════
# 예산
# ════════════════════════════════════════════════════════════════
def _item(p, ctx):
    iid = _to_int(p.get('item_id'))
    it = ctx['items_by_id'].get(iid)
    if it is None:
        raise Reject('없는 예산 항목이에요')
    return it


def _cat(p, ctx, key='category_id'):
    cid = _to_int(p.get(key))
    c = ctx['cats_by_id'].get(cid)
    if c is None:
        raise Reject('없는 세부항목이에요')
    return c


def _won(n):
    return f'{int(n):,}원'


def _item_label(it):
    return it.get('description') or it.get('cost_type') or '항목'


class _SpentBase:
    kind = None

    @classmethod
    def check(cls, p, ctx):
        from budget_ai import _check
        it = _item(p, ctx)
        if it.get('completed'):
            raise Reject('완료된 항목이에요')
        key = {'spend': 'amount', 'set_remaining': 'remaining', 'set_spent': 'spent'}[cls.kind]
        value = _to_int(p.get(key))
        new, reason = _check(cls.kind, value, it)
        if reason:
            raise Reject(reason)
        if new is None:
            raise Skip()
        purpose = _text(p, 'purpose', required=False, limit=40) if cls.kind == 'spend' else ''
        b = int(it['budget'])
        v = {'item_id': it['id'], 'kind': cls.kind, 'value': value, 'purpose': purpose, 'new_spent': new}
        head = {'spend': f'{_won(value)} 썼어요', 'set_remaining': f'잔액을 {_won(value)}로',
                'set_spent': f'집행액을 {_won(value)}로'}[cls.kind]
        pur = f' ({purpose})' if purpose else ''
        return v, (f'예산: {_item_label(it)} · {head}{pur} → 집행 {_won(new)} · 잔액 {_won(b - new)}')

    @staticmethod
    def run(v, app):
        import db
        cur = next((i for i in db.get_budget_items() if i['id'] == v['item_id']), None)
        if cur is None or cur.get('completed'):
            return None
        from budget_ai import target_spent
        new = target_spent(v['kind'], v['value'], cur)
        rec = {'item_id': cur['id'], 'spent': cur['spent'], 'expense_id': None}
        if v['kind'] == 'spend':
            rec['expense_id'] = db.add_budget_expense(cur['id'], v['value'], v.get('purpose', ''),
                                                      date.today().isoformat())
        else:
            db.set_budget_item_spent(cur['id'], new)
        app.refresh_budget()
        return rec

    @staticmethod
    def undo(rec, app):
        import db
        if not rec:
            return
        if rec.get('expense_id'):
            db.delete_budget_expense(rec['expense_id'], adjust_spent=False)
        db.set_budget_item_spent(rec['item_id'], rec['spent'])
        app.refresh_budget()


@action('budget.spend', CHANGE, 'budget.spend {item_id, amount, purpose?} — 돈을 썼다(집행액에 더하고 사용 내역에 남김)')
class BudgetSpend(_SpentBase):
    kind = 'spend'


@action('budget.set_remaining', CHANGE, 'budget.set_remaining {item_id, remaining} — 잔액을 이 금액으로 맞추기')
class BudgetSetRemaining(_SpentBase):
    kind = 'set_remaining'


@action('budget.set_spent', CHANGE, 'budget.set_spent {item_id, spent} — 집행액을 이 금액으로 고치기')
class BudgetSetSpent(_SpentBase):
    kind = 'set_spent'


@action('budget.item_add', CHANGE, 'budget.item_add {category_id, cost_type, description, budget, spent?} — 예산 항목 추가')
class BudgetItemAdd:
    @staticmethod
    def check(p, ctx):
        c = _cat(p, ctx)
        desc = _text(p, 'description', label='산출내역')
        cost = _text(p, 'cost_type', required=False, limit=40)
        budget = _money(p, 'budget')
        spent = _money(p, 'spent', required=False) or 0
        v = {'category_id': c['id'], 'cost_type': cost, 'description': desc, 'budget': budget, 'spent': spent}
        return v, f'예산 항목 추가: {c["name"]} · {cost + " · " if cost else ""}{desc} · 예산 {_won(budget)}'

    @staticmethod
    def run(v, app):
        import db
        iid = db.add_budget_item(v['category_id'], v['cost_type'], v['description'], v['budget'], v['spent'])
        app.refresh_budget()
        return iid

    @staticmethod
    def undo(iid, app):
        import db
        db.delete_budget_item(iid)
        app.refresh_budget()


@action('budget.item_edit', CHANGE, 'budget.item_edit {item_id, cost_type?, description?, budget?} — 예산 항목 고치기')
class BudgetItemEdit:
    @staticmethod
    def check(p, ctx):
        it = _item(p, ctx)
        cost = _text(p, 'cost_type', required=False, limit=40) if 'cost_type' in p else it.get('cost_type', '')
        desc = _text(p, 'description', required=False) if 'description' in p else it.get('description', '')
        budget = _money(p, 'budget', required=False)
        budget = it['budget'] if budget is None else budget
        if not desc:
            raise Reject('산출내역이 비어 있어요')
        if (cost, desc, int(budget)) == (it.get('cost_type', ''), it.get('description', ''), int(it['budget'])):
            raise Skip()
        v = {'item_id': it['id'], 'cost_type': cost, 'description': desc, 'budget': int(budget)}
        return v, f'예산 항목 고치기: {_item_label(it)} → {cost + " · " if cost else ""}{desc} · 예산 {_won(budget)}'

    @staticmethod
    def run(v, app):
        import db
        cur = next((i for i in db.get_budget_items() if i['id'] == v['item_id']), None)
        if cur is None:
            return None
        db.update_budget_item(cur['id'], cur['category_id'], v['cost_type'], v['description'], v['budget'], cur['spent'])
        app.refresh_budget()
        return dict(cur)

    @staticmethod
    def undo(rec, app):
        import db
        if rec:
            db.update_budget_item(rec['id'], rec['category_id'], rec['cost_type'], rec['description'],
                                  rec['budget'], rec['spent'])
            app.refresh_budget()


@action('budget.item_done', CHANGE, 'budget.item_done {item_id, on: true/false} — 예산 항목 완료 표시')
class BudgetItemDone:
    @staticmethod
    def check(p, ctx):
        it = _item(p, ctx)
        on = _to_bool(p.get('on'), True)
        if bool(it.get('completed')) == on:
            raise Skip()
        return {'item_id': it['id'], 'on': on}, f'예산 항목 완료{"" if on else " 취소"}: {_item_label(it)}'

    @staticmethod
    def run(v, app):
        import db
        db.set_budget_item_completed(v['item_id'], v['on'])
        app.refresh_budget()
        return v

    @staticmethod
    def undo(v, app):
        import db
        db.set_budget_item_completed(v['item_id'], not v['on'])
        app.refresh_budget()


@action('budget.item_delete', DELETE, 'budget.item_delete {item_id} — 예산 항목 지우기 (사용 내역도 함께)')
class BudgetItemDelete:
    @staticmethod
    def check(p, ctx):
        it = _item(p, ctx)
        return {'item_id': it['id']}, f'예산 항목 지우기: {_item_label(it)} (예산 {_won(it["budget"])} · 집행 {_won(it["spent"])})'

    @staticmethod
    def run(v, app):
        import db
        cur = next((i for i in db.get_budget_items() if i['id'] == v['item_id']), None)
        if cur is None:
            return None
        ex = [e for e in db.get_budget_expenses() if e['item_id'] == cur['id']]
        db.delete_budget_item(cur['id'])
        app.refresh_budget()
        return {'item': dict(cur), 'expenses': ex}

    @staticmethod
    def undo(rec, app):
        import db
        if rec:
            db.restore_budget_item(rec['item'], rec['expenses'])
            app.refresh_budget()


@action('budget.cat_add', CHANGE, 'budget.cat_add {name} — 세부항목 만들기')
class BudgetCatAdd:
    @staticmethod
    def check(p, ctx):
        name = _text(p, 'name', label='세부항목 이름')
        if any(c['name'] == name for c in ctx['budget']['cats']):
            raise Reject('같은 이름의 세부항목이 이미 있어요')
        return {'name': name}, f'세부항목 만들기: {name}'

    @staticmethod
    def run(v, app):
        import db
        cid = db.add_budget_category(v['name'])
        app.refresh_budget()
        return cid

    @staticmethod
    def undo(cid, app):
        import db
        if cid:
            db.delete_budget_category(cid)
            app.refresh_budget()


@action('budget.cat_rename', CHANGE, 'budget.cat_rename {category_id, name} — 세부항목 이름 바꾸기')
class BudgetCatRename:
    @staticmethod
    def check(p, ctx):
        c = _cat(p, ctx)
        name = _text(p, 'name', label='세부항목 이름')
        if name == c['name']:
            raise Skip()
        if any(x['name'] == name for x in ctx['budget']['cats']):
            raise Reject('같은 이름의 세부항목이 이미 있어요')
        return {'category_id': c['id'], 'name': name, 'old': c['name']}, f'세부항목 이름: {c["name"]} → {name}'

    @staticmethod
    def run(v, app):
        import db
        db.rename_budget_category(v['category_id'], v['name'])
        app.refresh_budget()
        return v

    @staticmethod
    def undo(v, app):
        import db
        db.rename_budget_category(v['category_id'], v['old'])
        app.refresh_budget()


@action('budget.cat_exclude', CHANGE, 'budget.cat_exclude {category_id, on: true(전체 합계에서 빼기)/false(넣기)}')
class BudgetCatExclude:
    @staticmethod
    def check(p, ctx):
        c = _cat(p, ctx)
        on = _to_bool(p.get('on'), True)
        if bool(c.get('excluded')) == on:
            raise Skip()
        return {'category_id': c['id'], 'on': on}, f'{c["name"]}: 전체 합계에서 {"빼기" if on else "넣기"}'

    @staticmethod
    def run(v, app):
        import db
        db.set_budget_category_excluded(v['category_id'], v['on'])
        app.refresh_budget()
        return v

    @staticmethod
    def undo(v, app):
        import db
        db.set_budget_category_excluded(v['category_id'], not v['on'])
        app.refresh_budget()


@action('budget.cat_group', CHANGE, 'budget.cat_group {category_id, group: 그룹 이름 또는 ""(그룹에서 빼기)} — 세부항목 그룹 (없는 그룹은 새로 만듦)')
class BudgetCatGroup:
    @staticmethod
    def check(p, ctx):
        c = _cat(p, ctx)
        group = _text(p, 'group', required=False, label='그룹 이름')
        cur = ctx['groups_by_id'].get(c.get('group_id'), {}).get('name', '')
        if group == cur:
            raise Skip()
        return ({'category_id': c['id'], 'group': group},
                f'{c["name"]}: ' + (f'그룹 「{group}」에 넣기' if group else '그룹에서 빼기'))

    @staticmethod
    def run(v, app):
        import db
        cur = next((c for c in db.get_budget_categories() if c['id'] == v['category_id']), None)
        if cur is None:
            return None
        created = None
        gid = None
        if v['group']:
            gid = next((g['id'] for g in db.get_budget_groups() if g['name'] == v['group']), None)
            if gid is None:
                gid = created = db.add_budget_group(v['group'])
        db.set_budget_category_group(cur['id'], gid)
        app.refresh_budget()
        return {'category_id': cur['id'], 'old': cur.get('group_id'), 'created': created}

    @staticmethod
    def undo(rec, app):
        import db
        if not rec:
            return
        db.set_budget_category_group(rec['category_id'], rec['old'])
        if rec.get('created'):
            db.delete_budget_group(rec['created'])
        app.refresh_budget()


@action('budget.cat_delete', DELETE, 'budget.cat_delete {category_id} — 세부항목 지우기 (안의 항목·사용 내역도 함께)')
class BudgetCatDelete:
    @staticmethod
    def check(p, ctx):
        c = _cat(p, ctx)
        items = [i for i in ctx['budget']['items'] if i['category_id'] == c['id']]
        return {'category_id': c['id']}, f'세부항목 지우기: {c["name"]} (안의 항목 {len(items)}개도 함께)'

    @staticmethod
    def run(v, app):
        import db
        cat = next((c for c in db.get_budget_categories() if c['id'] == v['category_id']), None)
        if cat is None:
            return None
        items = db.get_budget_items(cat['id'])
        ids = {i['id'] for i in items}
        ex = [e for e in db.get_budget_expenses() if e['item_id'] in ids]
        db.delete_budget_category(cat['id'])
        app.refresh_budget()
        return {'cat': dict(cat), 'items': items, 'expenses': ex}

    @staticmethod
    def undo(rec, app):
        import db
        if not rec:
            return
        cat = rec['cat']
        cid = db.add_budget_category(cat['name'])
        if cid is None:
            return
        if cat.get('excluded'):
            db.set_budget_category_excluded(cid, True)
        if cat.get('group_id') and any(g['id'] == cat['group_id'] for g in db.get_budget_groups()):
            db.set_budget_category_group(cid, cat['group_id'])
        for it in rec['items']:
            db.restore_budget_item({**it, 'category_id': cid}, [e for e in rec['expenses'] if e['item_id'] == it['id']])
        app.refresh_budget()


# ════════════════════════════════════════════════════════════════
# 공문 작성
# ════════════════════════════════════════════════════════════════
DOC_MAX = 8000


@action('document.write', CHANGE, 'document.write {title?, ref?: 관련 번호, body?: 고쳐 쓴 공문 본문 전체} — 공문 작성 창에 쓰기 '
        '(공문을 새로 쓰거나 고칠 때. 본문은 바뀐 곳만이 아니라 전체를 넣음)')
class DocumentWrite:
    @staticmethod
    def check(p, ctx):
        v = {}
        for key, limit in (('title', 200), ('ref', 120)):
            if p.get(key) not in (None, ''):
                v[key] = ' '.join(str(p[key]).split())[:limit]
        if p.get('body') not in (None, ''):
            v['body'] = str(p['body']).strip()[:DOC_MAX]          # 개인정보를 가린 채 온 글도 그대로 (예전과 같음)
        if not v:
            raise Skip()
        doc = ctx.get('document') or {}
        if all(v.get(k) == doc.get(k) for k in v):
            raise Skip()
        parts = []
        if 'title' in v:
            parts.append(f'제목 「{v["title"]}」')
        if 'ref' in v:
            parts.append(f'관련 번호 {v["ref"]}')
        if 'body' in v:
            parts.append(f'본문 {len(v["body"])}자')
        return v, '공문 작성: ' + ' · '.join(parts)

    @staticmethod
    def run(v, app):
        return app.write_document(v.get('title'), v.get('ref'), v.get('body'))

    @staticmethod
    def undo(old, app):
        if old:
            app.write_document(old['title'], old['ref'], old['body'])


# ════════════════════════════════════════════════════════════════
# 설정 · AI 대화 창
# ════════════════════════════════════════════════════════════════
@action('settings.alarm', CHANGE, 'settings.alarm {minutes: 0(끄기)|30|60|180} — 마감 임박 알림 간격')
class SettingsAlarm:
    @staticmethod
    def check(p, ctx):
        m = _to_int(p.get('minutes'))
        if m not in ALARM_MINUTES:
            raise Reject('알림 간격은 끄기·30분·1시간·3시간 중 하나예요')
        if m == ctx['settings'].get('alarm'):
            raise Skip()
        return {'minutes': m}, '마감 알림: ' + ('끄기' if m == 0 else f'{m // 60}시간마다' if m >= 60 else f'{m}분마다')

    @staticmethod
    def run(v, app):
        old = app.get_alarm()
        app.set_alarm(v['minutes'])
        return old

    @staticmethod
    def undo(old, app):
        if old is not None:
            app.set_alarm(old)


@action('settings.auto_update', CHANGE, 'settings.auto_update {on: true/false} — 자동 업데이트')
class SettingsAutoUpdate:
    @staticmethod
    def check(p, ctx):
        on = _to_bool(p.get('on'), True)
        if on == ctx['settings'].get('auto_update'):
            raise Skip()
        return {'on': on}, f'자동 업데이트 {"켜기" if on else "끄기"}'

    @staticmethod
    def run(v, app):
        import updater
        updater.set_auto_enabled(v['on'])
        return v['on']

    @staticmethod
    def undo(on, app):
        import updater
        updater.set_auto_enabled(not on)


@action('settings.autostart', CHANGE, 'settings.autostart {on: true/false} — Windows 켤 때 자동 실행')
class SettingsAutostart:
    @staticmethod
    def check(p, ctx):
        on = _to_bool(p.get('on'), True)
        if on == ctx['settings'].get('autostart'):
            raise Skip()
        return {'on': on}, f'시작 시 자동 실행 {"켜기" if on else "끄기"}'

    @staticmethod
    def run(v, app):
        app.set_autostart(v['on'])
        return v['on']

    @staticmethod
    def undo(on, app):
        app.set_autostart(not on)


@action('chat.side', CHANGE, 'chat.side {side: inside(업무창 안)|left|right|bottom} — AI 대화 붙는 자리')
class ChatSide:
    @staticmethod
    def check(p, ctx):
        s = str(p.get('side') or '').strip().lower()
        s = {v: k for k, v in CHAT_SIDES.items()}.get(s, s)        # 「왼쪽」처럼 우리말로 와도
        if s not in CHAT_SIDES:
            raise Reject('AI 대화 자리는 업무창 안·왼쪽·오른쪽·아래쪽 중 하나예요')
        if s == ctx['chat'].get('side'):
            raise Skip()
        return {'side': s}, f'AI 대화를 {CHAT_SIDES[s]}에'

    @staticmethod
    def run(v, app):
        old = app.chat.side
        app.chat.set_side(v['side'])
        return old

    @staticmethod
    def undo(old, app):
        app.chat.set_side(old)


@action('chat.style', CHANGE, 'chat.style {color?: 색 이름|"="(업무창과 같게), font?, size?, gap?} — AI 대화 배경색·글자·줄간격')
class ChatStyle:
    @staticmethod
    def check(p, ctx):
        c = p.get('color')
        color = '=' if str(c or '').strip() in ('=', '업무창과 같게') else _color(c)
        v = {'color': color, 'font': _font(p.get('font')), 'size': _size(p.get('size')), 'gap': _gap(p.get('gap'))}
        if all(x is None for x in v.values()):
            raise Skip()
        parts = [f'배경색 {"업무창과 같게" if color == "=" else COLOR_NAMES[color]}'] if color else []
        if v['font']:
            parts.append(f'글꼴 {FONT_KEYS[v["font"]]}')
        if v['size'] is not None:
            parts.append('글자 ' + next(k for k, x in SIZE_VALUES.items() if x == v['size']))
        if v['gap'] is not None:
            parts.append(f'줄 간격 {v["gap"]}')
        return v, 'AI 대화 꾸미기: ' + ' · '.join(parts)

    @staticmethod
    def run(v, app):
        ch = app.chat
        old = {'color': ch.color, 'font': ch.font_key, 'size': ch.scale, 'gap': ch.gap}
        if v['color'] is not None:
            ch.set_color(v['color'])
        if v['font']:
            ch.set_font(v['font'])
        if v['size'] is not None:
            ch.set_scale(v['size'])
        if v['gap'] is not None:
            ch.set_gap(v['gap'])
        return old

    @staticmethod
    def undo(old, app):
        ch = app.chat
        ch.set_color(old['color'])
        ch.set_font(old['font'])
        ch.set_scale(old['size'])
        ch.set_gap(old['gap'])


@action('chat.clear', VIEW, 'chat.clear {} — AI 대화 지우기')
class ChatClear:
    @staticmethod
    def check(p, ctx):
        return {}, '대화 지우기'

    @staticmethod
    def run(v, app):
        app.clear_chat_later()


# ════════════════════════════════════════════════════════════════
# 지금 상태 · 글 만들기 · 답 읽기 · 실행
# ════════════════════════════════════════════════════════════════
def build_state(app, doc_extra=None) -> dict:
    """지금 앱 상태 (DB + 열린 창). 검사(check)와 AI에게 보낼 글에 쓴다.
    doc_extra: 공문 작성 창에서 물을 때 함께 보낼 {'received': 받은 공문 정리, 'notes': 나만의 작성 규칙}."""
    import db
    rows = sorted(db.get_all_windows(), key=lambda w: w['id'])
    task_n = memo_n = 0
    windows = []
    for r in rows:
        kind = 'memo' if r.get('memo_mode') else 'task'
        if kind == 'memo':
            memo_n += 1
            label = f'메모장 {memo_n}'
        else:
            task_n += 1
            label = f'업무창 {task_n}'
        live = app.window(r['id']) if app is not None else None
        windows.append({'id': r['id'], 'label': label, 'kind': kind, 'color': (r.get('color') or '').upper(),
                        'font': r.get('font') or 'pretendard', 'scale': float(r.get('scale') or 1.0),
                        'gap': int(r.get('row_gap') or 0), 'collapsed': bool(r.get('collapsed')),
                        'pin': bool(getattr(live, 'pin_active', False)) if live is not None else False,
                        'memo': (live.memo_editor.toPlainText() if live is not None and kind == 'memo'
                                 else r.get('memo_text') or '')})
    tasks = []
    for w in windows:
        if w['kind'] == 'task':
            tasks.extend(db.get_tasks(w['id']))
    cats, items, groups = db.get_budget_categories(), db.get_budget_items(), db.get_budget_groups()
    settings = {}
    if app is not None:
        settings = {'alarm': app.get_alarm(), 'auto_update': app.get_auto_update(), 'autostart': app.get_autostart()}
    chat_obj = getattr(app, '_chat', None)
    chat = {'side': getattr(chat_obj, 'side', None) or _saved_chat_side()}
    document = None
    try:
        document = app.document_state() if app is not None else None
    except Exception:
        document = None
    if document is not None and doc_extra:
        document = {**document, **doc_extra}
    return {
        'document': document,
        'today': date.today(), 'now': datetime.now(),
        'current_window': getattr(app, 'host_window_id', None),
        'windows': windows, 'windows_by_id': {w['id']: w for w in windows},
        'tasks': tasks, 'tasks_by_id': {t['id']: t for t in tasks},
        'budget': {'cats': cats, 'items': items, 'groups': groups},
        'cats_by_id': {c['id']: c for c in cats}, 'items_by_id': {i['id']: i for i in items},
        'groups_by_id': {g['id']: g for g in groups},
        'settings': settings, 'chat': chat,
    }


SYSTEM = (
    "당신은 한국 고등학교 선생님이 쓰는 업무 메모 앱 「서서니 노트」를 대신 다뤄 주는 도우미입니다.\n"
    "아래 '지금 앱 상태'만 근거로 답하고, 사용자가 시키면 '할 수 있는 기능' 중에서 골라 actions에 넣습니다.\n"
    "- 답(answer)은 짧고 분명하게, 목록이 필요하면 '- '로 시작하는 줄로. 마크다운(**, ##, 표)은 쓰지 않습니다.\n"
    "- 날짜는 '11월 21일(금)'처럼, 금액은 '1,234,000원'처럼 씁니다. 상태에 없는 것은 지어내지 않습니다.\n"
    "- '내일', '다음 주 금요일' 같은 말은 위의 오늘 날짜로 계산해 'YYYY-MM-DD'로 넣습니다.\n"
    "- 업무·창·예산 항목·세부항목은 상태에 적힌 id로 가리킵니다. answer에는 id를 절대 쓰지 말고 이름으로 말합니다.\n"
    "- 어느 것인지 애매하면 actions를 비우고 answer에서 후보를 보여 주며 되묻습니다.\n"
    "- 바꾸기·지우기는 사용자가 확인한 뒤에 앱이 실행합니다. answer에는 무엇을 할지 한 줄로 확인하듯 적습니다.\n"
    "- 예산 금액 계산은 하지 마세요. 쓴 금액·목표 잔액·목표 집행액만 넣으면 앱이 계산합니다.\n"
    "- 할 수 없는 일(예산 전체 지우기, 업무창·메모장 자체 지우기, AI 연결·API 키·개인정보 보호 설정, 업데이트, 앱 종료)은\n"
    "  actions에 넣지 말고 answer에서 메뉴 위치를 안내합니다.\n"
    f"- actions는 한 번에 {MAX_ACTIONS}개까지입니다. 그보다 많으면 나눠서 하자고 안내합니다.\n"
    "반드시 JSON만 출력합니다."
)

DOC_RULES = (
    "공문을 쓰거나 고칠 때(document.write)는 한국 학교 행정 공문 문체(개조식, '~함', '~바람', '~고자 함')로 쓰고,\n"
    "본문(body)에는 고친 곳만이 아니라 공문 본문 전체를 넣습니다. 지금 쓰는 공문이 있으면 그 내용을 바탕으로 고칩니다.\n"
    "받은 공문 정리·나만의 작성 규칙이 있으면 따릅니다. 공문 내용 자체는 answer에 다시 쓰지 말고 한 줄로 무엇을 했는지만 적습니다."
)

_FORMAT = '{"answer": "답", "actions": [{"do": "기능 이름", "...": "값"}]}'


def catalog() -> str:
    lines = []
    for spec in _ACTIONS.values():
        tag = {'view': '바로', 'change': '확인 후', 'delete': '지우기 확인 후'}[spec['risk']]
        lines.append(f'- [{tag}] {spec["help"]}')
    colors = ', '.join(COLOR_NAMES.values())
    return '\n'.join(lines) + f'\n(색 이름: {colors})'


def _saved_chat_side():
    try:
        from ui.ai_chat_window import saved_side
        return saved_side()
    except Exception:
        return 'inside'


def state_text(ctx) -> str:
    out = ['[창]']
    for w in ctx['windows']:
        extra = []
        if w['id'] == ctx.get('current_window'):
            extra.append('지금 창')
        if w['collapsed']:
            extra.append('접힘')
        if w['pin']:
            extra.append('항상 위')
        color = COLOR_NAMES.get(w['color'], w['color'] or '노랑')
        line = (f'window_id={w["id"]} | {w["label"]} | 배경 {color} | 글꼴 {FONT_KEYS.get(w["font"], w["font"])}'
                f' | 줄간격 {w["gap"]}' + (f' | {", ".join(extra)}' if extra else ''))
        if w['kind'] == 'memo':
            memo = ' '.join(w['memo'].split())
            line += f' | 글: {memo[:200]}{"…" if len(memo) > 200 else ""}' if memo else ' | 글 없음'
        out.append(line)
    labels = {w['id']: w['label'] for w in ctx['windows']}
    out.append('\n[업무]')
    if not ctx['tasks']:
        out.append('(업무 없음)')
    for t in ctx['tasks']:
        parts = [f'task_id={t["id"]}', labels.get(t['window_id'], '업무창'),
                 str(t.get('name') or '').replace('\n', ' '),
                 '마감 ' + (f'{t["deadline"]} ({fmt_deadline(t["deadline"])})' if t.get('deadline') else '없음')]
        if t.get('priority'):
            parts.append('중요')
        if t.get('recurrence'):
            parts.append('반복 ' + REPEATS.get(t['recurrence'], '사용자 지정'))
        if t.get('strikethrough'):
            parts.append('완료')
        if t.get('notes'):
            n = ' '.join(t['notes'].split())
            parts.append(f'메모: {n[:80]}{"…" if len(n) > 80 else ""}')
        if t.get('related_no'):
            parts.append(f'관련번호 {t["related_no"]}')
        out.append(' | '.join(parts))
    b = ctx['budget']
    if b['cats']:
        out.append('\n[예산]')
        gname = {g['id']: g['name'] for g in b['groups']}
        tb = ts = 0
        for c in b['cats']:
            its = [i for i in b['items'] if i['category_id'] == c['id']]
            cb, cs = sum(i['budget'] for i in its), sum(i['spent'] for i in its)
            if not c.get('excluded'):
                tb, ts = tb + cb, ts + cs
            tags = (' (전체 합계에서 제외)' if c.get('excluded') else '') + \
                (f' (그룹 {gname[c["group_id"]]})' if c.get('group_id') in gname else '')
            out.append(f'category_id={c["id"]} | 세부항목 {c["name"]}{tags} | 예산 {cb:,} · 집행 {cs:,} · 잔액 {cb - cs:,}')
            for i in its[:120]:
                done = ' | 완료' if i.get('completed') else ''
                out.append(f'  item_id={i["id"]} | {i.get("cost_type", "")} | {i.get("description", "")} | '
                           f'예산 {i["budget"]:,} | 집행 {i["spent"]:,} | 잔액 {i["budget"] - i["spent"]:,}{done}')
        out.append(f'[예산 전체] 예산 {tb:,} · 집행 {ts:,} · 잔액 {tb - ts:,}')
    d = ctx.get('document')
    if d is not None:
        out.append('\n[공문 작성 창 — 지금 쓰는 공문]')
        out.append(f'제목: {d.get("title") or "(없음)"}')
        out.append(f'관련 번호: {d.get("ref") or "(없음)"}')
        body = d.get('body') or ''
        out.append('본문:\n' + (body[:DOC_MAX] if body.strip() else '(비어 있음)'))
        if d.get('received'):
            out.append('\n[받은 공문 정리]\n' + d['received'])
        if d.get('notes'):
            out.append('\n' + d['notes'])
    s = ctx.get('settings') or {}
    if s:
        out.append(f'\n[설정] 마감 알림 {s.get("alarm")}분마다(0=끔) · 자동 업데이트 {"켜짐" if s.get("auto_update") else "꺼짐"}'
                   f' · 시작 시 자동 실행 {"켜짐" if s.get("autostart") else "꺼짐"}'
                   f' · AI 대화 자리 {CHAT_SIDES.get(ctx["chat"]["side"], "")}')
    return '\n'.join(out)


def build_prompt(question, ctx, history=()) -> str:
    now = ctx.get('now') or datetime.now()
    today = ctx['today']
    past = ''.join(f"\n이전 질문: {q}\n이전 답: {a}" for q, a in list(history)[-2:])
    rules = f"\n\n공문 쓰기 규칙\n{DOC_RULES}" if ctx.get('document') is not None else ''
    return (f"오늘: {today.isoformat()} ({_WEEKDAYS[today.weekday()]}요일), 지금 시각 {now:%H:%M}\n\n"
            f"할 수 있는 기능\n{catalog()}{rules}\n\n지금 앱 상태\n{state_text(ctx)}\n{past}\n\n"
            f"질문: {question}\n\n다음 형식의 JSON만 출력하세요:\n{_FORMAT}")


def parse_reply(raw, ctx) -> dict:
    """AI 응답 → {'answer', 'now': [바로 실행], 'changes': [확인 후], 'deletes': [지우기 확인 후], 'rejected': [...]}.
    항목: {'do': 기능, 'v': 검사한 값, 'desc': 설명}. 거절: {'do', 'desc', 'reason'}."""
    data = _extract_json(raw)
    if not isinstance(data, dict):
        return {'answer': (raw or '').strip(), 'now': [], 'changes': [], 'deletes': [], 'rejected': []}
    out = {'answer': str(data.get('answer') or '').strip(), 'now': [], 'changes': [], 'deletes': [], 'rejected': []}
    acts = data.get('actions') or []
    if not isinstance(acts, list):
        acts = []
    seen = set()
    for i, a in enumerate(acts):
        if not isinstance(a, dict):
            continue
        name = str(a.get('do') or a.get('action') or '').strip()
        if i >= MAX_ACTIONS:
            out['rejected'].append({'do': name, 'desc': name, 'reason': f'한 번에 {MAX_ACTIONS}개까지만 할 수 있어요'})
            continue
        spec = _ACTIONS.get(name)
        if spec is None:
            out['rejected'].append({'do': name, 'desc': name or '(이름 없음)', 'reason': '앱에 없는 기능이에요'})
            continue
        try:
            v, desc = spec['check'](a, ctx)
        except Skip:
            continue
        except Reject as e:
            out['rejected'].append({'do': name, 'desc': _reject_label(name, a, ctx), 'reason': str(e)})
            continue
        key = (name, json.dumps(v, sort_keys=True, ensure_ascii=False, default=str))
        if key in seen:
            continue
        seen.add(key)
        bucket = {VIEW: 'now', CHANGE: 'changes', DELETE: 'deletes'}[spec['risk']]
        out[bucket].append({'do': name, 'v': v, 'desc': desc})
    return out


def _reject_label(name, a, ctx):
    t = ctx['tasks_by_id'].get(_to_int(a.get('task_id')))
    if t:
        return t['name']
    it = ctx['items_by_id'].get(_to_int(a.get('item_id')))
    if it:
        return _item_label(it)
    for k in ('name', 'description', 'title'):
        if a.get(k):
            return str(a[k])[:30]
    return name


def execute(items, app) -> list:
    """검사를 통과한 항목들을 실행 → 되돌리기 기록 [(기능, 기록)]."""
    journal = []
    for it in items:
        spec = _ACTIONS[it['do']]
        rec = spec['run'](it['v'], app)
        if spec['risk'] != VIEW and spec['undo'] is not None:
            journal.append((it['do'], rec))
    return journal


def undo(journal, app):
    """execute 전으로 (나중에 한 것부터)."""
    for name, rec in reversed(journal or []):
        if rec is not None:
            _ACTIONS[name]['undo'](rec, app)


# ── 실행 (뒤에서) ────────────────────────────────────────────────
try:
    from PyQt5.QtCore import QThread, pyqtSignal
except ImportError:          # 테스트 환경 등
    QThread = object

    def pyqtSignal(*_a):     # noqa: N802
        return None


class AskWorker(QThread):
    """질문 → AI 답.  done(reply: dict, masked: dict) / failed(message: str). ctx는 화면 쪽에서 만들어 넘긴다."""
    done = pyqtSignal(dict, dict)
    failed = pyqtSignal(str)

    def __init__(self, question, ctx, history=(), parent=None):
        super().__init__(parent)
        self._q = question
        self._ctx = ctx
        self._history = list(history)

    def run(self):
        try:
            from ai_client import get_adapter
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있어요. 메뉴 → 설정 → AI 연결에서 연결하세요.')
                return
            if not adapter.get_api_key():
                self.failed.emit('AI API 키가 없어요. 메뉴 → 설정 → AI 연결에서 키를 넣어 주세요.')
                return
            prompt = build_prompt(self._q, self._ctx, self._history)
            masked = {}
            try:
                from ai_client import masking_required
                from school_profile import mask_for_ai
                if masking_required():
                    prompt, masked = mask_for_ai(prompt)
            except Exception:
                pass
            raw = adapter.complete(SYSTEM, prompt, max_tokens=2500)
            reply = parse_reply(raw, self._ctx)
            if not any(reply[k] for k in ('answer', 'now', 'changes', 'deletes', 'rejected')):
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
