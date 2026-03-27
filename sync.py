"""
Supabase 동기화 모듈 (Push-only, v1)
- 로컬 SQLite → Supabase 단방향 동기화
- 전체 교체 방식: 기존 클라우드 데이터 삭제 후 재삽입
- 실패해도 예외 발생 안 함 (로컬은 항상 정상 동작)
"""

import urllib.request
import urllib.error
import json

try:
    from supabase_secret import SUPABASE_URL, SUPABASE_ANON_KEY
except ImportError:
    SUPABASE_URL = ''
    SUPABASE_ANON_KEY = ''

import auth


def _req(method, table, data=None, query=''):
    token = auth.get_access_token()
    if not token or not SUPABASE_URL:
        return None

    url = f'{SUPABASE_URL}/rest/v1/{table}{query}'
    headers = {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': f'Bearer {token}',
        'Prefer': 'return=minimal',
    }
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            content = r.read()
            return json.loads(content) if content else []
    except Exception:
        return None


def push_all(windows, tasks_by_window):
    """
    전체 데이터를 Supabase에 push (전체 교체).
    windows: list of window dicts (from get_all_windows)
    tasks_by_window: {window_id: [task dicts]}
    """
    user_id = auth.get_user_id()
    if not user_id:
        return

    # 기존 클라우드 데이터 삭제 (tasks 먼저, windows 나중)
    _req('DELETE', 'tasks',   query=f'?user_id=eq.{user_id}')
    _req('DELETE', 'windows', query=f'?user_id=eq.{user_id}')

    # 창 일괄 삽입
    if windows:
        window_records = [
            {
                'user_id':   user_id,
                'local_id':  w['id'],
                'x':         w['x'],
                'y':         w['y'],
                'width':     w['width'],
                'height':    w['height'],
                'collapsed': bool(w.get('collapsed', False)),
                'color':     w.get('color', ''),
            }
            for w in windows
        ]
        _req('POST', 'windows', data=window_records)

    # 태스크 일괄 삽입
    all_tasks = []
    for window_id, tasks in tasks_by_window.items():
        for t in tasks:
            all_tasks.append({
                'user_id':        user_id,
                'local_id':       t['id'],
                'window_local_id': window_id,
                'name':           t['name'],
                'deadline':       t.get('deadline', ''),
                'strikethrough':  bool(t.get('strikethrough', False)),
            })
    if all_tasks:
        _req('POST', 'tasks', data=all_tasks)
