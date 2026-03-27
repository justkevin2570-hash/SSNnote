"""
익명 인증 모듈 (Supabase Anonymous Auth)
- 앱 시작 시 init() 호출
- 저장된 토큰이 있으면 갱신, 없으면 새 익명 세션 생성
- 실패해도 앱은 로컬 전용으로 정상 동작
"""

import urllib.request
import urllib.error
import json
import os

try:
    from supabase_secret import SUPABASE_URL, SUPABASE_ANON_KEY
except ImportError:
    SUPABASE_URL = ''
    SUPABASE_ANON_KEY = ''

_AUTH_FILE = os.path.join(os.environ.get('APPDATA', '.'), 'Tmemo', 'auth.json')
_session = None


def _post(path, data):
    url = f'{SUPABASE_URL}{path}'
    headers = {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
    }
    body = json.dumps(data).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method='POST')
    with urllib.request.urlopen(req, timeout=10) as r:
        content = r.read()
        return json.loads(content) if content else {}


def _load_saved():
    try:
        with open(_AUTH_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def _save(session):
    os.makedirs(os.path.dirname(_AUTH_FILE), exist_ok=True)
    with open(_AUTH_FILE, 'w', encoding='utf-8') as f:
        json.dump(session, f)


def init():
    """
    앱 시작 시 호출. 세션 초기화.
    - 저장된 refresh_token으로 갱신 시도
    - 실패 시 새 익명 세션 생성
    - 오프라인/오류 시 None 반환 (앱은 정상 동작)
    """
    global _session
    if not SUPABASE_URL:
        return None

    saved = _load_saved()
    if saved and saved.get('refresh_token'):
        try:
            new_session = _post(
                '/auth/v1/token?grant_type=refresh_token',
                {'refresh_token': saved['refresh_token']}
            )
            if new_session.get('access_token'):
                _session = new_session
                _save(_session)
                return _session
        except Exception:
            pass  # 갱신 실패 → 새 익명 로그인 시도

    try:
        session = _post('/auth/v1/signup', {})
        if session.get('access_token'):
            _session = session
            _save(_session)
            return _session
    except Exception:
        pass

    return None


def get_access_token():
    return _session.get('access_token') if _session else None


def get_user_id():
    if _session and 'user' in _session:
        return _session['user']['id']
    return None


def is_authenticated():
    return _session is not None and bool(_session.get('access_token'))
