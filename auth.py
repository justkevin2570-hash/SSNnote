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
import sys

try:
    from supabase_secret import SUPABASE_URL, SUPABASE_ANON_KEY
except ImportError:
    SUPABASE_URL = ''
    SUPABASE_ANON_KEY = ''

_AUTH_FILE = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'auth.json')
_session = None


# ── Windows DPAPI 암호화 (Windows 전용, 미지원 시 평문 폴백) ──
def _dpapi_available():
    # Windows에서만 DPAPI 사용. WSL/리눅스/맥은 평문 폴백.
    if sys.platform != 'win32':
        return False
    try:
        import ctypes
        ctypes.windll.crypt32  # Windows에서만 존재
    except Exception:
        return False
    # 실제 라운드트립이 동작하는지 확인 (일부 WSL/wine 환경에선 DLL은 있으나 실패)
    try:
        probe = _protect_raw('__dpapi_probe__')
        if probe is None:
            return False
        back = _unprotect_raw(probe)
        return back == '__dpapi_probe__'
    except Exception:
        return False


def _protect_raw(plain: str):
    import ctypes
    from ctypes import wintypes, POINTER
    blob_in = ctypes.create_string_buffer(plain.encode('utf-8'))
    class DATA_BLOB(ctypes.Structure):
        _fields_ = [('cbData', wintypes.DWORD), ('pbData', POINTER(ctypes.c_char))]
    out = DATA_BLOB()
    CryptProtectData = ctypes.windll.crypt32.CryptProtectData
    CryptProtectData.argtypes = [POINTER(DATA_BLOB), ctypes.c_wchar_p, ctypes.c_void_p,
                                ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, POINTER(DATA_BLOB)]
    CryptProtectData.restype = wintypes.BOOL
    in_blob = DATA_BLOB(len(blob_in), ctypes.cast(blob_in, POINTER(ctypes.c_char)))
    if CryptProtectData(ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out)):
        size = out.cbData
        buf = (ctypes.c_char * size).from_address(ctypes.addressof(out.pbData.contents))
        return 'dpapi:' + bytes(buf).hex()
    return None


def _unprotect_raw(payload: str):
    import ctypes
    from ctypes import wintypes, POINTER
    raw = bytes.fromhex(payload[len('dpapi:'):])
    class DATA_BLOB(ctypes.Structure):
        _fields_ = [('cbData', wintypes.DWORD), ('pbData', POINTER(ctypes.c_char))]
    in_blob = DATA_BLOB(len(raw), ctypes.cast(ctypes.create_string_buffer(raw, len(raw)), POINTER(ctypes.c_char)))
    out = DATA_BLOB()
    CryptUnprotectData = ctypes.windll.crypt32.CryptUnprotectData
    CryptUnprotectData.argtypes = [POINTER(DATA_BLOB), POINTER(ctypes.c_wchar_p), ctypes.c_void_p,
                                  ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, POINTER(DATA_BLOB)]
    CryptUnprotectData.restype = wintypes.BOOL
    if CryptUnprotectData(ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out)):
        size = out.cbData
        buf = (ctypes.c_char * size).from_address(ctypes.addressof(out.pbData.contents))
        return bytes(buf).decode('utf-8', errors='replace')
    return None


_DPAPI_OK = _dpapi_available()


def _protect(plain: str) -> str:
    """DPAPI로 암호화. 미지원/실패 시 평문 그대로 반환."""
    if not _DPAPI_OK:
        return plain
    try:
        res = _protect_raw(plain)
        return res if res is not None else plain
    except Exception:
        return plain


def _unprotect(payload: str) -> str:
    """DPAPI로 복호화. 'dpapi:' 접두사가 없으면 평문으로 간주."""
    if not payload.startswith('dpapi:'):
        return payload
    if not _DPAPI_OK:
        # Windows가 아닌 환경에서는 복호화 불가 → 빈 값 반환
        return ''
    try:
        res = _unprotect_raw(payload)
        return res if res is not None else ''
    except Exception:
        return ''


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
            raw = f.read().strip()
        if not raw:
            return None
        plain = _unprotect(raw)
        if not plain:
            return None
        return json.loads(plain)
    except Exception:
        return None


def _save(session):
    os.makedirs(os.path.dirname(_AUTH_FILE), exist_ok=True)
    try:
        data = json.dumps(session)
        encrypted = _protect(data)
        with open(_AUTH_FILE, 'w', encoding='utf-8') as f:
            f.write(encrypted)
    except Exception:
        pass


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
