"""claude_connect.py — 예전 「Claude 연결」 기능이 남긴 등록 지우기 (UI 의존 없음).

「Claude 연결」(MCP 커넥터)은 v3.06에서 메뉴에서 빠졌고, 연결 프로그램(SSNnote_MCP.exe)도 더 이상 설치하지 않는다.
예전에 연결해 둔 사람은 Claude 데스크톱 앱 설정에 `ssnnote` 커넥터가 남아 「커넥터 실행이 중지되었습니다」
알림이 계속 뜨므로, 서서니 노트를 켤 때 그 등록만 지운다.
  - Claude 데스크톱:   claude_desktop_config.json 의 mcpServers.ssnnote 항목 (다른 설정은 그대로, 원본은 백업)
  - Claude Code:      `claude mcp remove --scope user ssnnote` (한 번만)
"""

import glob
import json
import os
import shutil
import subprocess
import sys

SERVER_NAME = 'ssnnote'
BACKUP_SUFFIX = '.ssnnote-backup'
CODE_CLEANUP_KEY = 'claude_code_mcp_removed'


# ── Claude 데스크톱 앱 ────────────────────────────────────────

def desktop_config_dirs():
    """Claude 데스크톱 앱 설정 폴더 후보 중 실제로 있는 것.

    - 요즘 스토어판: %LOCALAPPDATA%\\Claude-Data\\claude_<게시자ID>
    - 일반 설치:     %APPDATA%\\Claude
    - 예전 스토어판: %LOCALAPPDATA%\\Packages\\Claude_*\\LocalCache\\Roaming\\Claude
    """
    dirs = []
    appdata, local = os.environ.get('APPDATA', ''), os.environ.get('LOCALAPPDATA', '')
    if local:
        dirs += sorted(glob.glob(os.path.join(local, 'Claude-Data', '*')))
    if appdata:
        dirs.append(os.path.join(appdata, 'Claude'))
    if local:
        dirs += sorted(glob.glob(os.path.join(local, 'Packages', 'Claude_*', 'LocalCache', 'Roaming', 'Claude')))
    return [d for d in dirs if os.path.isdir(d)]


def desktop_config_paths():
    """Claude 앱이 읽을 수 있는 설정 파일 경로 전부 (스토어판은 두 군데)."""
    return [os.path.join(d, 'claude_desktop_config.json') for d in desktop_config_dirs()]


def _read_config(path):
    """설정 내용(dict). 없거나 비었으면 {}, 형식이 이상하면 None (그 파일은 건드리지 않음)."""
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding='utf-8-sig') as f:
            text = f.read()
    except OSError:
        return None
    if not text.strip():
        return {}
    try:
        data = json.loads(text)
    except ValueError:
        return None
    return data if isinstance(data, dict) else None


def _write_config(path, data):
    shutil.copy2(path, path + BACKUP_SUFFIX)                             # 바꾸기 전 원본 백업
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def remove_desktop_entries(paths=None):
    """모든 설정 파일에서 ssnnote 커넥터만 지운다. 지운 파일 경로 목록."""
    removed = []
    for p in (desktop_config_paths() if paths is None else paths):
        data = _read_config(p)
        servers = (data or {}).get('mcpServers')
        if isinstance(servers, dict) and SERVER_NAME in servers:
            del servers[SERVER_NAME]
            try:
                _write_config(p, data)
                removed.append(p)
            except OSError as e:
                print(f'[claude] 커넥터 등록 지우기 실패: {p}: {e}')
    return removed


# ── Claude Code ────────────────────────────────────────────────

def find_claude_code():
    """claude 명령 경로 (없으면 None)."""
    found = shutil.which('claude')
    if found:
        return found
    home = os.path.expanduser('~')
    candidates = [
        os.path.join(home, '.local', 'bin', 'claude.exe'),
        os.path.join(home, '.local', 'bin', 'claude'),
        os.path.join(os.environ.get('APPDATA', ''), 'npm', 'claude.cmd'),
    ]
    return next((c for c in candidates if c and os.path.exists(c)), None)


def remove_claude_code_entry(claude=None):
    """Claude Code에 등록된 ssnnote 커넥터를 지운다. 지웠거나 원래 없으면 True."""
    claude = claude or find_claude_code()
    if not claude:
        return True
    kw = {'cwd': os.path.expanduser('~')}
    if sys.platform == 'win32':
        kw['creationflags'] = 0x08000000                 # CREATE_NO_WINDOW: 검은 창이 뜨지 않게
    try:
        p = subprocess.run([claude, 'mcp', 'remove', '--scope', 'user', SERVER_NAME], capture_output=True,
                           text=True, encoding='utf-8', errors='replace', timeout=60, **kw)
    except (OSError, subprocess.TimeoutExpired):
        return False
    out = ((p.stdout or '') + (p.stderr or '')).lower()
    return p.returncode == 0 or 'no mcp server' in out or 'not found' in out


# ── 앱을 켤 때 ─────────────────────────────────────────────────

def cleanup_old_connector():
    """예전 「Claude 연결」 등록을 지운다. 데스크톱 설정은 매번 확인(가벼움), Claude Code는 성공할 때까지 한 번."""
    removed = remove_desktop_entries()
    if removed:
        print(f'[claude] 예전 커넥터 등록 지움: {removed}')
    import db
    if db.get_setting(CODE_CLEANUP_KEY, '') != '1' and remove_claude_code_entry():
        db.set_setting(CODE_CLEANUP_KEY, '1')
    return removed
