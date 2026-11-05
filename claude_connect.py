"""claude_connect.py — 「Claude 연결」 버튼 뒤의 로직 (UI 의존 없음).

서서니 노트 MCP 서버를 Claude Code / Claude 데스크톱 앱에 등록·해제하고 상태를 확인한다.
  - Claude Code:      `claude mcp add --scope user ssnnote -- <서버 실행 명령>`
  - Claude 데스크톱:   claude_desktop_config.json 의 mcpServers.ssnnote 항목 (원본은 백업)

서버 실행 명령
  - 설치판(exe):  설치 폴더의 SSNnote_MCP.exe
  - 소스 실행:    지금 쓰는 python + mcp_server.py (mcp 패키지 필요)
"""

import glob
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys

SERVER_NAME = 'ssnnote'
MCP_EXE_NAME = 'SSNnote_MCP.exe'
BACKUP_SUFFIX = '.ssnnote-backup'
_ROOT = os.path.dirname(os.path.abspath(__file__))


class ConnectError(Exception):
    """사용자에게 그대로 보여줄 한국어 오류."""


# ── 서버 실행 명령 ─────────────────────────────────────────────

def server_command():
    """(command, args) — 연결 프로그램을 실행하는 명령. 준비가 안 됐으면 ConnectError."""
    if getattr(sys, 'frozen', False):
        exe = os.path.join(os.path.dirname(sys.executable), MCP_EXE_NAME)
        if not os.path.exists(exe):
            raise ConnectError('설치 폴더에 연결 프로그램(SSNnote_MCP.exe)이 없습니다.\n'
                               '최신 버전의 서서니 노트를 다시 설치해 주세요.')
        return exe, []
    if importlib.util.find_spec('mcp') is None:
        raise ConnectError('연결에 필요한 구성요소(mcp)가 설치되어 있지 않습니다.')
    py = sys.executable
    base = os.path.basename(py).lower()
    if base.startswith('pythonw'):                      # 창 없는 python → 콘솔 python
        alt = os.path.join(os.path.dirname(py), base.replace('pythonw', 'python'))
        if os.path.exists(alt):
            py = alt
    return py, [os.path.join(_ROOT, 'mcp_server.py')]


def install_requirements_command():
    """소스 실행 환경에서 mcp 패키지를 설치하는 명령 (설치판에는 필요 없음)."""
    return [sys.executable, '-m', 'pip', 'install', '-r', os.path.join(_ROOT, 'requirements-mcp.txt')]


def _run(cmd, timeout=90):
    # 홈 폴더에서 실행: claude는 폴더마다 설정(local scope)이 달라서, 어디서 켜도 같은 결과가 나오게
    kw = {'cwd': os.path.expanduser('~')}
    if sys.platform == 'win32':
        kw['creationflags'] = 0x08000000                 # CREATE_NO_WINDOW: 검은 창이 뜨지 않게
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace',
                           timeout=timeout, **kw)
    except FileNotFoundError:
        raise ConnectError('프로그램을 실행하지 못했습니다: ' + str(cmd[0]))
    except subprocess.TimeoutExpired:
        raise ConnectError('응답이 너무 늦습니다. 잠시 후 다시 시도해 주세요.')
    return p.returncode, (p.stdout or '') + (p.stderr or '')


# ── Claude Code ────────────────────────────────────────────────

def find_claude_code():
    """claude 명령 경로 (없으면 None). 로그인 직후 켜진 앱은 PATH가 오래됐을 수 있어 흔한 위치도 찾는다."""
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


def claude_code_status(claude=None):
    """{'installed', 'registered', 'connected', 'detail'}."""
    claude = claude or find_claude_code()
    if not claude:
        return {'installed': False, 'registered': False, 'connected': False, 'detail': ''}
    code, out = _run([claude, 'mcp', 'get', SERVER_NAME])
    registered = code == 0 and 'no mcp server' not in out.lower()
    m = re.search(r'^\s*Status:\s*(.+)$', out, re.M | re.I)              # 예: "Status: √ Connected"
    status = m.group(1) if m else ''
    connected = registered and bool(re.search(r'connected', status, re.I)) \
        and not re.search(r'fail|✗|not connected|disconnected', status, re.I)
    return {'installed': True, 'registered': registered, 'connected': connected, 'detail': out.strip()}


def connect_claude_code(claude=None):
    claude = claude or find_claude_code()
    if not claude:
        raise ConnectError('Claude Code가 설치되어 있지 않습니다.')
    cmd, args = server_command()
    _run([claude, 'mcp', 'remove', '--scope', 'user', SERVER_NAME])      # 예전 경로로 등록돼 있으면 교체
    code, out = _run([claude, 'mcp', 'add', '--scope', 'user', SERVER_NAME, '--', cmd, *args])
    if code != 0:
        raise ConnectError('Claude Code 등록에 실패했습니다.\n' + out.strip()[-400:])
    return claude_code_status(claude)


def disconnect_claude_code(claude=None):
    claude = claude or find_claude_code()
    if not claude:
        raise ConnectError('Claude Code가 설치되어 있지 않습니다.')
    code, out = _run([claude, 'mcp', 'remove', '--scope', 'user', SERVER_NAME])
    if code != 0 and 'no mcp server' not in out.lower():          # 이미 없으면 그대로 성공
        raise ConnectError('연결 해제에 실패했습니다.\n' + out.strip()[-400:])
    return claude_code_status(claude)


# ── Claude 데스크톱 앱 ────────────────────────────────────────

def desktop_config_path():
    """Claude 데스크톱 앱 설정 파일 경로 (앱이 없으면 None).

    일반 설치는 %APPDATA%\\Claude, Microsoft Store 설치는 %LOCALAPPDATA%\\Packages\\Claude_*\\...\\Claude.
    """
    dirs = []
    appdata, local = os.environ.get('APPDATA', ''), os.environ.get('LOCALAPPDATA', '')
    if appdata:
        dirs.append(os.path.join(appdata, 'Claude'))
    if local:
        dirs += sorted(glob.glob(os.path.join(local, 'Packages', 'Claude_*', 'LocalCache', 'Roaming', 'Claude')))
    existing = [d for d in dirs if os.path.isdir(d)]
    with_file = [d for d in existing if os.path.exists(os.path.join(d, 'claude_desktop_config.json'))]
    chosen = (with_file or existing or [None])[0]
    return os.path.join(chosen, 'claude_desktop_config.json') if chosen else None


def _read_config(path):
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8-sig') as f:
        text = f.read()
    if not text.strip():
        return {}
    try:
        data = json.loads(text)
    except ValueError:
        raise ConnectError('Claude 앱 설정 파일의 형식이 올바르지 않아 건드리지 않았습니다.\n' + path)
    if not isinstance(data, dict):
        raise ConnectError('Claude 앱 설정 파일의 형식이 예상과 달라 건드리지 않았습니다.\n' + path)
    return data


def _write_config(path, data):
    if os.path.exists(path):
        shutil.copy2(path, path + BACKUP_SUFFIX)                         # 바꾸기 전 원본 백업
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def desktop_status(path=None):
    path = path or desktop_config_path()
    if not path:
        return {'installed': False, 'registered': False, 'up_to_date': False, 'path': ''}
    try:
        entry = (_read_config(path).get('mcpServers') or {}).get(SERVER_NAME)
    except ConnectError:
        entry = None
    up_to_date = False
    if entry:
        try:
            cmd, args = server_command()
            up_to_date = entry.get('command') == cmd and entry.get('args', []) == args
        except ConnectError:
            pass
    return {'installed': True, 'registered': bool(entry), 'up_to_date': up_to_date, 'path': path}


def connect_desktop(path=None):
    path = path or desktop_config_path()
    if not path:
        raise ConnectError('Claude 데스크톱 앱이 설치되어 있지 않습니다.')
    cmd, args = server_command()
    data = _read_config(path)
    servers = data.get('mcpServers')
    if not isinstance(servers, dict):
        servers = {}
    servers[SERVER_NAME] = {'command': cmd, 'args': args}
    data['mcpServers'] = servers
    os.makedirs(os.path.dirname(path), exist_ok=True)
    _write_config(path, data)
    return desktop_status(path)


def disconnect_desktop(path=None):
    path = path or desktop_config_path()
    if not path:
        raise ConnectError('Claude 데스크톱 앱이 설치되어 있지 않습니다.')
    data = _read_config(path)
    servers = data.get('mcpServers')
    if isinstance(servers, dict) and SERVER_NAME in servers:
        del servers[SERVER_NAME]
        _write_config(path, data)
    return desktop_status(path)
