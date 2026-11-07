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

def desktop_config_dirs():
    """Claude 데스크톱 앱 설정 폴더 후보 중 실제로 있는 것.

    - 요즘 스토어판: %LOCALAPPDATA%\\Claude-Data\\claude_<게시자ID>   (설정 → 개발자 → 구성 편집이 여는 곳)
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
    """Claude 앱이 읽을 수 있는 설정 파일 경로 전부. 스토어판은 일반 위치와 앱 전용 위치를 함께 볼 수 있어서
    (앱 전용 위치에 파일이 있으면 그쪽을 읽음) 한쪽에만 쓰면 등록이 안 보일 수 있다 → 있는 곳 모두에 쓴다."""
    return [os.path.join(d, 'claude_desktop_config.json') for d in desktop_config_dirs()]


def desktop_config_path():
    """Claude 데스크톱 앱 설정 파일 경로 (앱이 없으면 None).

    일반 설치는 %APPDATA%\\Claude, Microsoft Store 설치는 %LOCALAPPDATA%\\Packages\\Claude_*\\...\\Claude.
    """
    paths = desktop_config_paths()
    with_file = [p for p in paths if os.path.exists(p)]
    chosen = (with_file or paths or [None])[0]
    return chosen


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


def _entry_state(path):
    """(등록됨, 지금 위치와 같음, 등록된 명령)."""
    try:
        entry = (_read_config(path).get('mcpServers') or {}).get(SERVER_NAME)
    except ConnectError:
        entry = None
    if not isinstance(entry, dict):
        return False, False, ''
    try:
        cmd, args = server_command()
        current = entry.get('command') == cmd and entry.get('args', []) == args
    except ConnectError:
        current = False
    return True, current, entry.get('command', '')


def desktop_status(path=None):
    paths = [path] if path else desktop_config_paths()
    if not paths:
        return {'installed': False, 'registered': False, 'up_to_date': False, 'path': ''}
    states = [_entry_state(p) for p in paths]
    registered = [st for st in states if st[0]]
    return {'installed': True, 'registered': bool(registered),
            'up_to_date': bool(registered) and all(st[1] for st in registered),
            'path': path or desktop_config_path()}


def connect_desktop(path=None):
    paths = [path] if path else desktop_config_paths()
    if not paths:
        raise ConnectError('Claude 데스크톱 앱이 설치되어 있지 않습니다.')
    cmd, args = server_command()
    for p in paths:
        data = _read_config(p)
        servers = data.get('mcpServers')
        if not isinstance(servers, dict):
            servers = {}
        servers[SERVER_NAME] = {'command': cmd, 'args': args}
        data['mcpServers'] = servers
        os.makedirs(os.path.dirname(p), exist_ok=True)
        _write_config(p, data)
    return desktop_status(path)


def disconnect_desktop(path=None):
    paths = [path] if path else desktop_config_paths()
    if not paths:
        raise ConnectError('Claude 데스크톱 앱이 설치되어 있지 않습니다.')
    for p in paths:
        data = _read_config(p)
        servers = data.get('mcpServers')
        if isinstance(servers, dict) and SERVER_NAME in servers:
            del servers[SERVER_NAME]
            _write_config(p, data)
    return desktop_status(path)


def refresh_registrations():
    """앱을 켤 때: 사용자가 Claude 데스크톱 연결을 해 둔 경우(어느 한 설정 파일에라도 등록됨),
    Claude 앱이 읽을 수 있는 모든 설정 위치에 지금 위치로 등록돼 있게 맞춘다 (설치 위치가 바뀌었거나,
    Claude 앱이 새 설정 위치를 쓰기 시작한 경우). 연결한 적 없으면 건드리지 않는다.
    Claude Code는 확인이 몇 초 걸려서 설치 위치가 바뀐 뒤 처음 켤 때 한 번만."""
    if not getattr(sys, 'frozen', False):
        return
    try:
        cmd, _ = server_command()
    except ConnectError:
        return
    paths = desktop_config_paths()
    states = {p: _entry_state(p) for p in paths}
    if any(reg for reg, _, _ in states.values()) and not all(reg and cur for reg, cur, _ in states.values()):
        try:
            connect_desktop()
        except ConnectError as e:
            print(f'[claude] 데스크톱 연결 맞추기 실패: {e}')
    import db
    if db.get_setting('claude_paths_for', '') == cmd:
        return
    claude = find_claude_code()
    if claude:
        code, out = _run([claude, 'mcp', 'get', SERVER_NAME])
        if code == 0 and 'no mcp server' not in out.lower() and cmd.lower() not in out.lower():
            connect_claude_code(claude)
    db.set_setting('claude_paths_for', cmd)


# ── 연결 진단 ─────────────────────────────────────────────────

def self_test(timeout=25.0):
    """연결 프로그램을 Claude 앱처럼 직접 실행해 MCP 인사(initialize)와 도구 목록(tools/list)을 주고받는다.
    (ok, 설명) — mcp 패키지 없이 줄 단위 JSON-RPC로 직접 말한다."""
    import threading
    import time
    try:
        cmd, args = server_command()
    except ConnectError as e:
        return False, str(e)
    env = {**os.environ, 'SSNNOTE_MCP_SELFTEST': '1'}
    kw = {'creationflags': 0x08000000} if sys.platform == 'win32' else {}
    t0 = time.monotonic()
    try:
        p = subprocess.Popen([cmd, *args], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                             cwd=os.path.expanduser('~'), env=env, **kw)
    except OSError as e:
        return False, f'실행하지 못함: {e}'
    lines, errs = [], []

    def reader(stream, out):
        for raw in iter(stream.readline, b''):
            out.append(raw.decode('utf-8', 'replace').strip())
    threading.Thread(target=reader, args=(p.stdout, lines), daemon=True).start()
    threading.Thread(target=reader, args=(p.stderr, errs), daemon=True).start()

    def send(obj):
        p.stdin.write((json.dumps(obj) + '\n').encode('utf-8'))
        p.stdin.flush()

    def wait_for(msg_id):
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            for ln in list(lines):
                try:
                    m = json.loads(ln)
                except ValueError:
                    continue
                if isinstance(m, dict) and m.get('id') == msg_id:
                    return m
            if p.poll() is not None and not any(ln.startswith('{') for ln in lines):
                break
            time.sleep(0.05)
        return None
    try:
        send({'jsonrpc': '2.0', 'id': 1, 'method': 'initialize',
              'params': {'protocolVersion': '2025-06-18', 'capabilities': {},
                         'clientInfo': {'name': 'ssnnote-selftest', 'version': '1'}}})
        init = wait_for(1)
        if not init or 'result' not in init:
            tail = ' / '.join((errs or lines)[-5:])[-400:]
            return False, '연결 프로그램이 응답하지 않음' + (f' ({tail})' if tail else '')
        send({'jsonrpc': '2.0', 'method': 'notifications/initialized'})
        send({'jsonrpc': '2.0', 'id': 2, 'method': 'tools/list'})
        tl = wait_for(2)
        tools = [t.get('name') for t in ((tl or {}).get('result') or {}).get('tools', [])]
        if not tools:
            return False, '도구 목록을 받지 못함'
        return True, f'정상 — 도구 {len(tools)}개, {time.monotonic() - t0:.1f}초 ({", ".join(tools)})'
    except OSError as e:
        return False, f'주고받기 실패: {e}'
    finally:
        try:
            p.kill()
        except OSError:
            pass


def _tail_lines(path, n=20, needle=None):
    try:
        with open(path, encoding='utf-8', errors='replace') as f:
            lines = f.read()[-200_000:].splitlines()
    except OSError:
        return []
    if needle:
        lines = [ln for ln in lines if needle in ln.lower()]
    return [ln[:300] for ln in lines[-n:]]


def desktop_logs():
    """Claude 앱이 남긴 연결 기록: [(경로, 수정 시각, 마지막 줄들)]."""
    from datetime import datetime
    out = []
    for d in desktop_config_dirs():
        for name, needle in (('mcp-server-ssnnote.log', None), ('mcp.log', SERVER_NAME)):
            path = os.path.join(d, 'logs', name)
            if os.path.exists(path):
                when = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M')
                out.append((path, when, _tail_lines(path, 20, needle)))
    return out


def diagnose() -> str:
    """사용자가 복사해 보낼 수 있는 연결 진단 글."""
    from datetime import datetime
    lines = [f'[서서니 노트 Claude 연결 진단] {datetime.now():%Y-%m-%d %H:%M}']
    try:
        from updater import APP_VERSION
    except Exception:
        APP_VERSION = '?'
    lines.append(f"버전: {APP_VERSION} ({'설치판' if getattr(sys, 'frozen', False) else '소스 실행'})")
    try:
        cmd, args = server_command()
        lines.append(f"연결 프로그램: {' '.join([cmd, *args])}")
    except ConnectError as e:
        lines.append(f'연결 프로그램: 문제 — {e}')
    ok, msg = self_test()
    lines.append(f"직접 실행 시험: {'✓ ' if ok else '✗ '}{msg}")
    try:
        import db
        lines.append(f'DB: {db.DB_PATH}')
        lines.append(f"Claude 앱이 연결 프로그램을 마지막으로 켠 시각: {db.get_setting('mcp_last_start', '') or '기록 없음'}")
        lines.append(f"Claude가 마지막으로 업무를 요청한 시각: {db.get_setting('mcp_last_call', '') or '기록 없음'}")
    except Exception as e:
        lines.append(f'DB 확인 실패: {e}')

    lines.append('')
    lines.append('Claude 데스크톱 앱 설정 파일:')
    paths = desktop_config_paths()
    if not paths:
        lines.append('  (Claude 데스크톱 앱 폴더를 찾지 못함 — 설치되어 있지 않거나 한 번도 실행하지 않음)')
    for p in paths:
        kind = ('스토어 설치' if os.sep + 'Claude-Data' + os.sep in p
                else '예전 스토어 설치 위치' if os.sep + 'Packages' + os.sep in p else '일반 설치')
        if not os.path.exists(p):
            lines.append(f'  - {p} ({kind}): 파일 없음')
            continue
        reg, cur, command = _entry_state(p)
        state = ('등록됨' + ('' if cur else f' (예전 위치: {command})')) if reg else '서서니 노트 등록 안 됨'
        try:
            others = sorted(k for k in (_read_config(p).get('mcpServers') or {}) if k != SERVER_NAME)
        except ConnectError as e:
            others, state = [], f'읽기 실패 — {e}'
        lines.append(f'  - {p} ({kind}): {state}' + (f" · 다른 연결: {', '.join(others)}" if others else ''))

    lines.append('')
    lines.append('Claude 앱 연결 기록(로그):')
    logs = desktop_logs()
    if not logs:
        lines.append('  (서서니 노트 관련 기록 없음 — Claude 앱이 아직 연결 프로그램을 실행하지 않음)')
    for path, when, tail in logs:
        lines.append(f'  - {path} (수정 {when})')
        lines += [f'    {ln}' for ln in tail] or ['    (서서니 노트 관련 줄 없음)']
    return '\n'.join(lines)
