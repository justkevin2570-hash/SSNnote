import json
import os
import stat
import sys

import pytest

import claude_connect as cc


# ── 서버 실행 명령 ─────────────────────────────────────────────

def test_server_command_source_mode():
    pytest.importorskip('mcp')
    cmd, args = cc.server_command()
    assert os.path.exists(cmd) and args[0].endswith('mcp_server.py') and os.path.exists(args[0])


def test_server_command_frozen(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    monkeypatch.setattr(sys, 'executable', str(tmp_path / 'SSNnote.exe'))
    with pytest.raises(cc.ConnectError, match='SSNnote_MCP.exe'):
        cc.server_command()
    (tmp_path / 'SSNnote_MCP.exe').write_text('')
    assert cc.server_command() == (str(tmp_path / 'SSNnote_MCP.exe'), [])


# ── Claude 데스크톱 앱 설정 파일 ───────────────────────────────

@pytest.fixture
def desk(tmp_path, monkeypatch):
    monkeypatch.setenv('APPDATA', str(tmp_path / 'Roaming'))
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'Local'))
    monkeypatch.setattr(cc, 'server_command', lambda: ('C:\\SSNnote\\SSNnote_MCP.exe', []))
    return tmp_path


def test_desktop_not_installed(desk):
    assert cc.desktop_config_path() is None
    assert cc.desktop_status()['installed'] is False
    with pytest.raises(cc.ConnectError, match='설치되어 있지 않습니다'):
        cc.connect_desktop()


def test_desktop_connect_keeps_other_settings(desk):
    d = desk / 'Roaming' / 'Claude'
    d.mkdir(parents=True)
    cfg = d / 'claude_desktop_config.json'
    original = {'mcpServers': {'other': {'command': 'npx', 'args': ['x']}}, 'preferences': {'theme': 'dark'}}
    cfg.write_text(json.dumps(original), encoding='utf-8')

    st = cc.connect_desktop()
    assert st['registered'] and st['up_to_date']
    data = json.loads(cfg.read_text(encoding='utf-8'))
    assert data['mcpServers']['other'] == original['mcpServers']['other']
    assert data['preferences'] == {'theme': 'dark'}
    assert data['mcpServers']['ssnnote'] == {'command': 'C:\\SSNnote\\SSNnote_MCP.exe', 'args': []}
    assert json.loads((d / ('claude_desktop_config.json' + cc.BACKUP_SUFFIX)).read_text()) == original

    st = cc.disconnect_desktop()
    assert not st['registered']
    assert json.loads(cfg.read_text(encoding='utf-8')) == original


def test_desktop_connect_creates_file_and_detects_outdated(desk):
    (desk / 'Roaming' / 'Claude').mkdir(parents=True)
    cc.connect_desktop()
    cfg = cc.desktop_config_path()
    data = json.loads(open(cfg, encoding='utf-8').read())
    data['mcpServers']['ssnnote']['command'] = 'D:\\예전\\SSNnote_MCP.exe'
    open(cfg, 'w', encoding='utf-8').write(json.dumps(data))
    st = cc.desktop_status()
    assert st['registered'] and not st['up_to_date']


def test_desktop_invalid_json_is_left_untouched(desk):
    d = desk / 'Roaming' / 'Claude'
    d.mkdir(parents=True)
    cfg = d / 'claude_desktop_config.json'
    cfg.write_text('{ "mcpServers": { broken', encoding='utf-8')
    with pytest.raises(cc.ConnectError, match='형식'):
        cc.connect_desktop()
    assert cfg.read_text(encoding='utf-8') == '{ "mcpServers": { broken'


def test_desktop_store_install_path(desk):
    store = desk / 'Local' / 'Packages' / 'Claude_pzs8sxrjxfjjc' / 'LocalCache' / 'Roaming' / 'Claude'
    store.mkdir(parents=True)
    assert cc.desktop_config_path() == str(store / 'claude_desktop_config.json')


# ── Claude Code (가짜 claude 명령) ─────────────────────────────

FAKE_CLAUDE = r'''#!/usr/bin/env python3
import json, os, sys
db = os.environ['FAKE_CLAUDE_DB']
reg = json.load(open(db)) if os.path.exists(db) else {}
a = sys.argv[1:]
open(db + '.log', 'a').write(' '.join(a) + '\n')
if a[:2] == ['mcp', 'get']:
    if a[2] in reg:
        print(f"{a[2]}:\n  Scope: User config\n  Status: {os.environ.get('FAKE_STATUS', '√ Connected')}\n  Command: {reg[a[2]][0]}")
        sys.exit(0)
    print(f'No MCP server named "{a[2]}".'); sys.exit(1)
if a[:2] == ['mcp', 'add']:
    i = a.index('--'); name = a[i - 1]; reg[name] = a[i + 1:]
    json.dump(reg, open(db, 'w')); print('Added'); sys.exit(0)
if a[:2] == ['mcp', 'remove']:
    name = a[-1]
    if name in reg:
        del reg[name]; json.dump(reg, open(db, 'w')); print('Removed'); sys.exit(0)
    print(f'No MCP server named "{name}" in user scope'); sys.exit(1)   # 실제 claude 문구
'''


@pytest.fixture
def fake_claude(tmp_path, monkeypatch):
    if sys.platform == 'win32':
        pytest.skip('가짜 claude 스크립트는 Linux/macOS에서만')
    exe = tmp_path / 'claude'
    exe.write_text(FAKE_CLAUDE)
    exe.chmod(exe.stat().st_mode | stat.S_IEXEC)
    monkeypatch.setenv('FAKE_CLAUDE_DB', str(tmp_path / 'reg.json'))
    monkeypatch.setenv('PATH', str(tmp_path) + os.pathsep + os.environ['PATH'])
    monkeypatch.setattr(cc, 'server_command', lambda: ('/opt/ssnnote/SSNnote_MCP', []))
    return tmp_path


def test_claude_code_connect_status_disconnect(fake_claude):
    assert cc.find_claude_code() == str(fake_claude / 'claude')
    assert cc.claude_code_status() == {'installed': True, 'registered': False, 'connected': False,
                                       'detail': 'No MCP server named "ssnnote".'}
    st = cc.connect_claude_code()
    assert st['registered'] and st['connected']
    log = (fake_claude / 'reg.json.log').read_text().splitlines()
    assert 'mcp add --scope user ssnnote -- /opt/ssnnote/SSNnote_MCP' in log
    st = cc.connect_claude_code()                       # 다시 연결해도 중복 없이 교체
    assert json.loads((fake_claude / 'reg.json').read_text()) == {'ssnnote': ['/opt/ssnnote/SSNnote_MCP']}
    st = cc.disconnect_claude_code()
    assert not st['registered']
    cc.disconnect_claude_code()                         # 이미 없어도 오류 없음


def test_claude_code_registered_but_failing(fake_claude, monkeypatch):
    cc.connect_claude_code()
    monkeypatch.setenv('FAKE_STATUS', '✗ Failed to connect')
    st = cc.claude_code_status()
    assert st['registered'] and not st['connected']


def test_claude_code_not_installed(monkeypatch, tmp_path):
    monkeypatch.setenv('PATH', str(tmp_path))
    monkeypatch.setattr(os.path, 'expanduser', lambda p: str(tmp_path))
    monkeypatch.setenv('APPDATA', str(tmp_path))
    assert cc.find_claude_code() is None
    assert cc.claude_code_status()['installed'] is False
