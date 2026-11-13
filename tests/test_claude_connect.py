"""예전 「Claude 연결」이 남긴 ssnnote 커넥터 등록을 지우는지."""
import json
import subprocess

import claude_connect as cc
import db


def _cfg(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding='utf-8')
    return path


def test_removes_ssnnote_from_every_desktop_config(tmp_path, monkeypatch):
    local, appdata = tmp_path / 'local', tmp_path / 'roaming'
    monkeypatch.setenv('LOCALAPPDATA', str(local))
    monkeypatch.setenv('APPDATA', str(appdata))
    store = _cfg(local / 'Claude-Data' / 'claude_abc' / 'claude_desktop_config.json',
                 {'mcpServers': {'ssnnote': {'command': 'C:\\x\\SSNnote_MCP.exe'}, 'other': {'command': 'npx'}},
                  'preferences': {'theme': 'dark'}})
    normal = _cfg(appdata / 'Claude' / 'claude_desktop_config.json',
                  {'mcpServers': {'ssnnote': {'command': 'C:\\x\\SSNnote_MCP.exe'}}})
    removed = cc.remove_desktop_entries()
    assert sorted(removed) == sorted([str(store), str(normal)])
    a = json.loads(store.read_text(encoding='utf-8'))
    assert a == {'mcpServers': {'other': {'command': 'npx'}}, 'preferences': {'theme': 'dark'}}   # 다른 설정은 그대로
    assert json.loads(normal.read_text(encoding='utf-8')) == {'mcpServers': {}}
    backup = json.loads(open(str(store) + cc.BACKUP_SUFFIX, encoding='utf-8').read())
    assert 'ssnnote' in backup['mcpServers']                                     # 원본은 백업
    assert cc.remove_desktop_entries() == []                                     # 두 번째는 아무것도 안 바꿈


def test_leaves_unrelated_or_broken_configs_alone(tmp_path):
    clean = _cfg(tmp_path / 'a' / 'claude_desktop_config.json', {'mcpServers': {'other': {}}})
    broken = tmp_path / 'b' / 'claude_desktop_config.json'
    broken.parent.mkdir()
    broken.write_text('{ "mcpServers": { "ssnnote": ', encoding='utf-8')
    before = clean.stat().st_mtime_ns
    assert cc.remove_desktop_entries([str(clean), str(broken), str(tmp_path / 'none.json')]) == []
    assert clean.stat().st_mtime_ns == before and not (tmp_path / 'a' / ('claude_desktop_config.json' + cc.BACKUP_SUFFIX)).exists()
    assert broken.read_text(encoding='utf-8').startswith('{ "mcpServers"')


def test_claude_code_entry_removed_once(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setattr(cc, 'desktop_config_paths', lambda: [])
    monkeypatch.setattr(cc, 'find_claude_code', lambda: 'claude')
    calls = []

    def fake_run(cmd, **kw):
        calls.append(cmd)
        return subprocess.CompletedProcess(cmd, 1, '', 'No MCP server found with name: ssnnote')
    monkeypatch.setattr(cc.subprocess, 'run', fake_run)
    cc.cleanup_old_connector()
    cc.cleanup_old_connector()
    assert calls == [['claude', 'mcp', 'remove', '--scope', 'user', 'ssnnote']]


def test_claude_code_retried_when_it_fails(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setattr(cc, 'desktop_config_paths', lambda: [])
    monkeypatch.setattr(cc, 'find_claude_code', lambda: 'claude')

    def boom(cmd, **kw):
        raise subprocess.TimeoutExpired(cmd, 60)
    monkeypatch.setattr(cc.subprocess, 'run', boom)
    cc.cleanup_old_connector()
    assert db.get_setting(cc.CODE_CLEANUP_KEY, '') != '1'


def test_no_claude_code_installed_is_fine(monkeypatch):
    monkeypatch.setattr(cc, 'find_claude_code', lambda: None)
    assert cc.remove_claude_code_entry() is True
