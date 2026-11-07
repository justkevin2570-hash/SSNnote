import pytest

pytest.importorskip('PyQt5')
import updater  # noqa: E402

REL = {
    'tag_name': 'v2.69', 'draft': False, 'prerelease': False, 'body': '- 변경 1\n- 변경 2\n',
    'assets': [
        {'name': 'other.zip', 'browser_download_url': 'https://x/other.zip'},
        {'name': 'SSNnote_Setup.exe', 'browser_download_url': 'https://x/v2.69/SSNnote_Setup.exe',
         'digest': 'sha256:2f8325a9df1e72c53e2bfdd286f9de2b37590b08dbf95886b7ed0df347e141fe'},
    ],
}


def test_release_to_version_info():
    v = updater.release_to_version_info(REL)
    assert v == {'version': 'v2.69', 'download_url': 'https://x/v2.69/SSNnote_Setup.exe',
                 'sha256': '2F8325A9DF1E72C53E2BFDD286F9DE2B37590B08DBF95886B7ED0DF347E141FE',
                 'changelog': '- 변경 1\n- 변경 2'}


@pytest.mark.parametrize('change', [
    {'draft': True}, {'prerelease': True}, {'assets': []}, {'tag_name': ''},
])
def test_release_ignored_when_not_usable(change):
    assert updater.release_to_version_info({**REL, **change}) is None


def test_release_without_digest_still_usable():
    rel = {**REL, 'assets': [{'name': 'SSNnote_Setup.exe', 'browser_download_url': 'https://x/s.exe'}]}
    assert updater.release_to_version_info(rel)['sha256'] == ''


def test_pick_newest_prefers_higher_version_from_either_source():
    gist = {'version': 'v2.66', 'download_url': 'g'}
    rel = {'version': 'v2.68', 'download_url': 'r'}
    assert updater.pick_newest(rel, gist) is rel
    assert updater.pick_newest(None, gist) is gist           # 릴리스 확인 실패해도 다른 출처로
    assert updater.pick_newest(rel, None) is rel
    assert updater.pick_newest({'version': 'v2.10', 'download_url': 'a'},
                               {'version': 'v2.9', 'download_url': 'b'})['version'] == 'v2.10'
    assert updater.pick_newest(None, None, {'version': 'v3.0'}) is None   # 다운로드 주소 없으면 제외


# ── 자동 업데이트 ─────────────────────────────────────────────

import hashlib  # noqa: E402
import json  # noqa: E402
import sys  # noqa: E402

GOOD_URL = 'https://github.com/justkevin2570-hash/SSNnote/releases/download/v9.1/SSNnote_Setup.exe'


@pytest.mark.parametrize('url,ok', [
    (GOOD_URL, True),
    ('https://github.com/evil/SSNnote/releases/download/v9.1/SSNnote_Setup.exe', False),
    ('https://evil.example/justkevin2570-hash/SSNnote/releases/download/v9.1/SSNnote_Setup.exe', False),
    ('http://github.com/justkevin2570-hash/SSNnote/releases/download/v9.1/SSNnote_Setup.exe', False),
    ('https://github.com/justkevin2570-hash/SSNnote/releases/download/v9.1/other.exe', False),
    ('https://github.com/justkevin2570-hash/SSNnote/releases/download/../../x/SSNnote_Setup.exe', False),
    ('https://github.com/justkevin2570-hash/SSNnote/releases/download/v9.1/SSNnote_Setup.exe?x=1', False),
    ('', False),
])
def test_only_this_repos_release_is_trusted(url, ok):
    assert updater.is_trusted({'download_url': url}) is ok


def test_fetch_version_info_ignores_untrusted_sources(monkeypatch):
    evil = {'version': 'v99.0', 'download_url': 'https://evil.example/SSNnote_Setup.exe', 'sha256': 'AA'}
    good = {'version': 'v9.1', 'download_url': GOOD_URL, 'sha256': 'BB'}
    monkeypatch.setattr(updater, 'release_to_version_info', lambda r: r)
    monkeypatch.setattr(updater, '_get_json', lambda url, accept='': evil if 'raw.' in url else good)
    assert updater.fetch_version_info() is good


@pytest.fixture
def upd(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    d = tmp_path / 'update'
    monkeypatch.setattr(updater, '_UPDATE_DIR', str(d))
    monkeypatch.setattr(updater, '_READY_FILE', str(d / 'ready.json'))
    monkeypatch.setattr(updater, 'APP_VERSION', 'v9.0')
    monkeypatch.setattr(updater, '_launched', False)
    payload = b'installer-bytes'
    sha = hashlib.sha256(payload).hexdigest().upper()

    def fake_download(info, dest, progress=None):
        assert updater.is_trusted(info) and info['sha256']
        import os
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        open(dest, 'wb').write(payload)
        return dest
    monkeypatch.setattr(updater, 'download', fake_download)
    launched = []
    monkeypatch.setattr(updater.subprocess, 'Popen', lambda cmd, **kw: launched.append(cmd))
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    info = {'version': 'v9.1', 'download_url': GOOD_URL, 'sha256': sha, 'changelog': '- 새 기능'}
    return info, launched, d


def test_stage_then_install_on_startup_relaunches(upd):
    info, launched, d = upd
    data = updater.stage(info)
    assert data['version'] == 'v9.1' and (d / 'ready.json').exists()
    assert updater.stage(info)['path'] == data['path']                   # 같은 버전은 다시 받지 않음
    assert updater.install_on_startup() is True
    cmd = launched[-1]
    assert cmd[0] == data['path'] and '/VERYSILENT' in cmd and '/NORELAUNCH' not in cmd
    assert json.loads((d / 'ready.json').read_text())['attempts'] == 1
    assert updater.install_on_quit() is False                           # 같은 실행에서 두 번 띄우지 않음


def test_install_on_quit_does_not_relaunch(upd):
    info, launched, _ = upd
    updater.stage(info)
    assert updater.install_on_quit() is True and launched[-1][-1] == '/NORELAUNCH'


def test_tampered_file_is_not_installed(upd):
    info, launched, d = upd
    data = updater.stage(info)
    open(data['path'], 'wb').write(b'tampered')
    assert updater.install_on_startup() is False and not launched
    assert not (d / 'ready.json').exists()                              # 받아 둔 것은 지움


def test_gives_up_after_two_failed_attempts(upd, monkeypatch):
    info, launched, _ = upd
    updater.stage(info)
    for _ in range(updater.MAX_ATTEMPTS):
        monkeypatch.setattr(updater, '_launched', False)               # 다음 실행
        assert updater.install_on_startup() is True
    monkeypatch.setattr(updater, '_launched', False)
    assert updater.install_on_startup() is False and len(launched) == updater.MAX_ATTEMPTS


def test_installed_version_clears_staging(upd, monkeypatch):
    info, _, d = upd
    updater.stage(info)
    monkeypatch.setattr(updater, 'APP_VERSION', 'v9.1')                  # 설치가 끝나 새 판으로 켜짐
    assert updater.ready_update() is None and not any(d.iterdir())


def test_auto_update_off_does_nothing(upd):
    info, launched, _ = upd
    updater.stage(info)
    updater.set_auto_enabled(False)
    assert updater.install_on_startup() is False and updater.install_on_quit() is False and not launched


def test_just_updated(upd, monkeypatch):
    assert updater.just_updated() == ''                                 # 처음 설치
    monkeypatch.setattr(updater, 'APP_VERSION', 'v9.1')
    assert updater.just_updated() == 'v9.0'
    assert updater.just_updated() == ''


def test_download_refuses_untrusted_or_unsigned(tmp_path):
    with pytest.raises(RuntimeError, match='믿을 수 없는'):
        updater.download({'download_url': 'https://evil.example/SSNnote_Setup.exe', 'sha256': 'AA'}, str(tmp_path / 'x'))
    with pytest.raises(RuntimeError, match='SHA256'):
        updater.download({'download_url': GOOD_URL, 'sha256': ''}, str(tmp_path / 'x'))
