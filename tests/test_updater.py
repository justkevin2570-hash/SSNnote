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
    assert updater.pick_newest(None, gist) is gist           # 릴리스 확인 실패해도 Gist로
    assert updater.pick_newest(rel, None) is rel             # Gist 실패해도 릴리스로
    assert updater.pick_newest({'version': 'v2.10', 'download_url': 'a'},
                               {'version': 'v2.9', 'download_url': 'b'})['version'] == 'v2.10'
    assert updater.pick_newest(None, None, {'version': 'v3.0'}) is None   # 다운로드 주소 없으면 제외
