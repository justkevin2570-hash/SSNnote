"""
자동 업데이트

흐름 (자동 업데이트 켬 — 기본)
  1) 앱을 켜면 뒤에서 최신 버전 확인 (GitHub 최신 릴리스, 안 되면 저장소의 version.json)
  2) 새 버전이 있으면 조용히 내려받아 SHA256으로 확인하고 업데이트 폴더에 둔다 (ready.json)
  3) 앱을 끌 때(트레이 → 종료) 또는 다음에 켤 때 화면 없이 설치 → 다음에 켤 때 설치한 경우엔 자동으로 다시 켜짐
  4) 업데이트 뒤 처음 켜면 제목 표시줄에 「vX.YY로 업데이트됐어요」

안전장치
  - 설치 파일은 이 저장소의 GitHub 릴리스 주소에서 받은 것만 쓴다 (버전 정보가 다른 곳을 가리키면 거부)
  - SHA256이 없거나 맞지 않으면 설치하지 않는다. 설치 직전에도 다시 확인한다.
  - 두 번 연속 설치에 실패하면 자동 설치를 멈추고 한 번만 알린다.

자동 업데이트를 끄면(메뉴 → 설정 → 자동 업데이트) 예전처럼 새 버전 알림만 한 번 띄운다.
메뉴 → 도움말 → 업데이트 확인: [지금 업데이트] 한 번이면 내려받아 설치하고 다시 켜진다.
"""

import hashlib
import json
import os
import ssl
import subprocess
import sys
import threading
import urllib.error
import urllib.request

from PyQt5.QtCore import QObject, Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressDialog


def _read_app_version() -> str:
    """번들된 version.json에서 현재 버전을 읽는다. 실패 시 'v0.0' 반환."""
    try:
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base, 'version.json')
        with open(path, encoding='utf-8') as f:
            return json.load(f).get('version', 'v0.0')
    except Exception:
        return 'v0.0'


APP_VERSION = _read_app_version()

_REPO = 'justkevin2570-hash/SSNnote'
# 최신 GitHub 릴리스 — 릴리스를 올리면 바로 업데이트
_LATEST_RELEASE_URL = f'https://api.github.com/repos/{_REPO}/releases/latest'
# 예비: 저장소의 version.json (학교망에서 GitHub API 호출 한도에 걸렸을 때)
_VERSION_JSON_URL = f'https://raw.githubusercontent.com/{_REPO}/master/version.json'
# 설치 파일은 이 주소로 시작하는 것만 받는다
_TRUSTED_DOWNLOAD_PREFIX = f'https://github.com/{_REPO}/releases/download/'
_SETUP_ASSET_NAME = 'SSNnote_Setup.exe'

_APPDATA_DIR = os.path.join(os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming'),
                            'SSNnote')
_NOTIFIED_FILE = os.path.join(_APPDATA_DIR, 'last_notified_version.txt')
_UPDATE_DIR = os.path.join(os.environ.get('LOCALAPPDATA') or _APPDATA_DIR, 'SSNnote', 'update')
_READY_FILE = os.path.join(_UPDATE_DIR, 'ready.json')
_REQUEST_TIMEOUT = 10
MAX_ATTEMPTS = 2                 # 같은 버전 설치를 이만큼 실패하면 자동 설치를 멈춤

AUTO_KEY = 'auto_update'         # app_settings: '1'(기본) / '0'
LAST_VERSION_KEY = 'last_run_version'

# 화면 없이 설치: 실행 중인 서서니 노트·연결 프로그램은 설치 프로그램이 닫는다
SILENT_ARGS = ['/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CLOSEAPPLICATIONS', '/FORCECLOSEAPPLICATIONS']
NO_RELAUNCH_ARG = '/NORELAUNCH'  # 설치 뒤 앱을 다시 켜지 않음 (사용자가 끈 경우)


# ── 버전 정보 ─────────────────────────────────────────────────

def _get_json(url: str, accept: str = 'application/json') -> dict | None:
    try:
        req = urllib.request.Request(url)
        req.add_header('Accept', accept)
        req.add_header('User-Agent', 'SSNnote')
        req.add_header('Cache-Control', 'no-cache')
        with urllib.request.urlopen(req, timeout=_REQUEST_TIMEOUT, context=ssl.create_default_context()) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except Exception:
        return None


def release_to_version_info(release: dict) -> dict | None:
    """GitHub 릴리스 API 응답 → version.json과 같은 모양 {version, download_url, sha256, changelog}.

    설치 파일(SSNnote_Setup.exe)이 없거나 초안·시험판이면 None.
    sha256은 GitHub가 계산해 주는 파일 digest('sha256:...')를 쓴다.
    """
    if not isinstance(release, dict) or release.get('draft') or release.get('prerelease'):
        return None
    tag = release.get('tag_name') or ''
    asset = next((a for a in release.get('assets') or [] if a.get('name') == _SETUP_ASSET_NAME), None)
    if not tag or not asset or not asset.get('browser_download_url'):
        return None
    digest = asset.get('digest') or ''
    sha = digest.split(':', 1)[1].upper() if digest.lower().startswith('sha256:') else ''
    return {
        'version': tag,
        'download_url': asset['browser_download_url'],
        'sha256': sha,
        'changelog': (release.get('body') or '').strip(),
    }


def is_trusted(info) -> bool:
    """이 저장소의 릴리스 설치 파일인지 (버전 정보가 다른 곳을 가리키면 받지 않는다)."""
    url = (info or {}).get('download_url') or ''
    return url.startswith(_TRUSTED_DOWNLOAD_PREFIX) and url.endswith('/' + _SETUP_ASSET_NAME) \
        and '..' not in url and '?' not in url and '#' not in url


def pick_newest(*infos) -> dict | None:
    """여러 출처 중 가장 높은 버전. 모두 실패면 None."""
    best = None
    for info in infos:
        if not info or not info.get('version') or not info.get('download_url'):
            continue
        if best is None or is_newer_version(info['version'], best['version']):
            best = info
    return best


def fetch_version_info() -> dict | None:
    """최신 버전 정보 (믿을 수 있는 주소만). 실패 시 None."""
    infos = [release_to_version_info(_get_json(_LATEST_RELEASE_URL, 'application/vnd.github+json')),
             _get_json(_VERSION_JSON_URL)]
    data = pick_newest(*[i for i in infos if isinstance(i, dict) and is_trusted(i)])
    if data is None:
        print('[UPDATE] 버전 정보를 가져오지 못함')
    return data


def is_newer_version(remote_tag: str, local_tag: str) -> bool:
    """버전 비교 (정수 튜플 — v2.10 > v2.9). 파싱 실패 시 False."""
    try:
        r_parts = [int(x) for x in remote_tag.lstrip('v').split('.')]
        l_parts = [int(x) for x in local_tag.lstrip('v').split('.')]
        return tuple(r_parts) > tuple(l_parts)
    except Exception:
        return False


# ── 설정 ─────────────────────────────────────────────────────

def auto_enabled() -> bool:
    try:
        from db import get_setting
        return get_setting(AUTO_KEY, '1') != '0'
    except Exception:
        return True


def set_auto_enabled(on: bool):
    from db import set_setting
    set_setting(AUTO_KEY, '1' if on else '0')


def _get_last_notified_version() -> str:
    try:
        with open(_NOTIFIED_FILE, encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return ''


def _save_notified_version(tag: str):
    try:
        os.makedirs(_APPDATA_DIR, exist_ok=True)
        with open(_NOTIFIED_FILE, 'w', encoding='utf-8') as f:
            f.write(tag)
    except Exception:
        pass


# ── 내려받기 · 준비 ───────────────────────────────────────────

def _sha256_of(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest().upper()


def download(info: dict, dest: str, progress=None) -> str:
    """설치 파일을 받아 SHA256을 확인한다. 실패하면 RuntimeError (받은 파일은 지움).
    progress(받은 바이트, 전체 바이트) — False를 돌려주면 취소."""
    if not is_trusted(info):
        raise RuntimeError('믿을 수 없는 다운로드 주소라 받지 않았습니다.')
    sha = (info.get('sha256') or '').upper()
    if not sha:
        raise RuntimeError('설치 파일의 확인값(SHA256)이 없어 받지 않았습니다.')
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = dest + '.part'
    try:
        import requests
        with requests.get(info['download_url'], stream=True, timeout=(10, 120)) as r:
            r.raise_for_status()
            total = int(r.headers.get('Content-Length') or 0)
            done = 0
            with open(tmp, 'wb') as f:
                for chunk in r.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        done += len(chunk)
                        if progress and progress(done, total) is False:
                            raise RuntimeError('취소했습니다.')
        if _sha256_of(tmp) != sha:
            raise RuntimeError('내려받은 설치 파일의 확인값이 맞지 않습니다 (손상 또는 변조).')
        os.replace(tmp, dest)
        return dest
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _write_ready(data: dict):
    os.makedirs(_UPDATE_DIR, exist_ok=True)
    with open(_READY_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)


def clear_staged():
    """업데이트 폴더 비우기."""
    try:
        for name in os.listdir(_UPDATE_DIR):
            try:
                os.remove(os.path.join(_UPDATE_DIR, name))
            except OSError:
                pass
    except OSError:
        pass


def stage(info: dict) -> dict:
    """새 버전을 미리 내려받아 둔다 (이미 받아 둔 같은 버전이면 그대로). ready.json 내용을 돌려준다."""
    cur = ready_update()
    if cur and cur['version'] == info['version']:
        return cur
    clear_staged()
    path = os.path.join(_UPDATE_DIR, f"SSNnote_Setup-{info['version']}.exe")
    download(info, path)
    data = {'version': info['version'], 'path': path, 'sha256': info['sha256'].upper(),
            'download_url': info['download_url'], 'changelog': info.get('changelog', ''), 'attempts': 0}
    _write_ready(data)
    return data


def ready_update() -> dict | None:
    """설치할 준비가 된 업데이트 (지금보다 새 버전이고 파일·확인값이 맞을 때만). 아니면 None (필요하면 정리)."""
    try:
        with open(_READY_FILE, encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or not is_newer_version(data.get('version', ''), APP_VERSION):
        clear_staged()               # 이미 설치됐거나 잘못된 기록
        return None
    path = data.get('path') or ''
    if not is_trusted(data) or os.path.dirname(os.path.abspath(path)) != os.path.abspath(_UPDATE_DIR) \
            or not os.path.exists(path):
        clear_staged()
        return None
    return data


def verify_ready(data: dict) -> bool:
    """설치 직전 다시 확인 (받아 둔 뒤 파일이 바뀌지 않았는지)."""
    try:
        return bool(data.get('sha256')) and _sha256_of(data['path']) == data['sha256'].upper()
    except OSError:
        return False


def installer_command(data: dict, relaunch: bool) -> list:
    return [data['path'], *SILENT_ARGS, *([] if relaunch else [NO_RELAUNCH_ARG])]


_launched = False                # 이번 실행에서 이미 설치 프로그램을 띄웠는지 (종료할 때 두 번 띄우지 않게)


def launch_installer(data: dict, relaunch: bool) -> bool:
    """받아 둔 설치 파일을 화면 없이 실행한다 (이 프로그램이 끝나기를 설치 프로그램이 기다림).
    성공적으로 실행했으면 True — 호출한 쪽은 곧바로 앱을 끝내야 한다."""
    if not verify_ready(data):
        clear_staged()
        return False
    data = {**data, 'attempts': int(data.get('attempts', 0)) + 1}
    _write_ready(data)
    kw = {}
    if sys.platform == 'win32':
        kw['creationflags'] = 0x00000008 | 0x00000200      # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    try:
        subprocess.Popen(installer_command(data, relaunch), close_fds=True, cwd=_UPDATE_DIR, **kw)
    except OSError as e:
        print('[UPDATE] 설치 프로그램 실행 실패:', e)
        return False
    global _launched
    _launched = True
    return True


def gave_up(data) -> bool:
    return bool(data) and int(data.get('attempts', 0)) >= MAX_ATTEMPTS


# ── 앱 시작 · 종료 때 ─────────────────────────────────────────

def install_on_startup() -> bool:
    """앱을 켤 때 받아 둔 업데이트가 있으면 설치하고 True (호출한 쪽은 바로 끝내야 함).
    설치하면 앱이 자동으로 다시 켜진다."""
    if not getattr(sys, 'frozen', False) or not auto_enabled():
        return False
    data = ready_update()
    if not data or gave_up(data):
        return False
    return launch_installer(data, relaunch=True)


def install_on_quit() -> bool:
    """트레이 → 종료로 앱을 끌 때: 받아 둔 업데이트를 설치한다 (다시 켜지 않음)."""
    if _launched or not getattr(sys, 'frozen', False) or not auto_enabled():
        return False
    data = ready_update()
    if not data or gave_up(data):
        return False
    return launch_installer(data, relaunch=False)


def just_updated() -> str:
    """업데이트 뒤 처음 켰으면 이전 버전 문자열, 아니면 ''. (지금 버전을 기록한다)"""
    try:
        from db import get_setting, set_setting
        prev = get_setting(LAST_VERSION_KEY, '')
        if prev != APP_VERSION:
            set_setting(LAST_VERSION_KEY, APP_VERSION)
        return prev if prev and is_newer_version(APP_VERSION, prev) else ''
    except Exception:
        return ''


# ── 시작할 때 확인 (백그라운드 스레드) ───────────────────────

class StatusBannerNotifier(QObject):
    """확인 결과를 제목 표시줄에 잠깐 보여주기 위한 신호 (메인 스레드로)."""
    status_signal = pyqtSignal(bool)     # 예전 호환: True = 업데이트 있음
    message_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()


class UpdateNotifier(QObject):
    """백그라운드 스레드 → Qt 메인 스레드로 안전하게 알림 창 요청."""
    notify_signal = pyqtSignal(str, str, str)   # (version, download_url, changelog)
    failed_signal = pyqtSignal(str)             # 자동 설치를 여러 번 실패한 버전

    def __init__(self):
        super().__init__()
        self.notify_signal.connect(self._on_notify, Qt.QueuedConnection)
        self.failed_signal.connect(self._on_failed, Qt.QueuedConnection)

    def emit_notify(self, version: str, download_url: str, changelog: str):
        self.notify_signal.emit(version, download_url, changelog)

    def _on_notify(self, version: str, download_url: str, changelog: str):
        show_update_notification(version, changelog)

    def _on_failed(self, version: str):
        QMessageBox.information(
            None, 'SSNnote 업데이트',
            f'{version} 자동 설치가 잘 되지 않았어요.\n메뉴 → 도움말 → 업데이트 확인에서 직접 설치해 주세요.')


def check_for_update_on_startup(notifier: UpdateNotifier, banner_notifier=None):
    """앱 시작 시 백그라운드 스레드에서. 자동 업데이트면 조용히 내려받아 두고, 아니면 한 번 알린다."""
    def say(msg):
        if banner_notifier:
            banner_notifier.message_signal.emit(msg)

    info = fetch_version_info()
    if not info or not is_newer_version(info.get('version', ''), APP_VERSION):
        return
    version = info['version']

    if not auto_enabled() or not getattr(sys, 'frozen', False):
        if _get_last_notified_version() != version:
            notifier.emit_notify(version, info.get('download_url', ''), info.get('changelog', ''))
        return

    cur = ready_update()
    if cur and cur['version'] == version and gave_up(cur):
        if _get_last_notified_version() != version:
            _save_notified_version(version)
            notifier.failed_signal.emit(version)
        return
    try:
        stage(info)
    except Exception as e:
        print('[UPDATE] 미리 받기 실패:', e)
        return
    say(f'{version} 업데이트를 받아 두었어요 — 앱을 끄거나 다시 켜면 설치돼요.')


# ── 메뉴 → 업데이트 확인 ──────────────────────────────────────

class ManualUpdateSignal(QObject):
    show_dialog = pyqtSignal(object)      # version info dict
    show_offline = pyqtSignal()
    show_uptodate = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.show_dialog.connect(lambda info: show_update_dialog(info), Qt.QueuedConnection)
        self.show_offline.connect(lambda: _show_offline_msg(), Qt.QueuedConnection)
        self.show_uptodate.connect(lambda: _show_up_to_date_msg(), Qt.QueuedConnection)


_manual_signal = ManualUpdateSignal()


def check_for_update_manual(parent=None):
    def _check():
        info = fetch_version_info()
        if info is None:
            _manual_signal.show_offline.emit()
        elif not is_newer_version(info.get('version', ''), APP_VERSION):
            _manual_signal.show_uptodate.emit()
        else:
            _manual_signal.show_dialog.emit(info)
    threading.Thread(target=_check, daemon=True).start()


def _korean_buttons(msg: QMessageBox):
    for btn in msg.buttons():
        if 'Details' in btn.text():
            btn.setText('자세히')


def show_update_notification(version: str, changelog: str = '', parent=None):
    """자동 업데이트를 끈 경우: 새 버전 안내 한 번."""
    msg = QMessageBox(parent)
    msg.setWindowTitle('SSNnote 업데이트')
    msg.setText(f'새 버전이 있습니다.\n\n현재 버전: {APP_VERSION}\n최신 버전: {version}\n\n'
                '메뉴 → 도움말 → 업데이트 확인에서 설치할 수 있어요.')
    msg.setDetailedText(changelog or '업데이트 내역 없음')
    msg.setStandardButtons(QMessageBox.Ok)
    msg.button(QMessageBox.Ok).setText('확인')
    _korean_buttons(msg)
    msg.exec_()
    _save_notified_version(version)


def show_update_dialog(info: dict, parent=None):
    """메뉴에서 확인: [지금 업데이트] 한 번이면 받아서 설치하고 다시 켜진다."""
    msg = QMessageBox(parent)
    msg.setWindowTitle('SSNnote 업데이트')
    msg.setText(f"현재 버전: {APP_VERSION}\n최신 버전: {info['version']}\n\n"
                '지금 업데이트할까요? 잠깐 꺼졌다가 다시 켜져요.')
    msg.setDetailedText(info.get('changelog') or '업데이트 내역 없음')
    b_now = msg.addButton('지금 업데이트', QMessageBox.AcceptRole)
    msg.addButton('나중에', QMessageBox.RejectRole)
    msg.setDefaultButton(b_now)
    _korean_buttons(msg)
    msg.exec_()
    _save_notified_version(info['version'])
    if msg.clickedButton() is b_now:
        download_and_install(info, parent)


def download_and_install(info: dict, parent=None):
    """받기(진행 표시) → 확인 → 화면 없이 설치 → 앱 종료 (설치 뒤 자동으로 다시 켜짐)."""
    if not getattr(sys, 'frozen', False):
        QMessageBox.information(parent, '개발 환경', '개발 환경에서는 업데이트를 지원하지 않습니다.')
        return
    cur = ready_update()
    if not (cur and cur['version'] == info['version']):
        progress = QProgressDialog('업데이트 받는 중…', '취소', 0, 100, parent)
        progress.setWindowTitle('SSNnote 업데이트')
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        def on_progress(done, total):
            if total:
                progress.setValue(int(done * 100 / total))
            QApplication.processEvents()
            return not progress.wasCanceled()
        try:
            clear_staged()
            path = os.path.join(_UPDATE_DIR, f"SSNnote_Setup-{info['version']}.exe")
            download(info, path, on_progress)
            cur = {'version': info['version'], 'path': path, 'sha256': info['sha256'].upper(),
                   'download_url': info['download_url'], 'changelog': info.get('changelog', ''), 'attempts': 0}
            _write_ready(cur)
        except Exception as e:
            progress.close()
            if not progress.wasCanceled():
                QMessageBox.warning(parent, '업데이트 실패', f'업데이트를 받지 못했습니다.\n{e}')
            return
        progress.close()
    if not launch_installer(cur, relaunch=True):
        QMessageBox.warning(parent, '업데이트 실패', '설치 프로그램을 실행하지 못했습니다.')
        return
    QApplication.quit()


def _show_offline_msg(parent=None):
    QMessageBox.information(parent, '업데이트 확인', '서버에 연결할 수 없습니다.\n인터넷 연결을 확인해 주세요.')


def _show_up_to_date_msg(parent=None):
    QMessageBox.information(parent, '업데이트 확인', f'현재 최신 버전({APP_VERSION})을 사용 중입니다.')
