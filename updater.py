"""
자동 업데이트 모듈 (GitHub Releases API 기반)
- 앱 시작 시 자동으로 새 버전 확인 (백그라운드 스레드)
- 수동으로 "업데이트 확인" 메뉴에서 확인 가능
- exe 자동 다운로드 및 교체
- 1번만 알림 (last_notified_version.txt로 관리)
"""

import os
import sys
import json
import subprocess
import tempfile
import threading
import urllib.request
import urllib.error

from PyQt5.QtCore import QObject, pyqtSignal, Qt, QTimer
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication

APP_VERSION = 'v0.2'

_GITHUB_API_URL = 'https://api.github.com/repos/justkevin2570-hash/SSNnote/releases/latest'
_APPDATA_DIR = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
_NOTIFIED_FILE = os.path.join(_APPDATA_DIR, 'last_notified_version.txt')
_REQUEST_TIMEOUT = 10


def fetch_latest_release() -> dict | None:
    """GitHub Releases API에서 최신 릴리스 정보를 가져온다. 실패 시 None 반환."""
    try:
        req = urllib.request.Request(
            _GITHUB_API_URL,
            headers={'User-Agent': f'SSNnote/{APP_VERSION}'},
        )
        with urllib.request.urlopen(req, timeout=_REQUEST_TIMEOUT) as r:
            return json.loads(r.read())
    except Exception:
        return None


def is_newer_version(remote_tag: str, local_tag: str) -> bool:
    """'v1.8' > 'v1.7' 형식 비교. 파싱 실패 시 False 반환."""
    try:
        def _parse(tag):
            return tuple(int(x) for x in tag.lstrip('v').split('.'))
        return _parse(remote_tag) > _parse(local_tag)
    except Exception:
        return False


def _get_last_notified_version() -> str:
    """마지막 알림 버전을 읽음. 파일 없음/실패 시 '' 반환."""
    try:
        with open(_NOTIFIED_FILE, encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return ''


def _save_notified_version(tag: str):
    """알림 보낸 버전을 파일에 기록."""
    os.makedirs(_APPDATA_DIR, exist_ok=True)
    try:
        with open(_NOTIFIED_FILE, 'w', encoding='utf-8') as f:
            f.write(tag)
    except Exception:
        pass


class UpdateNotifier(QObject):
    """백그라운드 스레드 → Qt 메인 스레드로 안전하게 신호 전달."""
    notify_signal = pyqtSignal(str, str)  # (tag_name, changelog)

    def __init__(self):
        super().__init__()
        self.notify_signal.connect(self._on_notify, Qt.QueuedConnection)

    def emit_notify(self, tag_name: str, changelog: str):
        """백그라운드 스레드에서 호출 가능."""
        self.notify_signal.emit(tag_name, changelog)

    def _on_notify(self, tag_name: str, changelog: str):
        """메인 스레드에서 실행됨."""
        show_update_dialog(tag_name, changelog)


class ManualUpdateSignal(QObject):
    """수동 업데이트 확인용 신호."""
    show_dialog = pyqtSignal(str, str)  # (tag_name, changelog)
    show_offline = pyqtSignal()
    show_uptodate = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.show_dialog.connect(self._on_show_dialog, Qt.QueuedConnection)
        self.show_offline.connect(self._on_show_offline, Qt.QueuedConnection)
        self.show_uptodate.connect(self._on_show_uptodate, Qt.QueuedConnection)

    def _on_show_dialog(self, tag_name: str, changelog: str):
        show_update_dialog(tag_name, changelog)

    def _on_show_offline(self):
        _show_offline_msg()

    def _on_show_uptodate(self):
        _show_up_to_date_msg()


_manual_signal = ManualUpdateSignal()


def check_for_update_on_startup(notifier: UpdateNotifier):
    """
    백그라운드 스레드 타깃. 앱 시작 시 호출.
    새 버전 있음 + 미알림 → 팝업 표시.
    """
    release = fetch_latest_release()
    if not release:
        return

    remote_tag = release.get('tag_name', '')
    if not is_newer_version(remote_tag, APP_VERSION):
        return

    if _get_last_notified_version() == remote_tag:
        return  # 이미 이 버전으로 알림 완료

    changelog = release.get('body', '업데이트 내역 없음')
    notifier.emit_notify(remote_tag, changelog)


def check_for_update_manual(parent=None):
    """트레이 메뉴에서 수동 확인."""
    def _check():
        print(f'[UPDATE] 버전 확인 시작. 현재 버전: {APP_VERSION}')
        release = fetch_latest_release()
        if release is None:
            print('[UPDATE] API 응답 없음')
            _manual_signal.show_offline.emit()
            return
        remote_tag = release.get('tag_name', '')
        print(f'[UPDATE] 최신 버전: {remote_tag}')
        if not is_newer_version(remote_tag, APP_VERSION):
            print(f'[UPDATE] 최신 버전입니다 ({APP_VERSION})')
            _manual_signal.show_uptodate.emit()
            return
        print(f'[UPDATE] 새 버전 있음: {APP_VERSION} → {remote_tag}')
        changelog = release.get('body', '')
        _manual_signal.show_dialog.emit(remote_tag, changelog)

    threading.Thread(target=_check, daemon=True).start()


def show_update_dialog(tag_name: str, changelog: str, parent=None):
    """팝업 표시. 예/아니오 모두 버전 기록."""
    msg = QMessageBox(parent)
    msg.setWindowTitle('SSNnote 업데이트')
    msg.setText(f'현재 버전: {APP_VERSION}\n최신 버전: {tag_name}\n\n업데이트 하시겠어요?')
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msg.setDefaultButton(QMessageBox.Yes)
    msg.setDetailedText(changelog if changelog else '업데이트 내역 없음')

    # 버튼 텍스트를 한글로 변경 (setDetailedText 이후에 호출)
    yes_btn = msg.button(QMessageBox.Yes)
    no_btn = msg.button(QMessageBox.No)
    if yes_btn:
        yes_btn.setText('예')
    if no_btn:
        no_btn.setText('아니오')

    # Show Details 버튼도 한글로 변경
    for btn in msg.buttons():
        text = btn.text()
        if 'Show Details' in text or 'Details' in text:
            btn.setText('자세히')

    result = msg.exec_()

    _save_notified_version(tag_name)

    if result == QMessageBox.Yes:
        download_and_install(tag_name, changelog, parent)


def _find_exe_asset(release_data: dict, tag_name: str) -> str | None:
    """release 정보에서 .exe 다운로드 URL 찾음."""
    for asset in release_data.get('assets', []):
        if asset.get('name', '').endswith('.exe'):
            return asset['browser_download_url']
    return None


def download_and_install(tag_name: str, changelog: str, parent=None):
    """exe 다운로드 → 교체 배치 스크립트 실행 → 앱 종료."""
    if not getattr(sys, 'frozen', False):
        QMessageBox.information(parent, '개발 환경',
            '개발 환경에서는 업데이트를 지원하지 않습니다.\n'
            'PyInstaller로 빌드된 exe 환경에서 실행해주세요.')
        return

    release = fetch_latest_release()
    if not release:
        QMessageBox.warning(parent, '다운로드 실패', '서버에 연결할 수 없습니다.')
        return

    url = _find_exe_asset(release, tag_name)
    if not url:
        QMessageBox.warning(parent, '다운로드 실패', '설치 파일을 찾을 수 없습니다.')
        return

    tmp_exe = os.path.join(tempfile.gettempdir(), 'SSNnote_new.exe')

    progress = QProgressDialog('업데이트 다운로드 중...', '취소', 0, 0, parent)
    progress.setWindowTitle('SSNnote 업데이트')
    progress.setWindowModality(Qt.WindowModal)
    progress.show()
    QApplication.processEvents()

    try:
        urllib.request.urlretrieve(url, tmp_exe)
    except Exception as e:
        progress.close()
        QMessageBox.warning(parent, '다운로드 실패',
            f'다운로드 중 오류가 발생했습니다.\n{e}')
        return

    progress.close()

    current_exe = sys.executable
    bat_path = _write_update_batch(current_exe, tmp_exe)

    try:
        subprocess.Popen(
            ['cmd.exe', '/c', bat_path],
            creationflags=(
                subprocess.DETACHED_PROCESS
                | subprocess.CREATE_NEW_PROCESS_GROUP
                | subprocess.CREATE_NO_WINDOW
            ),
        )
    except Exception as e:
        QMessageBox.warning(parent, '업데이트 실패',
            f'업데이트를 시작할 수 없습니다.\n{e}')
        return

    QApplication.quit()


def _write_update_batch(current_exe: str, new_exe: str) -> str:
    """배치 스크립트 작성."""
    bat_path = os.path.join(tempfile.gettempdir(), 'ssnnote_update.bat')
    backup_exe = current_exe + '.bak'

    script = f"""@echo off
timeout /t 3 /nobreak > nul
move /y "{current_exe}" "{backup_exe}" > nul 2>&1
move /y "{new_exe}" "{current_exe}"
if errorlevel 1 (
    move /y "{backup_exe}" "{current_exe}" > nul 2>&1
    exit /b 1
)
start "" "{current_exe}"
del "%~f0"
"""
    with open(bat_path, 'w', encoding='cp949') as f:
        f.write(script)
    return bat_path


def _show_offline_msg(parent=None):
    msg = QMessageBox(parent)
    msg.setWindowTitle('업데이트 확인')
    msg.setText('서버에 연결할 수 없습니다.\n인터넷 연결을 확인해 주세요.')
    msg.setStandardButtons(QMessageBox.Ok)
    ok_btn = msg.button(QMessageBox.Ok)
    if ok_btn:
        ok_btn.setText('확인')
    msg.exec_()


def _show_up_to_date_msg(parent=None):
    msg = QMessageBox(parent)
    msg.setWindowTitle('업데이트 확인')
    msg.setText(f'현재 최신 버전({APP_VERSION})을 사용 중입니다.')
    msg.setStandardButtons(QMessageBox.Ok)
    ok_btn = msg.button(QMessageBox.Ok)
    if ok_btn:
        ok_btn.setText('확인')
    msg.exec_()
