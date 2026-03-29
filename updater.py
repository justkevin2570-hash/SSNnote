"""
자동 업데이트 모듈 (version.json 기반)
- 앱 시작 시 자동으로 새 버전 확인 (백그라운드 스레드)
- 수동으로 메뉴에서 "업데이트 확인" 가능
- Setup 파일 자동 다운로드 및 실행
- 1번만 알림 (last_notified_version.txt로 관리)
"""

import os
import sys
import json
import threading
import tempfile
import urllib.request
import urllib.error

from PyQt5.QtCore import QObject, pyqtSignal, Qt
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QApplication

APP_VERSION = 'v1.55'

_VERSION_JSON_URL = 'https://raw.githubusercontent.com/justkevin2570-hash/SSNnote/main/version.json'
_APPDATA_DIR = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
_NOTIFIED_FILE = os.path.join(_APPDATA_DIR, 'last_notified_version.txt')
_REQUEST_TIMEOUT = 10


def fetch_version_info() -> dict | None:
    """version.json에서 버전 정보를 가져온다. 실패 시 None 반환."""
    try:
        req = urllib.request.Request(
            _VERSION_JSON_URL,
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
    try:
        with open(_NOTIFIED_FILE, encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return ''


def _save_notified_version(tag: str):
    os.makedirs(_APPDATA_DIR, exist_ok=True)
    try:
        with open(_NOTIFIED_FILE, 'w', encoding='utf-8') as f:
            f.write(tag)
    except Exception:
        pass


class UpdateNotifier(QObject):
    notify_signal = pyqtSignal(str, str, str)  # (version, download_url, changelog)

    def __init__(self):
        super().__init__()
        self.notify_signal.connect(self._on_notify, Qt.QueuedConnection)

    def emit_notify(self, version: str, download_url: str, changelog: str):
        self.notify_signal.emit(version, download_url, changelog)

    def _on_notify(self, version: str, download_url: str, changelog: str):
        show_update_dialog(version, download_url, changelog)


class ManualUpdateSignal(QObject):
    show_dialog = pyqtSignal(str, str, str)
    show_offline = pyqtSignal()
    show_uptodate = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.show_dialog.connect(self._on_show_dialog, Qt.QueuedConnection)
        self.show_offline.connect(self._on_show_offline, Qt.QueuedConnection)
        self.show_uptodate.connect(self._on_show_uptodate, Qt.QueuedConnection)

    def _on_show_dialog(self, version: str, download_url: str, changelog: str):
        show_update_dialog(version, download_url, changelog)

    def _on_show_offline(self):
        _show_offline_msg()

    def _on_show_uptodate(self):
        _show_up_to_date_msg()


_manual_signal = ManualUpdateSignal()


def check_for_update_on_startup(notifier: UpdateNotifier):
    """백그라운드 스레드 타깃. 앱 시작 시 호출."""
    version_info = fetch_version_info()
    if not version_info:
        return

    remote_version = version_info.get('version', '')
    if not is_newer_version(remote_version, APP_VERSION):
        return

    if _get_last_notified_version() == remote_version:
        return

    download_url = version_info.get('download_url', '')
    changelog = version_info.get('changelog', '업데이트 내역 없음')
    notifier.emit_notify(remote_version, download_url, changelog)


def check_for_update_manual(parent=None):
    """메뉴에서 수동 확인."""
    def _check():
        version_info = fetch_version_info()
        if version_info is None:
            _manual_signal.show_offline.emit()
            return
        remote_version = version_info.get('version', '')
        if not is_newer_version(remote_version, APP_VERSION):
            _manual_signal.show_uptodate.emit()
            return
        download_url = version_info.get('download_url', '')
        changelog = version_info.get('changelog', '')
        _manual_signal.show_dialog.emit(remote_version, download_url, changelog)

    threading.Thread(target=_check, daemon=True).start()


def show_update_dialog(version: str, download_url: str, changelog: str, parent=None):
    msg = QMessageBox(parent)
    msg.setWindowTitle('SSNnote 업데이트')
    msg.setText(f'현재 버전: {APP_VERSION}\n최신 버전: {version}\n\n업데이트 하시겠어요?')
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msg.setDefaultButton(QMessageBox.Yes)
    msg.setDetailedText(changelog if changelog else '업데이트 내역 없음')

    yes_btn = msg.button(QMessageBox.Yes)
    no_btn = msg.button(QMessageBox.No)
    if yes_btn:
        yes_btn.setText('예')
    if no_btn:
        no_btn.setText('아니오')
    for btn in msg.buttons():
        if 'Details' in btn.text():
            btn.setText('자세히')

    result = msg.exec_()
    _save_notified_version(version)

    if result == QMessageBox.Yes:
        download_and_install(version, download_url, changelog, parent)


def download_and_install(version: str, download_url: str, changelog: str, parent=None):
    if not getattr(sys, 'frozen', False):
        QMessageBox.information(parent, '개발 환경',
            '개발 환경에서는 업데이트를 지원하지 않습니다.\n'
            'PyInstaller로 빌드된 환경에서 실행해주세요.')
        return

    if not download_url:
        QMessageBox.warning(parent, '다운로드 실패', '설치 파일 URL을 찾을 수 없습니다.')
        return

    tmp_setup = os.path.join(tempfile.gettempdir(), 'SSNnote_Setup.exe')

    progress = QProgressDialog('업데이트 다운로드 중...', '취소', 0, 0, parent)
    progress.setWindowTitle('SSNnote 업데이트')
    progress.setWindowModality(Qt.WindowModal)
    progress.show()
    QApplication.processEvents()

    try:
        urllib.request.urlretrieve(download_url, tmp_setup)
    except Exception as e:
        progress.close()
        QMessageBox.warning(parent, '다운로드 실패', f'다운로드 중 오류가 발생했습니다.\n{e}')
        return

    progress.close()

    try:
        os.startfile(tmp_setup)
    except Exception as e:
        QMessageBox.warning(parent, '설치 실패', f'설치 파일을 실행할 수 없습니다.\n{e}')
        return

    QApplication.quit()
    sys.exit(0)


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
