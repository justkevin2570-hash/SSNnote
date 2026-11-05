"""ui/claude_connect_dialog.py — 「Claude 연결」 창 (⚙️ 설정 → Claude 연결…).

Claude Code / Claude 데스크톱 앱을 찾아 서서니 노트 연결을 버튼 한 번으로 등록·해제한다.
실제 작업은 claude_connect.py, 오래 걸리는 확인은 백그라운드 스레드에서.
"""

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout,
)

import claude_connect as cc
from .common import pr_font
from .theme import T, tc

# 창을 닫아도 진행 중인 작업 스레드가 끝날 때까지 참조 유지
_ALIVE = set()

_STYLE = f"""
    QDialog {{ background: {T.BG_WARM}; }}
    QLabel {{ color: {T.TEXT_WARM}; background: transparent; }}
    QLabel#Title {{ color: {T.INK}; font-size: 14pt; font-weight: 700; }}
    QLabel#Hint, QLabel#Muted {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QLabel#Name {{ color: {T.INK}; font-size: 11pt; font-weight: 700; }}
    QLabel#Status[state="ok"] {{ color: {T.TEXT_ON_NOTE}; font-weight: 700; }}
    QLabel#Status[state="warn"] {{ color: {T.HIGHLIGHT}; }}
    QLabel#Status[state="off"] {{ color: {T.TEXT_WARM_MUTED}; }}
    QLabel#Banner {{ background: {T.HIGHLIGHT_BG}; color: {T.INK}; border-radius: 6px; padding: 8px 10px; }}
    QFrame#Card {{ background: {T.SURFACE}; border: 1px solid {T.BORDER_WARM}; border-radius: 8px; }}
    QPushButton {{
        background: {T.SURFACE_WARM}; color: {T.TEXT_ON_NOTE}; border: none; border-radius: 5px;
        padding: 6px 14px;
    }}
    QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
    QPushButton:disabled {{ color: {T.TEXT_WARM_MUTED}; }}
    QPushButton#Primary {{ background: {T.ACCENT}; color: {tc("white")}; font-weight: 700; }}
    QPushButton#Primary:hover {{ background: {T.ACCENT_HOVER}; }}
    QPushButton#Primary:disabled {{ background: {T.SURFACE_WARM}; color: {T.TEXT_WARM_MUTED}; }}
"""


class _Job(QThread):
    done = pyqtSignal(object, str)        # (결과, 오류 메시지)

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            self.done.emit(self._fn(), '')
        except cc.ConnectError as e:
            self.done.emit(None, str(e))
        except Exception as e:  # 예상 못 한 오류도 창에 표시
            self.done.emit(None, f'오류: {e}')


class _Row(QFrame):
    """Claude Code / 데스크톱 앱 한 줄: 이름·설명·상태 + [연결] [연결 해제]."""

    def __init__(self, name, desc):
        super().__init__()
        self.setObjectName('Card')
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 10, 12, 10)
        outer.setSpacing(4)
        lay = QHBoxLayout()
        text = QVBoxLayout()
        text.setSpacing(2)
        n = QLabel(name)
        n.setObjectName('Name')
        d = QLabel(desc)                      # 짧은 한 줄 (버튼 옆 좁은 칸이라 줄바꿈하지 않음)
        d.setObjectName('Muted')
        self.status = QLabel('확인 중…')
        self.status.setObjectName('Status')
        for w in (n, d, self.status):
            text.addWidget(w)
        lay.addLayout(text, 1)
        self.btn_connect = QPushButton('연결')
        self.btn_connect.setObjectName('Primary')
        self.btn_disconnect = QPushButton('연결 해제')
        lay.addWidget(self.btn_disconnect, 0, Qt.AlignVCenter)
        lay.addWidget(self.btn_connect, 0, Qt.AlignVCenter)
        outer.addLayout(lay)
        self.note = QLabel('')                # 긴 안내는 카드 전체 폭으로 아래에
        self.note.setObjectName('Muted')
        self.note.setWordWrap(True)
        self.note.hide()
        outer.addWidget(self.note)
        self.set_state('확인 중…', 'off', connect=False, disconnect=False)

    def set_state(self, text, state, connect=True, disconnect=False, connect_label='연결', note=''):
        self.note.setText(note)
        self.note.setVisible(bool(note))
        self.status.setText(text)
        self.status.setProperty('state', state)
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)
        self.btn_connect.setText(connect_label)
        self.btn_connect.setEnabled(connect)
        self.btn_disconnect.setVisible(disconnect)


class ClaudeConnectDialog(QDialog):
    def __init__(self, parent=None, auto_check=True):
        super().__init__(parent, Qt.Window | Qt.WindowStaysOnTopHint)
        self.setWindowTitle('Claude 연결')
        self.setStyleSheet(_STYLE)
        self.setFont(pr_font(10))
        self.setFixedWidth(500)
        self._jobs = []
        self._build()
        if auto_check:
            self.refresh()

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 14)
        root.setSpacing(10)
        title = QLabel('🤖 Claude 연결')
        title.setObjectName('Title')
        root.addWidget(title)
        hint = QLabel('연결하면 Claude 앱에서 서서니 노트 업무를 보고, 추가하고, 공문 초안을 쓸 수 있어요.\n'
                      'Claude 구독(Pro·Max)으로 쓰며 API 키는 필요 없습니다.')
        hint.setObjectName('Hint')
        hint.setWordWrap(True)
        root.addWidget(hint)

        self.banner = QLabel('')
        self.banner.setObjectName('Banner')
        self.banner.setWordWrap(True)
        self.banner.hide()
        root.addWidget(self.banner)
        self.btn_install = QPushButton('필요한 구성요소 설치')
        self.btn_install.hide()
        self.btn_install.clicked.connect(self._install)
        root.addWidget(self.btn_install, 0, Qt.AlignLeft)

        self.row_code = _Row('Claude Code', '명령 프롬프트용 Claude')
        self.row_desk = _Row('Claude 데스크톱 앱', 'Windows용 Claude 앱')
        root.addWidget(self.row_code)
        root.addWidget(self.row_desk)
        self.row_code.btn_connect.clicked.connect(lambda: self._do(self.row_code, cc.connect_claude_code, '연결하는 중…'))
        self.row_code.btn_disconnect.clicked.connect(lambda: self._do(self.row_code, cc.disconnect_claude_code, '해제하는 중…'))
        self.row_desk.btn_connect.clicked.connect(lambda: self._do(self.row_desk, cc.connect_desktop, '연결하는 중…'))
        self.row_desk.btn_disconnect.clicked.connect(lambda: self._do(self.row_desk, cc.disconnect_desktop, '해제하는 중…'))

        ex = QLabel('연결 후 Claude에게: "오늘 할 일 정리해줘", "다음 주 금요일까지 ○○ 제출 업무로 넣어줘",\n'
                    '"교무회의 안내 공문 초안 써줘"')
        ex.setObjectName('Muted')
        ex.setWordWrap(True)
        root.addWidget(ex)

        foot = QHBoxLayout()
        foot.addStretch()
        btn_close = QPushButton('닫기')
        btn_close.clicked.connect(self.close)
        foot.addWidget(btn_close)
        root.addLayout(foot)

    # ── 상태 확인 ────────────────────────────────────────────────
    def refresh(self):
        try:
            cc.server_command()
            self.banner.hide()
            self.btn_install.hide()
            ready = True
        except cc.ConnectError as e:
            self.banner.setText('⚠ ' + str(e))
            self.banner.show()
            self.btn_install.setVisible(not _frozen())
            ready = False
        self._ready = ready
        self.row_code.set_state('확인 중…', 'off', connect=False)
        self.row_desk.set_state('확인 중…', 'off', connect=False)
        self._start(cc.claude_code_status, self._show_code)
        self._show_desk(cc.desktop_status(), '')          # 파일만 읽으므로 바로

    def _show_code(self, s, err):
        r = self.row_code
        if err:
            r.set_state('문제가 있어요', 'warn', connect=self._ready, note=err)
        elif not s['installed']:
            r.set_state('설치되어 있지 않음', 'off', connect=False,
                        note='Claude Code를 설치한 뒤 이 창을 다시 열어주세요.')
        elif s['connected']:
            r.set_state('✓ 연결됨', 'ok', connect=self._ready, disconnect=True, connect_label='다시 연결')
        elif s['registered']:
            r.set_state('등록됐지만 응답 없음', 'warn', connect=self._ready, disconnect=True,
                        connect_label='다시 연결', note='[다시 연결]을 눌러보세요.')
        else:
            r.set_state('연결 안 됨', 'off', connect=self._ready)
        self._fit()

    def _show_desk(self, s, err):
        r = self.row_desk
        if err:
            r.set_state('문제가 있어요', 'warn', connect=self._ready, note=err)
        elif not s['installed']:
            r.set_state('설치되어 있지 않음', 'off', connect=False,
                        note='Claude 데스크톱 앱을 설치한 뒤 이 창을 다시 열어주세요.')
        elif s['registered'] and s['up_to_date']:
            r.set_state('✓ 연결됨', 'ok', connect=self._ready, disconnect=True, connect_label='다시 연결',
                        note='Claude 앱을 완전히 종료(트레이 아이콘까지)했다가 다시 켜면 적용됩니다.')
        elif s['registered']:
            r.set_state('예전 위치로 등록됨', 'warn', connect=self._ready, disconnect=True,
                        connect_label='다시 연결', note='서서니 노트를 옮기거나 다시 설치했다면 [다시 연결]을 눌러주세요.')
        else:
            r.set_state('연결 안 됨', 'off', connect=self._ready,
                        note='Claude 앱을 먼저 종료한 뒤 [연결]을 누르면 가장 확실합니다.')
        self._fit()

    def _fit(self):
        """안내 줄이 생기거나 사라지면 창 높이를 맞춘다 (폭은 고정)."""
        self.layout().activate()
        self.resize(self.width(), self.sizeHint().height())

    # ── 동작 ────────────────────────────────────────────────────
    def _start(self, fn, on_done):
        job = _Job(fn)
        _ALIVE.add(job)
        job.finished.connect(lambda: _ALIVE.discard(job))
        job.done.connect(lambda res, err: None if self._closed() else on_done(res, err))
        self._jobs.append(job)
        job.start()

    def _closed(self):
        try:
            return not self.isVisible()
        except RuntimeError:  # 창이 이미 정리됨
            return True

    def _do(self, row, fn, busy_text):
        row.set_state(busy_text, 'off', connect=False)
        on_done = self._show_code if row is self.row_code else self._show_desk
        self._start(fn, on_done)

    def _install(self):
        import subprocess
        self.btn_install.setEnabled(False)
        self.banner.setText('구성요소를 설치하는 중… (1~2분 걸릴 수 있어요)')

        def job():
            p = subprocess.run(cc.install_requirements_command(), capture_output=True, text=True)
            if p.returncode != 0:
                raise cc.ConnectError('설치에 실패했습니다.\n' + (p.stderr or p.stdout)[-300:])
            return True

        def done(_, err):
            self.btn_install.setEnabled(True)
            if err:
                self.banner.setText('⚠ ' + err)
                return
            # 같은 프로그램에서 새로 설치한 패키지를 찾도록 import 캐시 갱신
            import importlib
            importlib.invalidate_caches()
            self.refresh()
        self._start(job, done)


def _frozen():
    import sys
    return getattr(sys, 'frozen', False)
