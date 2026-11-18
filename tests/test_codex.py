"""ChatGPT 구독(Codex CLI) 연결 — 가짜 codex로 명령줄·표준입력·환경변수·답 파일·오류 확인 (실제 구독 호출 없음)."""
import os
import subprocess

import pytest

pytest.importorskip('PyQt5')

FAKE_EXE = 'C:/fake/codex.exe'


class FakeProc:
    script = {}
    calls = []

    def __init__(self, args, **kw):
        self.args, self.kw = args, kw
        self.input = None
        self.killed = False
        self.returncode = FakeProc.script.get('returncode', 0)
        FakeProc.calls.append(self)

    def communicate(self, input=None, timeout=None):
        self.input = input
        if FakeProc.script.get('timeout') and not self.killed:
            raise subprocess.TimeoutExpired(self.args, timeout)
        ans = FakeProc.script.get('answer')
        if ans is not None and '-o' in self.args:          # codex처럼 -o 파일에 최종 답을 쓴다
            with open(os.path.join(self.kw['cwd'], self.args[self.args.index('-o') + 1]), 'w', encoding='utf-8') as f:
                f.write(ans)
        return FakeProc.script.get('stdout', ''), FakeProc.script.get('stderr', '')

    def kill(self):
        self.killed = True

    def terminate(self):
        pass


@pytest.fixture
def cx(tmp_path, monkeypatch):
    import ai_client as ai
    monkeypatch.setattr(ai, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai, 'CODEX_PATH_FILE', str(tmp_path / 'codex_path.txt'))
    monkeypatch.setattr(ai, 'CODEX_EFFORT_FILE', str(tmp_path / 'codex_effort.txt'))
    monkeypatch.setattr(ai.shutil, 'which', lambda name: FAKE_EXE if name == 'codex' else None)
    monkeypatch.setattr(ai.subprocess, 'Popen', FakeProc)
    monkeypatch.setenv('OPENAI_API_KEY', 'sk-should-not-pass')
    FakeProc.script = {'answer': '안녕하세요'}
    FakeProc.calls = []
    yield ai
    FakeProc.script = {}


def test_ask_uses_readonly_ephemeral_exec_and_reads_answer_file(cx):
    FakeProc.script = {'answer': '2. 장소: 시청각실\n'}
    assert cx.CodexAdapter.complete('지침입니다', '장소 바꿔줘') == '2. 장소: 시청각실'
    p = FakeProc.calls[-1]
    a = p.args
    assert a[0] == FAKE_EXE and a[1] == 'exec' and a[-1] == '-'
    for flag in ('--ephemeral', '--skip-git-repo-check'):
        assert flag in a
    assert a[a.index('--sandbox') + 1] == 'read-only'
    assert 'model_reasoning_effort=low' in a                   # 기본은 빠르게
    assert '지침입니다' in p.input and '장소 바꿔줘' in p.input   # 지침 + 요청은 표준입력으로
    assert 'OPENAI_API_KEY' not in p.kw['env']                  # 구독 로그인을 쓰도록 API 키는 빼고
    assert not os.path.exists(p.kw['cwd'])                      # 빈 임시 폴더는 지움


def test_effort_setting_and_label(cx):
    cx.save_codex_effort('high')
    cx.CodexAdapter.complete('s', 'p')
    assert 'model_reasoning_effort=high' in FakeProc.calls[-1].args
    cx.save_codex_effort('weird')
    assert cx.load_codex_effort() == 'low'
    cx.save_ai_mode('codex')
    assert cx.load_ai_mode() == 'codex' and cx.get_adapter() is cx.CodexAdapter
    assert cx.ai_status_text() == 'ChatGPT 구독 · 빠르게'
    assert cx.CodexAdapter.supports_image is False


def test_errors_are_friendly(cx):
    FakeProc.script = {'returncode': 1, 'answer': None, 'stderr': 'Error: Not logged in'}
    with pytest.raises(cx.ClaudeCodeError, match='로그인'):
        cx.CodexAdapter.complete('s', 'p')
    FakeProc.script = {'returncode': 1, 'answer': None, 'stderr': "You've hit your usage limit"}
    with pytest.raises(cx.ClaudeCodeError, match='한도'):
        cx.CodexAdapter.complete('s', 'p')
    FakeProc.script = {'timeout': True}
    with pytest.raises(cx.ClaudeCodeError, match='오래'):
        cx.CodexAdapter.complete('s', 'p')
    errs = []
    FakeProc.script = {'returncode': 0, 'answer': ''}
    cx.CodexAdapter.stream_text('p', lambda c: None, lambda d: None, errs.append)
    assert errs and 'ChatGPT' in errs[0]


def test_not_installed(cx, monkeypatch):
    monkeypatch.setattr(cx, 'find_codex_cli', lambda: None)
    with pytest.raises(cx.ClaudeCodeError, match='설치'):
        cx.CodexAdapter.complete('s', 'p')
    assert FakeProc.calls == []
    assert cx.codex_status() == {'installed': False, 'logged_in': False, 'email': '', 'plan': '', 'path': ''}


def test_status_requires_chatgpt_login(cx):
    FakeProc.script = {'returncode': 0, 'stderr': 'Logged in using ChatGPT'}
    st = cx.codex_status()
    assert st['installed'] and st['logged_in'] and FakeProc.calls[-1].args[1:] == ['login', 'status']
    FakeProc.script = {'returncode': 0, 'stderr': 'Logged in using an API key - sk-***'}
    assert cx.codex_status()['logged_in'] is False              # API 키 로그인은 구독이 아님
    FakeProc.script = {'returncode': 1, 'stderr': 'Not logged in'}
    assert cx.codex_status()['logged_in'] is False


def test_finds_official_install_location(cx, monkeypatch, tmp_path):
    monkeypatch.setattr(cx.shutil, 'which', lambda name: None)
    exe = tmp_path / 'Programs' / 'OpenAI' / 'Codex' / 'bin' / 'codex.exe'
    exe.parent.mkdir(parents=True)
    exe.write_text('')
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    assert cx.find_codex_cli() == str(exe)


def test_menu_order_and_labels():
    from ui import ai_connect
    assert ai_connect.MENU_ORDER == ('claude_code', 'codex', None, 'opencode', 'claude', 'openai')
    assert [ai_connect.MENU_LABELS[k] for k in ai_connect.MENU_ORDER if k] == [
        'Claude 구독 계정 연동', 'ChatGPT 구독 계정 연동', 'OpenCode API 연동', 'Claude API 연동', 'ChatGPT API 연동']


def test_codex_dialog(tmp_path, monkeypatch):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    import ai_client as ai
    import document_editor as de
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai, 'CODEX_PATH_FILE', str(tmp_path / 'codex_path.txt'))
    monkeypatch.setattr(ai, 'CODEX_EFFORT_FILE', str(tmp_path / 'codex_effort.txt'))
    state = {'installed': True, 'logged_in': False, 'email': '', 'plan': '', 'path': 'x'}
    monkeypatch.setattr(de, 'codex_status', lambda: dict(state))
    calls = []
    monkeypatch.setattr(de, 'codex_login', lambda: calls.append('login'))
    monkeypatch.setattr(de.QMessageBox, 'warning', lambda *a, **k: calls.append(a[2]))
    dlg = de.ClaudeCodeDialog(kind='codex')
    for _ in range(20):
        if dlg._job is not None:
            dlg._job.wait(2000)
        app.processEvents()
    assert dlg.windowTitle() == 'ChatGPT 구독 계정 연동'
    assert dlg.btn_account.text() == 'ChatGPT로 로그인' and '로그인 필요' in dlg.lbl_status.text()
    assert [dlg.cmb_model.itemText(i) for i in range(dlg.cmb_model.count())] == ['빠르게', '보통', '꼼꼼하게']
    dlg._save()
    assert any('로그인' in str(c) for c in calls) and ai.load_ai_mode() != 'codex'
    state['logged_in'] = True
    dlg._st = dict(state)
    dlg.cmb_model.setCurrentIndex(1)
    dlg._save()
    assert ai.load_ai_mode() == 'codex' and ai.load_codex_effort() == 'medium'
