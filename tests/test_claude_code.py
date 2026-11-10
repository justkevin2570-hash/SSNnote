"""Claude 구독(Claude Code CLI) 연결 — 가짜 claude로 명령줄·표준입력·환경변수·오류 메시지 확인 (실제 구독 호출 없음)."""
import json
import os
import subprocess

import pytest

pytest.importorskip('PyQt5')

FAKE_EXE = 'C:/fake/claude.exe'


class FakeProc:
    """subprocess.Popen 대신 쓰는 가짜. script는 테스트마다 다른 응답을 준다."""
    script = {}
    calls = []

    def __init__(self, args, **kw):
        self.args = args
        self.kw = kw
        self.cwd_existed = os.path.isdir(kw.get('cwd', ''))
        self.input = None
        self.killed = False
        self.terminated = False
        self.returncode = FakeProc.script.get('returncode', 0)
        FakeProc.calls.append(self)
        if FakeProc.script.get('popen_error'):
            raise FileNotFoundError(FakeProc.script['popen_error'])

    def communicate(self, input=None, timeout=None):
        self.input = input
        if FakeProc.script.get('hook'):
            FakeProc.script['hook']()
        if FakeProc.script.get('timeout') and not self.killed:
            raise subprocess.TimeoutExpired(self.args, timeout)
        return FakeProc.script.get('stdout', ''), FakeProc.script.get('stderr', '')

    def kill(self):
        self.killed = True

    def terminate(self):
        self.terminated = True


@pytest.fixture
def cc(tmp_path, monkeypatch):
    import ai_client as ai
    monkeypatch.setattr(ai, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai, 'CLAUDE_CODE_PATH_FILE', str(tmp_path / 'claude_code_path.txt'))
    monkeypatch.setattr(ai, 'CLAUDE_CODE_MODEL_FILE', str(tmp_path / 'claude_code_model.txt'))
    monkeypatch.setattr(ai.shutil, 'which', lambda name: FAKE_EXE)     # PATH에서 찾은 것처럼
    monkeypatch.setattr(ai.subprocess, 'Popen', FakeProc)
    FakeProc.script = {'stdout': json.dumps({'is_error': False, 'result': '안녕하세요'})}
    FakeProc.calls = []
    yield ai
    FakeProc.script = {}


def _reply(**extra):
    FakeProc.script = {'stdout': json.dumps({'is_error': False, 'result': '2. 장소: 시청각실', **extra})}


def test_command_line_is_print_json_without_tools_or_bare(cc):
    cc.ClaudeCodeAdapter.complete('시스템', '장소를 알려줘')
    args = FakeProc.calls[-1].args
    assert args[0] == FAKE_EXE
    for flag in ('-p', '--no-session-persistence'):
        assert flag in args
    assert args[args.index('--output-format') + 1] == 'json'
    assert args[args.index('--tools') + 1] == ''                     # 도구 전부 끄기
    assert args[args.index('--system-prompt') + 1] == '시스템'
    assert '--bare' not in args and '--model' not in args            # 기본 모델이면 --model을 넣지 않음


def test_prompt_goes_through_stdin(cc):
    cc.ClaudeCodeAdapter.complete('지시', '질문 본문 ' * 50)
    proc = FakeProc.calls[-1]
    assert proc.input == '질문 본문 ' * 50
    assert '질문 본문' not in ' '.join(proc.args)                    # 명령줄에는 질문이 없다


def test_api_keys_removed_and_run_in_empty_temp_folder(cc, monkeypatch):
    for name in ('ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_BASE_URL', 'CLAUDE_CODE_OAUTH_TOKEN'):
        monkeypatch.setenv(name, 'secret')
    cc.ClaudeCodeAdapter.complete('지시', '질문')
    proc = FakeProc.calls[-1]
    for name in ('ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_BASE_URL', 'CLAUDE_CODE_OAUTH_TOKEN'):
        assert name not in proc.kw['env']
    assert proc.kw['encoding'] == 'utf-8' and proc.kw['errors'] == 'replace'
    assert 'ssnnote_claude_' in proc.kw['cwd'] and proc.cwd_existed     # 실행 중에는 빈 임시 폴더
    assert not os.path.exists(proc.kw['cwd'])                            # 끝나면 지운다


def test_result_text_is_extracted(cc):
    _reply()
    assert cc.ClaudeCodeAdapter.complete('지시', '질문') == '2. 장소: 시청각실'


def test_model_option_only_when_chosen(cc):
    cc.save_claude_code_model('sonnet')
    cc.ClaudeCodeAdapter.complete('지시', '질문')
    args = FakeProc.calls[-1].args
    assert args[args.index('--model') + 1] == 'sonnet'
    cc.save_claude_code_model('gpt-x')                                  # 목록에 없는 값은 기본으로
    assert cc.load_claude_code_model() == ''


def test_long_system_prompt_moves_to_stdin(cc):
    long_system = '규칙 ' * 9000
    cc.ClaudeCodeAdapter.complete(long_system, '질문')
    proc = FakeProc.calls[-1]
    assert '--system-prompt' not in proc.args
    assert proc.input.startswith(long_system) and proc.input.endswith('질문')


def test_error_result_is_turned_into_korean_messages(cc):
    cases = [
        ({'is_error': True, 'result': 'You have hit your usage limit'}, '구독 사용 한도'),
        ({'is_error': True, 'result': 'Invalid API key · Please run /login'}, '로그인'),
        ({'is_error': True, 'result': '무언가 이상함'}, 'Claude 연결 실패'),
    ]
    for body, expected in cases:
        FakeProc.script = {'stdout': json.dumps(body)}
        with pytest.raises(cc.ClaudeCodeError, match=expected):
            cc.ClaudeCodeAdapter.complete('지시', '질문')


def test_nonzero_exit_without_json_is_an_error(cc):
    FakeProc.script = {'returncode': 1, 'stdout': '', 'stderr': 'not logged in, please authenticate'}
    with pytest.raises(cc.ClaudeCodeError, match='로그인'):
        cc.ClaudeCodeAdapter.complete('지시', '질문')


def test_timeout_is_killed_and_explained(cc):
    FakeProc.script = {'timeout': True}
    with pytest.raises(cc.ClaudeCodeError, match='너무 오래'):
        cc.ClaudeCodeAdapter.complete('지시', '질문')
    assert FakeProc.calls[-1].killed


def test_missing_executable_gives_install_message(cc, monkeypatch):
    monkeypatch.setattr(cc, 'find_claude_cli', lambda: None)
    with pytest.raises(cc.ClaudeCodeError, match='설치돼 있지 않아요'):
        cc.ClaudeCodeAdapter.complete('지시', '질문')
    assert FakeProc.calls == []                                          # 실행 자체를 하지 않음


def test_popen_failure_gives_install_message(cc):
    FakeProc.script = {'popen_error': 'no such file'}
    with pytest.raises(cc.ClaudeCodeError, match='설치돼 있지 않아요'):
        cc.ClaudeCodeAdapter.complete('지시', '질문')


def test_cancel_terminates_and_reports_cancel(cc):
    FakeProc.script = {'hook': lambda: cc.cancel_claude_code()}
    with pytest.raises(cc.ClaudeCodeError, match='취소'):
        cc.ClaudeCodeAdapter.complete('지시', '질문')
    assert FakeProc.calls[-1].terminated


def test_image_is_ignored_and_adapter_is_text_only(cc):
    out = cc.ClaudeCodeAdapter.complete('지시', '질문', image_b64='aGk=')
    assert out == '안녕하세요' and cc.ClaudeCodeAdapter.supports_image is False
    assert cc.ClaudeCodeAdapter.requires_api_key is False


def test_stream_gives_whole_answer_or_friendly_error(cc):
    chunks, done, errors = [], [], []
    cc.ClaudeCodeAdapter.stream_text('장소는?', chunks.append, done.append, errors.append, system='지시')
    assert chunks == ['안녕하세요'] and done == ['안녕하세요'] and errors == []
    FakeProc.script = {'stdout': json.dumps({'is_error': True, 'result': 'usage limit reached'})}
    errors.clear()
    cc.ClaudeCodeAdapter.stream_text('x', lambda t: None, lambda t: None, errors.append)
    assert errors and '사용 한도' in errors[0]


def test_status_reads_auth_status_json(cc, monkeypatch):
    FakeProc.script = {'stdout': json.dumps({'loggedIn': True, 'authMethod': 'claude.ai',
                                              'email': 'a@b.c', 'subscriptionType': 'pro'})}
    st = cc.claude_code_status()
    assert st['installed'] is True and st['logged_in'] is True
    assert st['email'] == 'a@b.c' and st['plan'] == 'pro'
    assert FakeProc.calls[-1].args[1:] == ['auth', 'status', '--json']

    FakeProc.script = {'stdout': json.dumps({'loggedIn': False})}
    assert cc.claude_code_status()['logged_in'] is False

    FakeProc.script = {'stdout': json.dumps({'loggedIn': True, 'authMethod': 'api'})}      # 구독 로그인이 아니면 안 됨
    assert cc.claude_code_status()['logged_in'] is False

    FakeProc.script = {'stdout': 'not json'}
    assert cc.claude_code_status() == {'installed': True, 'logged_in': False, 'email': '', 'plan': '',
                                       'path': FAKE_EXE}


def test_status_when_not_installed(cc, monkeypatch):
    monkeypatch.setattr(cc, 'find_claude_cli', lambda: None)
    assert cc.claude_code_status()['installed'] is False
    assert FakeProc.calls == []


def test_login_and_logout_run_the_official_commands(cc):
    cc.claude_code_login()
    assert FakeProc.calls[-1].args[-3:] == ['auth', 'login', '--claudeai']
    cc.claude_code_logout()
    assert FakeProc.calls[-1].args[-2:] == ['auth', 'logout']
    FakeProc.script = {'returncode': 1, 'stderr': 'denied'}
    with pytest.raises(cc.ClaudeCodeError, match='로그인이 끝나지'):
        cc.claude_code_login()


def test_connection_test_reports_success_with_time(cc):
    ok, msg = cc.test_claude_code()
    assert ok and '연결 성공' in msg and '초' in msg
    assert FakeProc.calls[-1].input == '안녕이라고만 답해'
    FakeProc.script = {'stdout': json.dumps({'is_error': True, 'result': 'Please run /login'})}
    ok, msg = cc.test_claude_code()
    assert not ok and '로그인' in msg


def test_registered_as_a_provider(cc):
    assert 'claude_code' in cc._AI_MODES and cc.PROVIDER_LABELS['claude_code'] == 'Claude 구독'
    cc.save_ai_mode('claude_code')
    assert cc.get_adapter() is cc.ClaudeCodeAdapter
    assert cc.ai_status_text() == 'Claude 구독 · 기본 모델'
    cc.save_claude_code_model('opus')
    assert cc.ai_status_text() == 'Claude 구독 · opus'
    assert cc.load_external_model_name() == 'opus'


def test_text_only_and_privacy_follow_the_setting(cc, monkeypatch):
    import school_profile
    cc.save_ai_mode('claude_code')
    assert cc.text_only_required() is True                              # 이미지는 글자로만
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: False)
    assert cc.masking_required() is False                               # 보호 모드를 따름 (항상 가리지 않음)
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: True)
    assert cc.masking_required() is True


def test_path_setting_is_used_first(cc, tmp_path):
    exe = tmp_path / 'my_claude.exe'
    exe.write_text('x', encoding='utf-8')
    cc.save_claude_code_path(str(exe))
    assert cc.find_claude_cli() == str(exe)
    cc.save_claude_code_path(str(tmp_path / 'nope.exe'))
    assert cc.find_claude_cli() != str(tmp_path / 'nope.exe')           # 없는 경로는 무시
