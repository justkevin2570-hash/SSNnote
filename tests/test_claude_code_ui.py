"""Claude 구독 설정 창 (ClaudeCodeDialog) — 가짜 claude 함수로 상태 표시·설치·로그인·저장 확인."""
import pytest

pytest.importorskip('PyQt5')

_APP = None


@pytest.fixture
def ui(tmp_path, monkeypatch):
    global _APP
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    import ai_client as ai
    import document_editor as de
    monkeypatch.setattr(ai, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai, 'CLAUDE_CODE_PATH_FILE', str(tmp_path / 'claude_code_path.txt'))
    monkeypatch.setattr(ai, 'CLAUDE_CODE_MODEL_FILE', str(tmp_path / 'claude_code_model.txt'))
    state = {'installed': True, 'logged_in': True, 'email': 'me@example.com', 'plan': 'pro', 'path': 'x'}
    calls = []
    monkeypatch.setattr(de, 'claude_code_status', lambda: dict(state))
    monkeypatch.setattr(de, 'install_claude_code', lambda: calls.append('install') or None)
    monkeypatch.setattr(de, 'claude_code_login', lambda: calls.append('login') or None)
    monkeypatch.setattr(de, 'claude_code_logout', lambda: calls.append('logout') or None)
    monkeypatch.setattr(de, 'test_claude_code', lambda model: (True, '연결 성공 (2.0초)'))
    warned = []
    monkeypatch.setattr(de.QMessageBox, 'warning', lambda *a, **k: warned.append(a[2]))
    monkeypatch.setattr(de.QMessageBox, 'question', lambda *a, **k: de.QMessageBox.Yes)
    yield {'de': de, 'ai': ai, 'state': state, 'calls': calls, 'warned': warned, 'app': _APP}


def _settle(dlg):
    """뒤 스레드 작업이 끝나고 신호가 처리될 때까지 기다린다."""
    for _ in range(20):
        if dlg._job is not None:
            dlg._job.wait(2000)
        _APP.processEvents()


def _open(ui):
    dlg = ui['de'].ClaudeCodeDialog()
    _settle(dlg)
    return dlg


def test_logged_in_shows_plan_and_logout_button(ui):
    dlg = _open(ui)
    assert '✓ 로그인됨' in dlg.lbl_status.text() and 'me@example.com' in dlg.lbl_status.text()
    assert 'Pro 구독' in dlg.lbl_status.text()
    assert dlg.btn_account.text() == '로그아웃' and not dlg.btn_account.isHidden()
    assert dlg.btn_install.isHidden()


def test_not_installed_shows_install_and_link(ui):
    ui['state'].update(installed=False, logged_in=False)
    dlg = _open(ui)
    assert '설치 안 됨' in dlg.lbl_status.text()
    assert not dlg.btn_install.isHidden() and dlg.btn_account.isHidden()
    assert 'code.claude.com/docs' in dlg.lbl_link.text()


def test_not_logged_in_shows_login_button(ui):
    ui['state'].update(logged_in=False)
    dlg = _open(ui)
    assert '로그인 필요' in dlg.lbl_status.text()
    assert dlg.btn_account.text() == 'Claude로 로그인' and dlg.btn_install.isHidden()


def test_install_runs_then_refreshes(ui):
    ui['state'].update(installed=False, logged_in=False)
    dlg = _open(ui)
    dlg._install()
    _settle(dlg)
    assert ui['calls'] == ['install']
    assert dlg.lbl_progress.text() == '설치했어요.'


def test_login_button_runs_login(ui):
    ui['state'].update(logged_in=False)
    dlg = _open(ui)
    dlg._login_or_logout()
    _settle(dlg)
    assert ui['calls'] == ['login']


def test_logout_asks_then_runs(ui):
    dlg = _open(ui)
    dlg._login_or_logout()
    _settle(dlg)
    assert ui['calls'] == ['logout']


def test_connection_test_shows_result(ui):
    dlg = _open(ui)
    dlg._test()
    _settle(dlg)
    assert '연결 성공' in dlg.lbl_test.text() and '2.0초' in dlg.lbl_test.text()


def test_save_is_refused_until_logged_in(ui):
    ui['state'].update(logged_in=False)
    dlg = _open(ui)
    dlg._save()
    assert ui['warned'] and '로그인' in ui['warned'][0]
    assert dlg.result() != dlg.Accepted
    assert ui['ai'].load_ai_mode() == 'none'


def test_save_stores_mode_model_and_path(ui, tmp_path):
    ai = ui['ai']
    exe = tmp_path / 'claude.exe'
    exe.write_text('x', encoding='utf-8')
    dlg = _open(ui)
    dlg.cmb_model.setCurrentIndex(dlg.cmb_model.findData('opus'))
    dlg.edit_path.setText(str(exe))
    dlg._save()
    assert dlg.result() == dlg.Accepted
    assert ai.load_ai_mode() == 'claude_code'
    assert ai.load_claude_code_model() == 'opus'
    assert ai.load_claude_code_path() == str(exe)


def test_save_refuses_missing_path(ui, tmp_path):
    dlg = _open(ui)
    dlg.edit_path.setText(str(tmp_path / 'nope.exe'))
    dlg._save()
    assert ui['warned'] and '경로' in ui['warned'][0]
    assert ui['ai'].load_ai_mode() == 'none'


def test_settings_menu_offers_claude_code(ui):
    de = ui['de']
    assert de._PROVIDERS['claude_code']['desc'] == '유료 구독 필요'
    assert 'claude_code' in de._PROVIDERS and de.PROVIDER_LABELS['claude_code'] == 'Claude 구독'
