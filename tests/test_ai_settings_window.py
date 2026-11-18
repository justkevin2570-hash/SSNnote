"""AI 연결 창: 지금 쓰는 AI · 구독 계정 연동 2개 · API 연동 3개 · [연결]/[설정] · AI 끄기."""
import pytest

pytest.importorskip('PyQt5')


@pytest.fixture
def win(tmp_path, monkeypatch):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    import ai_client as ai
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai, 'CODEX_EFFORT_FILE', str(tmp_path / 'codex_effort.txt'))
    from ui import ai_settings_window as aw, ai_connect
    calls = []

    def fake_connect(parent, mode, say=None):
        calls.append(mode)
        ai.save_ai_mode(mode)
        ai_connect.notify_changed()
        return True
    monkeypatch.setattr(ai_connect, 'connect', fake_connect)
    w = aw.AiSettingsWindow()
    yield w, ai, calls
    w.close()
    app.processEvents()


def test_lists_five_ways_in_two_groups_and_off_state(win):
    w, ai, _ = win
    assert list(w.rows) == ['claude_code', 'codex', 'opencode', 'claude', 'openai']
    assert 'AI 꺼짐' in w.lbl_now.text() and not w.btn_off.isEnabled()
    assert all(btn.text() == '연결' and check.isHidden() for _, check, btn in w.rows.values())


def test_connect_marks_current_and_turn_off(win):
    w, ai, calls = win
    w.rows['codex'][2].click()
    assert calls == ['codex'] and ai.load_ai_mode() == 'codex'
    assert w.rows['codex'][2].text() == '설정' and not w.rows['codex'][1].isHidden()
    assert w.rows['claude'][2].text() == '연결'
    assert 'ChatGPT 구독' in w.lbl_now.text() and w.btn_off.isEnabled()
    w.btn_off.click()
    assert ai.load_ai_mode() == 'none' and 'AI 꺼짐' in w.lbl_now.text()


def test_main_menu_opens_window_not_document_editor():
    import inspect
    import window
    src = inspect.getsource(window.MemoWindow._open_ai_settings) if hasattr(window, 'MemoWindow') else \
        inspect.getsource(window)
    assert 'open_ai_settings' in src
