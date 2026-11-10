"""공문 선택 부분 다듬기 — 비교·개인정보 되살리기 (doc_polish), 창에서 적용·되돌리기 (document_editor)."""
import pytest

from doc_polish import (POLISH_SYSTEM, build_polish_prompt, diff_spans, leftover_masks, restore_masked)

pytest.importorskip('PyQt5')

_APP = None


def test_diff_marks_only_changed_words():
    left, right = diff_spans('가 나 다 라', '가 나 X 라')
    assert [('가 나 다 라'[s:e]) for s, e in left] == ['다']
    assert [('가 나 X 라'[s:e]) for s, e in right] == ['X']


def test_diff_ignores_identical_text():
    assert diff_spans('본문 그대로', '본문 그대로') == ([], [])


def test_restore_puts_back_masked_values_only():
    original = '연락처 010-1234-5678 입니다.'
    masked = '연락처 010-****-5678 입니다.'
    polished = '연락처 010-****-5678 임.'
    text, ok = restore_masked(polished, original, masked)
    assert ok and text == '연락처 010-1234-5678 임.'


def test_restore_refuses_when_word_count_differs():
    assert restore_masked('x', '가 나 다', '가 나') == ('x', False)


def test_leftover_mask_detection():
    assert leftover_masks('김○○ 학생') and leftover_masks('010-****-1234')
    assert not leftover_masks('정상 글입니다.')


def test_prompt_keeps_rules_notes_and_examples():
    assert '○○' in POLISH_SYSTEM and '마크다운' in POLISH_SYSTEM and '공문 문체' in POLISH_SYSTEM
    p = build_polish_prompt('교무회의를 개최함', '[작성 규칙] 시간은 24시간제', '[예시 1] 제목: 회의')
    assert '[작성 규칙]' in p and '[예시 1]' in p and '[다듬을 글]\n교무회의를 개최함' in p
    assert '다듬은 글만 출력' in p


# ── 창 동작 ───────────────────────────────────────────────────────

class _Signal:
    def __init__(self):
        self._subs = []

    def connect(self, fn):
        self._subs.append(fn)

    def emit(self, *args):
        for fn in self._subs:
            fn(*args)


@pytest.fixture
def win(tmp_path, monkeypatch):
    global _APP
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    import db
    import ai_client as ai
    import document_editor as de
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setattr(ai, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    ai.save_ai_mode('claude_code')
    monkeypatch.setattr(de.DocumentEditorWindow, '_STATE_FILE', str(tmp_path / 'state.json'))
    w = de.DocumentEditorWindow()
    yield w
    w._polish = None
    w.close()


class _FakeAI:
    """AiStreamThread 대신 쓰는 가짜: start()에서 바로 답을 보낸다. reply=None이면 답하지 않는다."""
    reply = ''
    seen = []

    def __init__(self, prompt, system=None, structured=False):
        _FakeAI.seen.append((prompt, system, structured))
        self.finished_signal = _Signal()
        self.error_signal = _Signal()
        self.new_text_signal = _Signal()

    def start(self):
        if _FakeAI.reply is not None:
            self.finished_signal.emit(_FakeAI.reply)

    def isRunning(self):
        return False

    def stop(self):
        pass


@pytest.fixture
def fake_ai(monkeypatch):
    import document_editor as de
    _FakeAI.reply, _FakeAI.seen = '', []
    monkeypatch.setattr(de, 'AiStreamThread', _FakeAI)
    accept = {'ok': True}
    monkeypatch.setattr(de.PolishDialog, 'exec_', lambda self: de.QDialog.Accepted if accept['ok'] else de.QDialog.Rejected)
    return accept


def _select(win, text):
    doc = win.editor.toPlainText()
    start = doc.index(text)
    from PyQt5.QtGui import QTextCursor
    cur = win.editor.textCursor()
    cur.setPosition(start)
    cur.setPosition(start + len(text), QTextCursor.KeepAnchor)
    win.editor.setTextCursor(cur)
    return start


def test_apply_changes_only_the_selection_and_undo_once(win):
    win.editor.setPlainText('첫 줄\n교무회의를 개최함\n끝.')
    start = win.editor.toPlainText().index('교무회의를 개최함')
    assert win._apply_polish(start, start + len('교무회의를 개최함'), '교무회의를 개최함', '교무회의를 개최하고자 함')
    assert win.editor.toPlainText() == '첫 줄\n교무회의를 개최하고자 함\n끝.'
    win.editor.undo()                                                # Ctrl+Z 한 번이면 원문
    assert win.editor.toPlainText() == '첫 줄\n교무회의를 개최함\n끝.'


def test_apply_refuses_when_the_text_moved_on(win):
    win.editor.setPlainText('가 나 다')
    assert win._apply_polish(0, 3, '다른 글', '새 글') is False
    assert win.editor.toPlainText() == '가 나 다'


def test_undo_link_restores_the_original(win):
    win.editor.setPlainText('첫 줄\n교무회의를 개최함\n끝.')
    old = '교무회의를 개최함'
    s = win.editor.toPlainText().index(old)
    win._apply_polish(s, s + len(old), old, '교무회의를 개최하고자 함')
    assert not win._polish_link.isHidden()
    win._undo_polish()
    assert win.editor.toPlainText() == '첫 줄\n교무회의를 개최함\n끝.' and win._polish_link.isHidden()


def test_polish_selection_sends_only_the_selection_and_applies(win, fake_ai, monkeypatch):
    win.editor.setPlainText('첫 줄\n교무회의를 개최함\n끝.')
    _select(win, '교무회의를 개최함')
    _FakeAI.reply = '교무회의를 개최하고자 함'
    win.polish_selection()
    prompt, system, structured = _FakeAI.seen[-1]
    assert '교무회의를 개최함' in prompt and '첫 줄' not in prompt
    assert system == POLISH_SYSTEM and structured is False
    assert win.editor.toPlainText() == '첫 줄\n교무회의를 개최하고자 함\n끝.'


def test_no_selection_polishes_whole_body(win, fake_ai):
    win.editor.setPlainText('본문 전체 글입니다.')
    _FakeAI.reply = '본문 전체 글임.'
    win.polish_selection()
    assert win.editor.toPlainText() == '본문 전체 글임.'


def test_rejected_in_dialog_keeps_body(win, fake_ai):
    win.editor.setPlainText('원래 글입니다.')
    _FakeAI.reply = '고친 글임.'
    fake_ai['ok'] = False
    win.polish_selection()
    assert win.editor.toPlainText() == '원래 글입니다.' and not win.editor.isReadOnly()


def test_privacy_mask_is_sent_hidden_and_restored(win, fake_ai):
    from school_profile import mask_for_ai
    text = '연락처 010-1234-5678 입니다.'
    win.editor.setPlainText(text)
    masked, _ = mask_for_ai(text)
    _FakeAI.reply = masked.replace('입니다', '임')
    win.polish_selection()
    prompt = _FakeAI.seen[-1][0]
    assert '1234' not in prompt and '010-****-5678' in prompt        # 보낼 때는 가려짐
    assert win.editor.toPlainText() == '연락처 010-1234-5678 임.'      # 적용할 때 원래 값으로


def test_cancel_during_polish_keeps_body_and_editable(win):
    win.editor.setPlainText('그대로 둘 글')
    win._polish = {'start': 0, 'end': 6, 'original': '그대로 둘 글', 'masked': '그대로 둘 글', 'counts': {}}
    win._set_polish_busy(True)
    assert win.editor.isReadOnly()
    win._cancel_polish()
    assert win.editor.toPlainText() == '그대로 둘 글' and not win.editor.isReadOnly() and win._polish is None


def test_ai_off_does_nothing(win, fake_ai, monkeypatch):
    import ai_client as ai
    ai.save_ai_mode('none')
    win.editor.setPlainText('글')
    win.polish_selection()
    assert _FakeAI.seen == [] and '꺼져' in win.statusBar().currentMessage()


def test_dialog_defaults_to_apply_and_marks_changes(win):
    import document_editor as de
    dlg = de.PolishDialog('가 나 다', '가 나 라')
    assert dlg._btn_apply.isDefault() and dlg.result_text() == '가 나 라'
    assert len(dlg._left.extraSelections()) == 1 and len(dlg._right.extraSelections()) == 1
    dlg.close()
