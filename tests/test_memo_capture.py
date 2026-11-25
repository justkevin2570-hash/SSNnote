"""업무 메모 [스크린샷]: 공문 캡처와 같은 방법 — PC 글자 인식으로 먼저, AI가 있으면 AI 값으로."""
import sys
import types

import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def ed(tmp_path, monkeypatch):
    from unittest.mock import MagicMock
    sys.modules.setdefault('winreg', MagicMock())
    monkeypatch.setitem(sys.modules, 'winrt', types.ModuleType('winrt'))
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    monkeypatch.setenv('APPDATA', str(tmp_path))
    db.init_db()
    global _APP
    from PyQt5.QtWidgets import QApplication, QMessageBox
    _APP = QApplication.instance() or QApplication([])
    from ui import task_row
    import doc_ai
    import ai_client
    import school_profile
    st = {'ocr': '', 'ai': None, 'mode': 'none', 'privacy': False, 'msgs': [], 'sent': []}
    monkeypatch.setattr(task_row, 'run_ocr', lambda pm, cb, err=None: cb(st['ocr']))
    monkeypatch.setattr(doc_ai.DocAnalysisWorker, 'start', lambda self: self.run())

    class Ad:
        supports_image = True

        @staticmethod
        def get_api_key():
            return 'x'

        @staticmethod
        def complete(system, prompt, image_b64=None):
            st['sent'].append('image' if image_b64 else 'text')
            return '{"doc_number": "%s"}' % (st['ai'] or '')
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: st['mode'])
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: None if st['mode'] == 'none' else Ad)
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: st['privacy'])
    monkeypatch.setattr(QMessageBox, 'information', lambda *a: st['msgs'].append(a[2]))
    w = task_row._InlineNoteEdit()
    w.task_id = 1
    monkeypatch.setattr(w, '_restore_inline_window', lambda: None)
    return w, st


def _shot(w):
    from PyQt5.QtGui import QPixmap
    pm = QPixmap(200, 30)
    pm.fill()
    w._on_inline_capture_complete(pm)


def test_pc_only_drops_label(ed):
    w, st = ed
    st['ocr'] = '2027학년도 운영 계획\n관련 교육과정과-12873(2026. 11. 18.)'
    _shot(w)
    assert w.ref_edit.text() == '교육과정과-12873(2026. 11. 18.)' and st['msgs'] == []


def test_ai_value_wins(ed):
    w, st = ed
    st.update(mode='claude_code', ocr='교육과정과-12B73 (2026. 11. 18.)', ai='교육과정과-12873')
    _shot(w)
    assert w.ref_edit.text().startswith('교육과정과-12873') and st['sent'] == ['image']


def test_privacy_sends_text_when_personal_info(ed):
    w, st = ed
    st.update(mode='claude_code', privacy=True, ocr='3반 홍길동 학생 010-1234-5678\n교육과정과-12873 (2026. 11. 18.)',
              ai='교육과정과-12873')
    _shot(w)
    assert st['sent'] == ['text'] and w.ref_edit.text().startswith('교육과정과-12873')


def test_user_edit_kept_and_not_found_message(ed):
    w, st = ed
    w.ref_edit.setText('직접 적은 번호')
    st['ocr'] = '교육과정과-12873 (2026. 11. 18.)'
    _shot(w)
    assert w.ref_edit.text() == '직접 적은 번호'
    w.ref_edit.setText('')
    st['ocr'] = '아무 번호 없음'
    _shot(w)
    assert w.ref_edit.text() == '' and st['msgs'] and '찾지 못했어요' in st['msgs'][-1]
