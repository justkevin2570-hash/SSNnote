"""공문 작성: 작은 시작 창(제목 · 관련 번호 · 종류 · 붙임 · 캡처 · 지난 공문 · [공문 작성]) → 작성 화면으로 펼치기."""
import sys

import pytest

pytest.importorskip('PyQt5')


@pytest.fixture
def ed(tmp_path, monkeypatch):
    from unittest.mock import MagicMock
    sys.modules.setdefault('winreg', MagicMock())
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setenv('APPDATA', str(tmp_path))
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    import ai_client
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    import document_editor as de
    monkeypatch.setattr(de, 'load_ai_mode', lambda: 'none')
    monkeypatch.setattr(de.DocumentEditorWindow, '_STATE_FILE', str(tmp_path / 'st.json'))
    w = de.DocumentEditorWindow()
    w.show()
    app.processEvents()
    yield w, de, db, app
    w.editor.clear()
    w.close()


def test_opens_as_small_start_window(ed):
    w, de, db, app = ed
    assert w._mode == 'start' and w.start_page.isVisible() and not w.splitter.isVisible()
    assert w.width() <= 480 and w.height() < 200
    assert w.title_input.placeholderText() == '공문 제목'
    assert w.btn_start_go.text() == '공문 작성' and w.btn_past.text() == '지난 공문'
    assert w.btn_start_capture.text() == '' and w.btn_start_capture.toolTip() == '받은 공문 캡처'
    assert w.title_input.parent() is w.start_page


def test_enter_with_title_expands_and_drafts(ed):
    w, de, db, app = ed
    w._on_return()                                    # 제목이 비면 펼치지 않음
    assert w._mode == 'start'
    w.title_input.setText('2026학년도 진로 특강 운영 계획')
    w._on_return()
    assert w._mode == 'edit' and w.splitter.isVisible() and not w.start_page.isVisible()
    assert '진로 특강' in w.editor.toPlainText()          # 뼈대 초안이 들어감
    assert w.width() >= 700


def test_new_document_returns_to_start(ed, monkeypatch):
    w, de, db, app = ed
    w.title_input.setText('제목')
    w._start_go()
    w.editor.setPlainText('직접 쓴 글')
    from PyQt5.QtWidgets import QMessageBox
    monkeypatch.setattr(QMessageBox, 'exec_', lambda self: 0)
    monkeypatch.setattr(QMessageBox, 'clickedButton', lambda self: next(b for b in self.buttons() if b.text() == '저장 안 함'))
    w._new_document()
    assert w._mode == 'start' and not w.title_input.text() and not w.editor.toPlainText()


def test_attach_button_and_copy_buttons(ed):
    w, de, db, app = ed
    w.spin_attach.setValue(0)
    assert w.btn_attach.text().startswith('붙임 없음')
    w.spin_attach.setValue(2)
    assert w.btn_attach.text().startswith('붙임 2')
    w.title_input.setText('제목')
    w._start_go()
    assert not w.btn_copy_ref.isVisible() and w.btn_copy_title.isVisible()
    w._copy_title()
    assert w.btn_copy_title.text() == '복사됨' and app.clipboard().text() == '제목'


def test_past_documents_load_with_year_bump(ed, monkeypatch):
    w, de, db, app = ed
    db.save_official_document('2020학년도 진로 특강 계획', '진로과-1', '1. 2020학년도 진로 특강을 합니다.', 'general')
    dlg = de.PastDocumentsDialog()
    assert dlg.list.count() == 1
    dlg.search.setText('없는 제목')
    assert dlg.list.count() == 0
    dlg.close()

    class Fake:
        def __init__(self, parent=None):
            self.chosen = db.get_official_documents()[0]

        def exec_(self):
            return de.QDialog.Accepted
    monkeypatch.setattr(de, 'PastDocumentsDialog', Fake)
    from PyQt5.QtWidgets import QMessageBox
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: QMessageBox.Yes)
    w._open_past_documents()
    old, new = de._school_year_bump('2020학년도')
    assert w._mode == 'edit' and f'{new}학년도' in w.title_input.text()
    assert f'{new}학년도 진로 특강' in w.editor.toPlainText() and w.ref_input.text() == '진로과-1'


def test_school_year_bump():
    import document_editor as de
    assert de._school_year_bump('글') == (None, None)
    assert de._school_year_bump('2099학년도') == (None, None)
    old, new = de._school_year_bump('2019학년도 · 2020학년도')
    assert old == 2020 and new > 2020
