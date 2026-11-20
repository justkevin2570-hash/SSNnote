"""공문 작성: 작은 시작 창(제목 · 관련 번호 · 종류 · 붙임 · 저장한 공문 · [공문 작성]) → 작성 화면으로 펼치기."""
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
    assert w._mode == 'start' and w.start_page.isVisible() and not w.fold.isVisible()
    assert w.width() <= 720 and w.height() < 260
    assert w.title_input.placeholderText() == '공문 제목'
    assert w.btn_start_go.text() == '공문 작성' and w.btn_past.text() == '저장한 공문'
    assert w.btn_start_capture.isHidden() and w.btn_past.objectName() == 'SavedDocs'   # 캡처 단추 없음 (단축키로)
    assert 'Ctrl' in w.lbl_hint.text() and 'Shift' in w.lbl_hint.text()          # 안내 줄: 캡처 단축키
    assert w.start_page.isAncestorOf(w.title_input) and w.start_page.isAncestorOf(w.ref_input)


def test_enter_with_title_expands_and_drafts(ed):
    w, de, db, app = ed
    w._on_return()                                    # 제목이 비면 펼치지 않음
    assert w._mode == 'start'
    w.title_input.setText('2026학년도 진로 특강 운영 계획')
    width0, top0 = w.width(), w.y()
    w._on_return()
    assert w._mode == 'edit' and w.fold.isVisible() and w.start_page.isVisible()   # 위쪽은 그대로, 아래가 펼쳐짐
    assert '진로 특강' in w.editor.toPlainText()          # 뼈대 초안이 들어감
    from PyQt5.QtTest import QTest
    QTest.qWait(400)
    assert w.width() == width0 and w.height() >= 420      # 같은 가로폭으로 아래로
    assert w.btn_start_go.text() == '다시 초안'


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
    # 제목 · 관련 번호는 칸 바로 옆 [복사] (칸이 비면 숨김), 본문은 위 도구줄 [본문 복사]
    assert w.btn_copy_title.isVisible() and not w.btn_copy_ref.isVisible() and w.btn_copy.isVisible()
    w.ref_input.setText('교무부-1')
    assert w.btn_copy_ref.isVisible()
    w._copy_title()
    assert w.btn_copy_title.text() == '복사됨' and app.clipboard().text() == '제목'
    w._copy_ref()
    assert w.btn_copy_ref.text() == '복사됨' and app.clipboard().text() == '교무부-1'


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


def test_capture_in_start_stays_folded(ed, monkeypatch):
    """접힌 시작 창에서 캡처해도 펼치지 않고, 안내 줄에 진행 상태만. [공문 작성]을 눌러야 펼쳐진다."""
    w, de, db, app = ed
    from PyQt5.QtGui import QPixmap
    monkeypatch.setattr(de, 'load_ai_mode', lambda: 'claude')
    monkeypatch.setattr(w.panel, '_analyze', lambda: None)
    w._overlay = type('O', (), {'region_captured': type('S', (), {'disconnect': lambda self: None})()})()
    monkeypatch.setattr(w, '_restore_after_capture', lambda: None)
    w._on_capture_complete(QPixmap(20, 20))
    assert w._mode == 'start' and not w.fold.isVisible() and '읽는 중' in w.lbl_hint.text()
    w._quick_fill_from_ocr('수신 각 학교장\n제목 2학기 진로교육 운영 계획 제출\n진로교육과-1234 (2026. 11. 10.)')
    assert w.title_input.text() == '2학기 진로교육 운영 계획 제출'           # AI를 기다리지 않고 바로
    assert w.ref_input.text() == '진로교육과-1234(2026. 11. 10.)' and w._mode == 'start'
    w._on_analyzed('2학기 진로교육 운영 계획 제출(정정)', '진로교육과-1234(2026. 11. 10.)')
    assert w.title_input.text() == '2학기 진로교육 운영 계획 제출(정정)'      # AI 결과로 고침
    w._quick_fill_from_ocr('제목 늦게 온 글자 인식')                       # AI가 먼저 끝났으면 덮지 않음
    assert w.title_input.text() == '2학기 진로교육 운영 계획 제출(정정)'
    w.panel._set_status('분석 완료', 'done')
    assert '채웠어요' in w.lbl_hint.text() and w._mode == 'start'
    monkeypatch.setattr(de, 'load_ai_mode', lambda: 'none')      # 초안은 AI 없이 (시험 중 AI 스레드를 띄우지 않게)
    w._start_go()
    assert w._mode == 'edit' and not w.panel.isVisible()            # 받은 공문 요약 칸은 없음


def test_quick_fields():
    import document_editor as de
    assert de._quick_fields('') == ('', '')
    t, n = de._quick_fields('제목: 학부모 상담 주간 운영 안내\n중 등 교 육 과-56 7 8(2026.3.2.)')
    assert t == '학부모 상담 주간 운영 안내' and n == '중등교육과-5678(2026. 3. 2.)'
    assert de._quick_fields('아무 글자')[0] == ''                       # 「제목」 줄이 없으면 제목은 비움
