"""공문 작성 창 디자인 고르기 (⋯ → 디자인): 메모장(기본) · 문서 · 기존 + 단축키 · 상태 줄."""
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
    de.apply_design(de.DEFAULT_DESIGN)


def _menu_texts(w, monkeypatch):
    from PyQt5.QtWidgets import QMenu
    monkeypatch.setattr(QMenu, 'exec_', lambda self, *a: None)
    w._show_more_menu()
    top = [a.text() for a in w._more_menu.actions() if not a.isSeparator()]
    return top, {k: a.text() for k, a in w._design_actions.items()}


def test_default_is_note_and_menu_has_three_designs(ed, monkeypatch):
    w, de, db, app = ed
    assert w.design == 'note' and de.C.NAME == 'note' and de.C.SHOW_KEYS
    assert not hasattr(w, 'btn_help') and w.btn_more.isVisible()          # ? 단추 대신 ⋯ 메뉴
    top, designs = _menu_texts(w, monkeypatch)
    assert top == ['디자인', '사용 방법']
    assert designs['note'].startswith('✓') and '(기본)' in designs['note']
    assert '문서' in designs['paper'] and '기존' in designs['base']
    assert not designs['paper'].startswith('✓')


def test_texts_changed(ed):
    w, de, db, app = ed
    assert w.ref_input.placeholderText() == '관련 번호 (선택)'
    assert w.btn_start_capture.toolTip().startswith('공문 캡처')
    assert '공문 캡처' in w.lbl_hint.text() and '받은 공문' not in w.lbl_hint.text()


def test_switch_keeps_text_and_is_remembered(ed):
    w, de, db, app = ed
    w.title_input.setText('2026학년도 진로 특강 운영 계획')
    w._start_go()
    w.editor.setPlainText('1. 본문입니다.')
    w.set_design('paper')
    assert de.C.NAME == 'paper' and db.get_setting(de.DESIGN_KEY) == 'paper' and de.current_design() == 'paper'
    assert w.title_input.text() == '2026학년도 진로 특강 운영 계획' and w.editor.toPlainText() == '1. 본문입니다.'
    assert w._mode == 'edit' and w.editor.font().family() == '바탕체' and not w.editor.rule_color
    assert w.lbl_ask.text() == '고칠 내용' and not w.lbl_ask.isHidden()
    w.set_design('base')
    assert de.C.PRIMARY == '#4F46E5' and not de.C.SHOW_KEYS and w.btn_save_exit._kbd.isHidden()
    assert w.lbl_ask.isHidden() and not w._spark.isHidden()            # 기존: 말풍선 아이콘
    w.set_design('note')
    assert w.editor.rule_color and w.editor.font().family() != '바탕체' and not w.btn_save_exit._kbd.isHidden()


def test_window_width_stays_when_expanding(ed):
    w, de, db, app = ed
    from PyQt5.QtTest import QTest
    for name in ('note', 'paper', 'base'):
        w.set_design(name)
        w.title_input.setText('2026학년도 진로 특강 운영 계획')
        width0 = w.width()
        w._start_go()
        QTest.qWait(350)
        assert w.width() == width0, name                 # 펼쳐도 가로폭은 그대로 (도구줄이 늘 들어감)
        w.editor.clear()
        w._enter_start()


def test_shortcuts(ed, monkeypatch):
    w, de, db, app = ed
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    w.activateWindow()
    w.title_input.setText('제목')
    w.ref_input.setText('교무부-7')
    w._start_go()
    w.editor.setPlainText('본문 글')
    w.editor.setFocus()
    app.processEvents()
    keys = {sc.key().toString(): sc for sc in w._shortcuts}
    assert set(keys) == {'Ctrl+1', 'Ctrl+2', 'Ctrl+3', 'Ctrl+N', 'Ctrl+O', 'Ctrl+R', 'Ctrl+L', 'Ctrl+S'}
    for key, want in (('Ctrl+1', '제목'), ('Ctrl+2', '교무부-7'), ('Ctrl+3', '본문 글')):
        keys[key].activated.emit()
        assert app.clipboard().text() == want
    keys['Ctrl+L'].activated.emit()                   # AI가 꺼져 있으면 요청 칸이 막혀 있어 그대로
    saved = []
    monkeypatch.setattr(w, 'save_and_close', lambda: saved.append(1))
    keys['Ctrl+S'].activated.emit()
    assert saved == [1]
    QTest.keyClick(w.editor, Qt.Key_1, Qt.ControlModifier)   # 본문에 커서가 있어도 단축키가 먹힘
    assert app.clipboard().text() == '제목'


def test_status_count(ed):
    w, de, db, app = ed
    w.editor.setPlainText('가나다\n라마')
    assert w._lbl_count.text() == '본문 2줄 · 5자'
    w.editor.clear()
    assert w._lbl_count.text() == '본문 0줄 · 0자'


def _ratio(a, b):
    def lum(h):
        r, g, b_ = (int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))
        f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b_)
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


@pytest.mark.parametrize('name', ['note', 'paper'])
def test_text_contrast(name):
    import document_editor as de
    de.apply_design(name)
    C = de.C
    try:
        for fg, bg in ((C.TEXT, C.BG), (C.MUTED, C.BG), (C.MUTED, C.HEADER), (C.TEXT, C.HEADER),
                       (C.BODY, C.SURFACE), ('#FFFFFF', C.PRIMARY), (C.SAVED_TEXT, C.SAVED_BG),
                       (C.KBD_TEXT, C.KBD_BG), (C.ACCENT, C.SURFACE)):
            assert _ratio(fg, bg) >= 4.5, (name, fg, bg)
    finally:
        de.apply_design(de.DEFAULT_DESIGN)
