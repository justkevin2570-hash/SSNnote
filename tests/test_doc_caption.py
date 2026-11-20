"""공문 작성 창: Windows 제목표시줄 숨기기 — 머리줄 ─ ✕ 단추, 머리줄 빈 곳 = 제목표시줄, 가장자리 = 크기 조절."""
import sys

import pytest

pytest.importorskip('PyQt5')
_APP = None


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
    global _APP
    _APP = QApplication.instance() or QApplication([])
    import document_editor as de
    monkeypatch.setattr(de, 'load_ai_mode', lambda: 'none')
    monkeypatch.setattr(de.DocumentEditorWindow, '_STATE_FILE', str(tmp_path / 'st.json'))
    w = de.DocumentEditorWindow()
    w.show()
    _APP.processEvents()
    yield w, de
    w.close()


def test_window_buttons_only_when_caption_hidden(ed):
    w, de = ed
    assert w.btn_win_min.toolTip() == '최소화' and w.btn_win_close.toolTip() == '닫기'
    assert w.btn_win_close.isVisible() == de._CAPTION_HIDDEN      # Windows에서만 (여기서는 기본 제목표시줄)


def test_caption_point_is_empty_header_only(ed):
    w, de = ed
    from PyQt5.QtCore import QPoint
    hdr = w._header
    empty = hdr.mapTo(w, QPoint(hdr.width() // 2, hdr.height() // 2))
    assert w._is_caption_point(empty)
    btn = w.btn_more.mapTo(w, w.btn_more.rect().center())
    assert not w._is_caption_point(btn)                            # 단추 위는 단추
    field = w.title_input.mapTo(w, w.title_input.rect().center())
    assert not w._is_caption_point(field)                          # 머리줄 밖


def test_hit_test_edges_and_fixed_height(ed):
    w, de = ed
    from PyQt5.QtCore import QPoint
    w.setMinimumHeight(100)
    w.setMaximumHeight(16777215)
    w.resize(600, 400)
    _APP.processEvents()
    H = de._HT
    assert de._hit_test(w, QPoint(2, 200)) == H['left']
    assert de._hit_test(w, QPoint(w.width() - 2, w.height() - 2)) == H['bottomright']
    assert de._hit_test(w, QPoint(300, 2)) == H['top']
    assert de._hit_test(w, QPoint(300, 2), zoomed=True) != H['top']   # 최대화면 크기 조절 없음
    w.setFixedHeight(220)                                              # 접힌 시작 창: 높이 고정
    _APP.processEvents()
    assert de._hit_test(w, QPoint(300, 2)) != H['top'] and de._hit_test(w, QPoint(2, 100)) == H['left']
    hdr = w._header
    assert de._hit_test(w, hdr.mapTo(w, QPoint(hdr.width() // 2, hdr.height() // 2))) == de._HTCAPTION
