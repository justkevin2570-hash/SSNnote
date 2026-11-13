"""메모장(메모 전용 창): 「…」 메뉴는 3가지뿐, 키보드 단축키 없음. 메인 창은 그대로."""
import os
import sys
import unittest.mock as um

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

_APP = None


@pytest.fixture
def make(tmp_path, monkeypatch):
    import db
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())      # 자동실행(레지스트리) 모듈은 Windows 전용
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication, QMenu
    global _APP
    _APP = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMenu, 'exec_', lambda self, *a: None)       # 메뉴를 띄우지 않고 내용만 본다
    import window
    made = []

    def _make(memo_only):
        w = window.MemoWindow(db.get_first_window_id(), force_memo_mode=memo_only)
        made.append(w)
        return w
    yield _make
    for w in made:
        w.close()


def _texts(menu):
    return [a.text() for a in menu.actions() if not a.isSeparator()]


def test_memo_window_menu_has_only_two_items(make):
    w = make(True)
    w.show_menu(w.title_bar.btn_menu)
    items = _texts(w._last_menu)
    assert items == ['배경색·글자·줄간격', '메모장 삭제']          # 다크 모드는 없앰
    style = next(a for a in w._last_menu.actions() if a.text() == '배경색·글자·줄간격').menu()
    assert '글자 크기' in _texts(style) and '줄 간격' in _texts(style)     # 메모장은 글 줄 높이
    assert not any('개인정보' in t or '다크' in t for t in items)


def test_no_dark_mode_anywhere(make):
    from ui import theme
    assert theme._load_mode() == 'light' and not theme.DARK
    w = make(False)
    w.show_menu(w.title_bar.btn_menu)
    settings = next(a for a in w._last_menu.actions() if a.text() == '설정').menu()
    assert not any('다크' in t for t in _texts(settings))
    assert not any('다크' in c.title for c in w._palette_commands())


def test_memo_window_has_no_shortcuts_and_no_palette(make):
    w = make(True)
    assert w._shortcuts == []
    w._open_command_palette()                      # 어떤 길로 불려도 열리지 않음
    assert getattr(w, '_palette', None) is None


def test_main_window_keeps_full_menu_and_shortcuts(make):
    w = make(False)
    assert len(w._shortcuts) == 3                  # 핀 · 접기 · 빠른 찾기
    w.show_menu(w.title_bar.btn_menu)
    items = _texts(w._last_menu)
    assert '설정' in items and '도움말' in items and '배경색·글자·줄간격' in items and len(items) > 5
    style = next(a for a in w._last_menu.actions() if a.text() == '배경색·글자·줄간격').menu()
    assert '줄 간격' in _texts(style)


def test_memo_font_is_pretendard_and_stays_after_size_change(make):
    from ui.common import pr_font
    w = make(True)
    want = pr_font(12).family()
    assert 'Pretendard' in want
    assert w.memo_editor.font().family() == want and w.memo_editor.font().pointSize() == 12
    for scale, pt in ((1.1, 13), (1.2, 14), (1.0, 12)):         # 글자 크기를 바꿔도 글꼴은 그대로
        w._set_scale(scale)
        assert w.memo_editor.font().family() == want and w.memo_editor.font().pointSize() == pt



def test_menu_words_match_app(make):
    """Claude 연결은 메뉴에서 뺐고, 업무창 메뉴는 「업무창」이라 부른다."""
    import db
    import window
    make(False)
    w = window.MemoWindow(db.create_window(), open_windows=[])     # 둘째 업무창: 삭제 메뉴가 있음
    try:
        w.show_menu(w.title_bar.btn_menu)
        items = _texts(w._last_menu)
        settings = next(a for a in w._last_menu.actions() if a.text() == '설정').menu()
        assert 'AI 연결' in _texts(settings)
        assert not any('Claude' in t for t in _texts(settings))
        assert '업무창 삭제' in items and '메모장 삭제' not in items
        titles = [c.title for c in w._palette_commands()]
        assert '↕️ 업무창 접기/펼치기' in titles and '🔑 AI 연결' in titles
        assert not any('Claude 연결' in t for t in titles)
    finally:
        w.close()


def test_bottom_buttons_same_height_and_row(make):
    """아래 줄 달력 · 공문 · 예산 · AI 버튼은 높이와 위쪽 끝이 같다."""
    from PyQt5.QtWidgets import QApplication, QPushButton
    w = make(False)
    w.resize(340, 420)
    w.show()
    QApplication.processEvents()
    bar = w.btn_ai_bottom.parentWidget()
    btns = [b for b in bar.findChildren(QPushButton) if b.isVisible()]
    assert len(btns) == 4
    assert len({b.height() for b in btns}) == 1 and len({b.y() for b in btns}) == 1
    top = btns[0].y()
    assert bar.height() - (top + btns[0].height()) == top          # 위아래 여백 같게
