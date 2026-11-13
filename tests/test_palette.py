"""배경색 새 팔레트 (v3.06): 예전 색 옮기기 · AI 색 이름 · 견본."""
import os
import sys
import unittest.mock as um

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.modules.setdefault('winreg', um.MagicMock())

from ui import theme


def test_palette_has_names_and_old_colors_map_into_it():
    assert len(theme.PALETTE) == 11 and theme.DEFAULT_NOTE_COLOR in theme.PALETTE
    assert all(new in theme.PALETTE for new in theme.OLD_COLOR_MAP.values())
    assert '#AAAAAA' not in theme.PALETTE                  # 어두운 회색은 뺌


def test_init_db_moves_old_colors(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 's.db'))
    db.init_db()
    a = db.get_first_window_id()
    b = db.create_window()
    c = db.create_window()
    with db._connect() as conn:
        conn.execute("UPDATE windows SET color='#ccff90' WHERE id=?", (a,))
        conn.execute("UPDATE windows SET color='#AAAAAA' WHERE id=?", (b,))
        conn.execute("UPDATE windows SET color='' WHERE id=?", (c,))
    db.set_setting('ai_chat_color', '#AECBFA')
    db.init_db()
    colors = {w['id']: w['color'] for w in db.get_all_windows()}
    assert colors[a] == '#E9F5D3' and colors[b] == '#EDEDEB' and colors[c] == ''
    assert db.get_setting('ai_chat_color', '') == '#DFECFC'
    db.init_db()                                           # 두 번 돌려도 그대로
    assert {w['id']: w['color'] for w in db.get_all_windows()} == colors


def test_ai_color_names_and_old_aliases():
    import ai_actions
    assert ai_actions._color('하늘') == '#DFECFC'
    assert ai_actions._color('보라') == '#EBE5FA'          # 예전 이름도 알아듣기
    assert ai_actions._color('진노랑') == '#FFEAC0'
    assert ai_actions._color('#FDCFE8') == '#FDE2EB'        # 예전 색 코드
    assert ai_actions._color('복숭아') == '#FFE4D8'


def test_swatches_mark_current_color():
    from PyQt5.QtWidgets import QApplication, QMenu
    QApplication.instance() or QApplication([])
    from ui import style_menu
    w = style_menu.color_swatches(QMenu(), '#DFECFC', lambda c: None)
    sw = [s for s in w.findChildren(style_menu.Swatch)]
    assert len(sw) == 11 and [s.color for s in sw if s.chosen] == ['#DFECFC']
    assert sw[0].toolTip() == '노랑'
    w2 = style_menu.color_swatches(QMenu(), '=', lambda c: None, same_as_note=True)
    assert [s.color for s in w2.findChildren(style_menu.Swatch) if s.chosen] == ['=']


def test_v306_palette_moves_to_brighter_same_name(tmp_path, monkeypatch):
    """v3.06~v3.10 색은 같은 이름의 더 밝은 새 색으로 (예: 노랑 #FFF5A5 → #FFF8C2)."""
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 's.db'))
    db.init_db()
    a = db.get_first_window_id()
    with db._connect() as conn:
        conn.execute("UPDATE windows SET color='#FFF5A5' WHERE id=?", (a,))
    db.set_setting('ai_chat_color', '#d2e4fa')
    db.init_db()
    assert {w['id']: w['color'] for w in db.get_all_windows()}[a] == '#FFF8C2' == theme.DEFAULT_NOTE_COLOR
    assert db.get_setting('ai_chat_color', '') == '#DFECFC'
    assert not set(theme.OLD_COLOR_MAP) & set(theme.PALETTE)          # 새 색이 다시 옮겨지지 않음
    assert theme.PALETTE_NAMES['#DFECFC'] == '하늘'
