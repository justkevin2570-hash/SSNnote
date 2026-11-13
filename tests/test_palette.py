"""배경색 팔레트 (v3.14: 밝은 6색 + 흰색): 예전 색 옮기기 · AI 색 이름 · 견본 · 글자 대비."""
import os
import sys
import unittest.mock as um

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.modules.setdefault('winreg', um.MagicMock())

from ui import theme

NEW = {'노랑': '#FFF5A5', '연두': '#E4F9E0', '분홍': '#FFE4F1', '라벤더': '#F2E6FF', '하늘': '#E2F1FF',
       '연회색': '#F3F2F1', '흰색': '#FFFFFF'}


def test_palette_is_seven_colors_and_old_colors_map_into_it():
    assert theme.PALETTE_NAMES == {hx: name for name, hx in NEW.items()}
    assert list(theme.PALETTE) == list(NEW.values()) and theme.DEFAULT_NOTE_COLOR == '#FFF5A5'
    assert all(new in theme.PALETTE for new in theme.OLD_COLOR_MAP.values())
    assert not set(theme.OLD_COLOR_MAP) & set(theme.PALETTE)          # 새 색이 다시 옮겨지지 않음


def test_old_colors_go_to_nearest_new_name():
    m = theme.OLD_COLOR_MAP
    # v3.11~v3.13
    assert m['#FFF8C2'] == m['#FFEAC0'] == NEW['노랑']                # 노랑·귤 → 노랑
    assert m['#E9F5D3'] == m['#DCF4EC'] == NEW['연두']                # 연두·민트 → 연두
    assert m['#FDE2EB'] == m['#FFE4D8'] == NEW['분홍']                # 분홍·복숭아 → 분홍
    assert m['#EBE5FA'] == NEW['라벤더'] and m['#DFECFC'] == NEW['하늘']
    assert m['#EDEDEB'] == m['#F3EBDE'] == NEW['연회색']              # 연회색·모래 → 연회색
    # v3.06~v3.10, v3.05까지
    assert '#FFF5A5' not in m and m['#FFF7D1'] == m['#FEFFA7'] == m['#FDD663'] == NEW['노랑']   # 예전 노랑 그대로, v3.14 노랑은 되돌림
    assert m['#AECBFA'] == m['#D2E4FA'] == NEW['하늘']
    assert m['#AAAAAA'] == m['#E6C9A8'] == NEW['연회색']


def test_init_db_moves_old_colors_and_keeps_white(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 's.db'))
    db.init_db()
    ids = [db.get_first_window_id()] + [db.create_window() for _ in range(4)]
    olds = ['#ffeac0', '#AAAAAA', '#FFFFFF', '', '#DCF4EC']
    with db._connect() as conn:
        for i, c in zip(ids, olds):
            conn.execute('UPDATE windows SET color=? WHERE id=?', (c, i))
    db.set_setting('ai_chat_color', '#dfecfc')
    db.init_db()
    colors = {w['id']: w['color'] for w in db.get_all_windows()}
    assert [colors[i] for i in ids] == [NEW['노랑'], NEW['연회색'], '#FFFFFF', '', NEW['연두']]
    assert db.get_setting('ai_chat_color', '') == NEW['하늘']
    db.init_db()                                                       # 두 번 돌려도 그대로
    assert {w['id']: w['color'] for w in db.get_all_windows()} == colors


def test_ai_color_names_and_old_aliases():
    import ai_actions
    assert ai_actions._color('하늘') == NEW['하늘']
    assert ai_actions._color('보라') == NEW['라벤더']
    assert ai_actions._color('귤') == ai_actions._color('주황') == NEW['노랑']      # 없어진 색은 가까운 색으로
    assert ai_actions._color('복숭아') == NEW['분홍'] and ai_actions._color('민트') == NEW['연두']
    assert ai_actions._color('모래') == ai_actions._color('회색') == NEW['연회색']
    assert ai_actions._color('흰색') == '#FFFFFF'
    assert ai_actions._color('#FDCFE8') == NEW['분홍']                 # 예전 색 코드


def test_swatches_one_row_and_mark_current_color():
    from PyQt5.QtWidgets import QApplication, QMenu
    QApplication.instance() or QApplication([])
    from ui import style_menu
    w = style_menu.color_swatches(QMenu(), NEW['하늘'], lambda c: None)
    sw = w.findChildren(style_menu.Swatch)
    assert len(sw) == 7 and [s.color for s in sw if s.chosen] == [NEW['하늘']]
    assert sw[0].toolTip() == '노랑'
    grid = w.layout()
    assert {grid.getItemPosition(i)[0] for i in range(grid.count())} == {0}     # 한 줄
    w2 = style_menu.color_swatches(QMenu(), '=', lambda c: None, same_as_note=True)
    assert [s.color for s in w2.findChildren(style_menu.Swatch) if s.chosen] == ['=']


def _lum(hx):
    c = [int(hx[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    c = [x / 12.92 if x <= 0.04045 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def test_body_text_contrast_on_every_background():
    """글자색은 그대로(본문 #333) — 모든 바탕에서 대비 7:1 이상."""
    for bg in theme.PALETTE:
        a, b = _lum(bg), _lum('#333333')
        assert (a + 0.05) / (b + 0.05) >= 7, bg
