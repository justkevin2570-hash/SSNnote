"""업무창·메모장 글꼴: 창마다 고르고 기억, 설치 안 된 글꼴은 Pretendard로, 줄 사이 여백이 있는 글꼴도 줄이 안 잘림."""
import os
import sys
import unittest.mock as um

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

_APP = None


@pytest.fixture
def env(tmp_path, monkeypatch):
    import db
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication, QMenu
    global _APP
    _APP = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMenu, 'exec_', lambda self, *a: None)
    from ui import common, ai_chat_window
    monkeypatch.setattr(ai_chat_window, 'saved_open', lambda: False)
    common.refresh_installed_fonts()
    # 이 PC에 리틱체가 설치돼 있다고 치기 (실제 글꼴 대신 Pretendard 이름을 빌려 씀)
    fam = common._pretendard_family()
    monkeypatch.setitem(common._NOTE_SPECS, 'reatic', (('없는글꼴', ''), (fam, '')))
    monkeypatch.setitem(common._NOTE_SPECS, 'reatic_narrow', (('없는 Narrow', ''),))
    monkeypatch.setitem(common._NOTE_SPECS, 'malgun', (('없는글꼴2', ''),))
    made = []
    import window

    def make(memo=False):
        w = window.MemoWindow(db.get_first_window_id(), force_memo_mode=memo)
        made.append(w)
        return w
    yield make, db, common
    for w in made:
        w.close()
    common.refresh_installed_fonts()
    common.set_current_note_font(common.DEFAULT_NOTE_FONT)


def test_font_choice_saved_per_window_and_used_by_rows(env):
    make, db, common = env
    wid = db.get_first_window_id()
    db.add_task(wid, '공문 회신', '2026-11-20')
    w = make()
    assert w._font_key == 'pretendard' and db.get_window_font(wid) == ''
    w._set_font('reatic')
    assert db.get_window_font(wid) == 'reatic'
    row = w._task_rows()[0]
    assert row.name_edit.font().family() == common.note_family('reatic')
    assert w.input_name.font().family() == common.note_family('reatic')
    w2 = make()                                          # 다시 켜도 그대로
    assert w2._font_key == 'reatic'
    w2._set_font('pretendard')
    assert db.get_window_font(wid) == ''


def test_missing_font_falls_back_and_is_disabled_in_menu(env):
    make, db, common = env
    assert not common.note_font_available('malgun')
    assert common.note_family('malgun') == common._pretendard_family()
    w = make()
    from ui import menu_style
    m = menu_style.make_menu(w)
    w._style_submenu(m, True)
    memo = [a.menu() for a in m.actions() if a.menu()][0]
    fm = [a.menu() for a in memo.actions() if a.menu() and a.text() == '글꼴'][0]
    acts = {a.text(): a for a in fm.actions()}
    assert acts['Pretendard'].isEnabled() and acts['리틱체'].isEnabled()
    assert not acts['맑은 고딕 (이 PC에 설치되지 않음)'].isEnabled()
    assert not acts['리틱체 Narrow (이 PC에 설치되지 않음)'].isEnabled()
    assert len(acts) == 4


def test_memo_window_font(env):
    make, db, common = env
    w = make(memo=True)
    w._set_font('reatic')
    assert w.memo_editor.font().family() == common.note_family('reatic')


def test_line_height_includes_font_leading():
    """줄 사이 여백(leading)이 있는 글꼴은 그만큼 업무 이름 칸이 높아져야 아랫줄이 잘리지 않는다."""
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])
    from PyQt5.QtGui import QFont, QTextLayout
    from ui.common import pr_font
    f = pr_font(12)
    lay = QTextLayout('가나다', f)
    lay.beginLayout()
    ln = lay.createLine()
    ln.setLineWidth(300)
    ln.setLeadingIncluded(True)
    lay.endLayout()
    assert ln.height() >= QFont(f).pointSize()           # 음수 여백이어도 줄이 줄어들지 않음
    import inspect
    from ui import task_row
    assert 'setLeadingIncluded(True)' in inspect.getsource(task_row._AutoHeightEdit._update_height)


def test_reatic_size_scaled_to_pretendard(env):
    """리틱체는 같은 pt에서 글자가 커서 한글 높이를 Pretendard에 맞춰 작게 그린다 (설치 안 됐으면 보정 없음)."""
    make, db, common = env
    assert common.note_scale('pretendard') == 1.0
    assert abs(common.note_scale('reatic') - 16 / 18) < 1e-9          # (테스트에서는 설치된 것으로 침)
    assert abs(common.note_font(12, 'reatic').pointSizeF() - 12 * 16 / 18) < 0.01
    assert common.note_scale('malgun') == 1.0                         # 없는 글꼴 → Pretendard로, 보정 없음
    wid = db.get_first_window_id()
    db.add_task(wid, '공문 회신', '2026-11-20')
    w = make()
    w._set_font('reatic')
    assert abs(w._task_rows()[0].name_edit.font().pointSizeF() - 12 * 16 / 18) < 0.01


def test_narrow_is_own_choice_and_bold_uses_narrow_bold(env, monkeypatch):
    """「리틱체 Narrow」는 따로 고르는 글꼴, 중요 표시(굵게)는 Narrow Bold 모양으로."""
    make, db, common = env
    fam = common._pretendard_family()
    styles = common.QFontDatabase().styles(fam)
    style = next((s for s in styles if s not in ('Regular', 'Bold')), None)
    if style is None:
        pytest.skip('모양 이름으로 고를 수 있는 글꼴이 없음')
    monkeypatch.setitem(common._NOTE_SPECS, 'reatic_narrow', (('없는 Narrow', ''), (fam, style)))
    common.refresh_installed_fonts()
    assert common._installed_spec('reatic_narrow') == (fam, style)
    assert common._installed_spec('reatic') == (fam, '')                 # 일반 리틱체는 따로
    f = common.note_font(12, 'reatic_narrow')
    assert f.styleName() == style
    assert abs(f.pointSizeF() - 12 * 16 / 18) < 0.01
    f2 = common.set_note_bold(common.note_font(12, 'reatic_narrow'))
    assert f2.styleName().endswith('Bold')
    g = common.set_note_bold(common.note_font(12, 'pretendard'))
    assert g.bold() and not g.styleName()
