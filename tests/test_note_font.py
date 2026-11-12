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


def test_missing_font_falls_back_and_menu_offers_install(env):
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
    assert len(acts) == 4 and acts['Pretendard'].isEnabled() and acts['리틱체'].isEnabled()
    assert acts['리틱체 Narrow (설치 필요)'].isEnabled()             # 눌러서 받는 곳 안내
    assert acts['맑은 고딕 (설치 필요)'].isEnabled()


def _fake_dialog(monkeypatch, clicks):
    """안내 창 버튼을 차례로 누른 것처럼. 누를 때마다 창의 글을 모아 둔다."""
    from PyQt5.QtWidgets import QDialog, QLabel
    shown = []

    def exec_(dlg):
        for name in clicks:
            if dlg.result() or not hasattr(dlg, '_buttons'):
                break
            dlg._buttons[name].click()
            shown.append(' '.join(lb.text() for lb in dlg.findChildren(QLabel)))
        return dlg.result()
    monkeypatch.setattr(QDialog, 'exec_', exec_)
    return shown


def test_missing_font_dialog_opens_official_download_then_checks(env, monkeypatch):
    make, db, common = env
    w = make()
    opened = []
    from PyQt5.QtGui import QDesktopServices
    monkeypatch.setattr(QDesktopServices, 'openUrl', lambda url: opened.append(url.toString()) or True)
    shown = _fake_dialog(monkeypatch, ['받기', '설치 확인', '닫기'])
    assert w._font_missing('reatic_narrow') is False
    assert opened and opened[0].startswith('https://www.reaticindustry.com/')
    assert 'REATIC-Narrow.ttf' in shown[0] and 'REATIC-NarrowBold.ttf' in shown[0]
    assert '아직 찾지 못했어요' in shown[1]
    # 이제 설치됐다고 치면 [설치 확인]으로 바로 바뀜
    fam = common._pretendard_family()
    monkeypatch.setitem(common._NOTE_SPECS, 'reatic_narrow', (('없는 Narrow', ''), (fam, '')))
    _fake_dialog(monkeypatch, ['설치 확인'])
    assert w._font_missing('reatic_narrow') is True
    assert w._font_key == 'reatic_narrow' and db.get_window_font(w.window_id) == 'reatic_narrow'


def test_malgun_missing_has_no_download(env, monkeypatch):
    make, db, common = env
    w = make()
    shown = _fake_dialog(monkeypatch, ['설치 확인'])
    w._font_missing('malgun')
    assert '받기' not in w._font_dialog._buttons
    assert 'Windows 설정' in shown[0] and 'reaticindustry' not in shown[0]


def test_font_installed_after_start_is_loaded_from_file(env, monkeypatch, tmp_path):
    """앱을 켠 뒤 설치한 글꼴: Windows 글꼴 파일을 찾아 앱에 불러와 바로 쓴다."""
    make, db, common = env
    import shutil
    src = os.path.join(common._base_path(), 'assets', 'font', 'Pretendard-Thin.ttf')
    dst = tmp_path / 'REATIC-Narrow.ttf'
    shutil.copy(src, dst)
    from ui import font_detect
    monkeypatch.setattr(font_detect, 'installed_font_files', lambda: [str(dst), str(tmp_path / 'other.ttf')])
    calls = []
    real = common.QFontDatabase.addApplicationFont
    monkeypatch.setattr(common.QFontDatabase, 'addApplicationFont',
                        staticmethod(lambda p: calls.append(p) or real(p)))
    monkeypatch.setitem(common._NOTE_SPECS, 'reatic_narrow', (('없는 Narrow', ''),))
    common.refresh_installed_fonts()
    assert not common.note_font_available('reatic_narrow')          # 파일은 불러왔지만 이름이 안 맞음
    assert calls == [str(dst)]
    common.refresh_installed_fonts()
    common.note_font_available('reatic_narrow')
    assert calls == [str(dst)]                                         # 같은 파일은 한 번만 불러옴


def test_find_files_matches_names(tmp_path):
    from ui import font_detect
    files = [r'C:\Users\a\AppData\Local\Microsoft\Windows\Fonts\REATIC-Narrow.otf',
             r'C:\Users\a\AppData\Local\Microsoft\Windows\Fonts\reatic-narrow.ttf',
             r'C:\Windows\Fonts\REATIC-NarrowBold.ttf', r'C:\Windows\Fonts\malgun.ttf']
    found = font_detect.find_files(('REATIC-Narrow', 'REATIC-NarrowBold'), files=[f.replace('\\', '/') for f in files])
    assert found['REATIC-Narrow'].endswith('reatic-narrow.ttf')            # .ttf 먼저
    assert found['REATIC-NarrowBold'].endswith('REATIC-NarrowBold.ttf')
    assert font_detect.find_files(('REATIC-Regular',), files=[]) == {}


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
