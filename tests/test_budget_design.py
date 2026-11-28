"""예산 창 디자인 고르기 (⋯ → 디자인): 기본 · 사무형. 틀은 같고 색 · 모양만, 고른 것은 기억."""
import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def win(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setenv('APPDATA', str(tmp_path))
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])
    import budget
    monkeypatch.setattr(budget, '_STATE_FILE', str(tmp_path / 'budget_state.json'))
    cid = db.add_budget_category('[목적]진로')
    db.add_budget_item(cid, '운영수당', '강사수당', 1000, 950)       # 95% → 거의 다 씀
    db.add_budget_item(cid, '여비', '인솔 여비', 1000, 100)
    w = budget.BudgetWindow()
    yield w, db, budget
    w.close()
    budget.apply_design('base')
    budget.apply_text_size('base')


def test_default_is_base_and_no_design_menu(win):
    """디자인은 하나로: ⋯에 「디자인」 메뉴가 없고, 예전에 사무형을 고른 PC도 기본으로 열림."""
    w, db, budget = win
    assert w.design == 'base' and not budget.C.OFFICE
    assert not any(a.text() == '디자인' for a in w._more_menu().actions())
    db.set_setting(budget.DESIGN_KEY, 'office')
    w2 = budget.BudgetWindow()
    assert w2.design == 'base' and not budget.C.OFFICE and db.get_setting(budget.DESIGN_KEY) == 'base'
    w2.close()


def _flush():
    from PyQt5.QtCore import QEvent
    from PyQt5.QtWidgets import QApplication
    QApplication.sendPostedEvents(None, QEvent.DeferredDelete)    # 예전 화면(다시 그리기 전) 정리


def test_switch_keeps_layout_and_is_remembered(win):
    w, db, budget = win
    from budget import _BudgetTable
    cols = [t.columnCount() for t in w.findChildren(_BudgetTable)]
    tab = w.active_tab
    w.set_design('office')
    _flush()
    assert budget.C.OFFICE and budget.C.ACCENT == '#1E6B47' and '#1E6B47' in budget.STYLE
    assert db.get_setting(budget.DESIGN_KEY) == 'office' and w.active_tab == tab
    tables = w.findChildren(_BudgetTable)
    assert [t.columnCount() for t in tables] == cols
    main = next(t for t in tables if t.row_items)
    assert main.showGrid() and list(main.tint_rows) == [0]          # 95% 줄만 연한 주황 바탕
    assert budget.current_design() == 'office'
    w.set_design('base')
    _flush()
    assert not budget.C.OFFICE and budget.C.ACCENT == '#4F46E5'
    main = next(t for t in w.findChildren(_BudgetTable) if t.row_items)
    assert not main.showGrid() and not main.tint_rows


def test_office_text_contrast(win):
    w, db, budget = win
    budget.apply_design('office')
    C = budget.C

    def lum(h):
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))
        f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)

    def ratio(a, b):
        la, lb = sorted((lum(a), lum(b)), reverse=True)
        return (la + 0.05) / (lb + 0.05)
    for fg in (C.TEXT, C.MUTED, C.FAINT, C.OK_TEXT, C.WARN_TEXT, C.OVER_TEXT):
        assert ratio(fg, '#FFFFFF') >= 4.5, fg
    assert ratio(C.WARN_TEXT, C.WARN_BG) >= 4.5
    for fg in (C.NAV_TEXT, C.NAV_MUTED, C.NAV_OK, C.NAV_WARN, C.NAV_OVER):
        assert ratio(fg, C.NAV_BG) >= 4.5, fg


def test_text_size_menu_and_scaling(win):
    w, db, budget = win
    menu = next(a.menu() for a in w._more_menu().actions() if a.text() == '글자 크기')
    assert [a.text() for a in menu.actions()] == ['기본', '크게', '더 크게', '더더 크게', '아주 크게']
    assert menu.actions()[0].isChecked()
    base_row = budget._sizes()['row_h']
    base_font = w.font().pointSizeF()
    w.set_text_size('xlarge')
    _flush()
    assert db.get_setting(budget.TEXT_SIZE_KEY) == 'xlarge' and budget.current_text_size() == 'xlarge'
    assert budget._pt(10) == 13.0 and 'font-size: 13.0pt' in budget.STYLE
    assert budget._sizes()['row_h'] > base_row                    # 줄 높이도 글자를 따라 커짐
    w.set_text_size('large')
    assert budget._pt(10) == 11.5
    w.set_text_size('xxlarge')
    assert budget._pt(10) == 14.5
    w.set_text_size('huge')
    assert budget._pt(10) == 16.0
    assert budget._sizes()['row_h'] > base_row
    w.set_text_size('base')
    assert budget._pt(10) == 10 and budget._sizes()['row_h'] == base_row
    assert w.font().pointSizeF() == base_font


def test_nav_total_and_group_have_no_bar(win):
    w, db, budget = win
    gid = db.add_budget_group('묶음')
    cid = db.get_budget_categories()[0]['id']
    db.set_budget_category_group(cid, gid)
    w.refresh()
    w.show()
    refs = w._nav_refs
    assert refs['ALL'].bar.isHidden() and refs['ALL'].lbl_pct.isHidden()
    assert refs[('G', gid)].bar.isHidden() and refs[('G', gid)].lbl_pct.isHidden()
    other = next(k for k in refs if isinstance(k, int))
    assert refs[other].bar.isHidden() and refs[other].lbl_pct.isHidden()        # 목록에는 남은 돈만
    assert all('남은 돈' in refs[k].lbl_meta.text() for k in ('ALL', ('G', gid), other))
