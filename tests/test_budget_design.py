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


def test_default_is_base_and_menu_offers_both(win):
    w, db, budget = win
    assert w.design == 'base' and not budget.C.OFFICE
    menu = next(a.menu() for a in w._more_menu().actions() if a.text() == '디자인')
    acts = [a for a in menu.actions()]
    assert [a.text().split()[0] for a in acts] == ['기본', '사무형']
    assert acts[0].isChecked() and not acts[1].isChecked()


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
