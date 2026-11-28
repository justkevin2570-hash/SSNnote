"""예산 창 1단계 개선: 한 줄 요약 · [지출 추가] 단추 · 빠른 금액 · 쓴 곳 이어 쓰기 · 다시 가져오기 보존 · Ctrl+Z."""
import os

import pytest

pytest.importorskip('PyQt5')
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


@pytest.fixture
def win(tmp_path, monkeypatch):
    import db
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    import budget
    monkeypatch.setattr(budget, '_STATE_FILE', str(tmp_path / 'budget_state.json'))
    w = budget.BudgetWindow()
    yield w, db, budget, app
    w.close()


CATS = [{'name': 'A', 'items': [{'cost_type': '운영수당', 'description': '강사 수당', 'budget': 1000, 'spent': 200},
                                {'cost_type': '일반수용비', 'description': '교재', 'budget': 500, 'spent': 500}]},
        {'name': 'B', 'items': [{'cost_type': '교육운영비', 'description': '간식', 'budget': 500, 'spent': 0}]}]


def _item(db, name):
    return next(i for i in db.get_budget_items() if i['description'] == name)


def test_header_and_summary_line(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    assert w.btn_import.isVisible() and w.btn_more.isVisible() and not w.btn_manage.isVisible()
    assert w.btn_ai.isVisible() and w.ask_card.isAncestorOf(w.btn_ai)          # AI 연결은 아래 묻기 줄에
    assert w.stat['remaining'].text() == '1,300원' and w.stat['budget'].text() == '2,000원'
    assert not w.stat['spent'].isVisible() and '700' in w.stat['remaining'].toolTip()   # 집행액은 말풍선
    assert w.lbl_pct.text() == '35%'


def test_spend_button_on_every_open_row(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    db.toggle_budget_item_completed(_item(db, '교재')['id'])
    w.refresh()
    w.show()
    app.processEvents()
    open_id, done_id = _item(db, '강사 수당')['id'], _item(db, '교재')['id']
    b = w._row_refs[open_id]['spend']
    assert b.isVisible() and b.text() == '＋ 지출 추가' and b.isEnabled()        # 마우스를 안 올려도 보임
    assert not w._row_refs[done_id]['spend'].isEnabled()                       # 완료한 줄은 빈 자리
    calls = []
    monkeypatch.setattr(w, '_open_spend', lambda iid: calls.append(iid))
    w.refresh()
    w._row_refs[open_id]['spend'].click()
    assert calls == [open_id]


def test_spend_dialog_quick_amounts_and_suggestions(win):
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': 'x', 'description': 'y', 'budget': 300000, 'spent': 120000,
            'completed': 0}
    dlg = budget._SpendDialog(item, [], w, suggestions=['4월 강사비', '3월 강사비', '4월 강사비', ''])
    assert dlg.windowTitle() == '지출 추가' and dlg.btn_save.text() == '저장'
    dlg.quick_buttons['+5만'].click()
    dlg.quick_buttons['+1만'].click()
    assert dlg.edit_spend_now.text() == '60,000'
    assert '180,000' in dlg.lbl_spend_preview.text() and '120,000원' in dlg.lbl_spend_preview.text()
    dlg.quick_buttons['남은 돈 전부'].click()
    assert dlg.edit_spend_now.text() == '180,000'
    comp = dlg.edit_purpose.completer()
    assert comp is not None and comp.model().stringList() == ['4월 강사비', '3월 강사비']   # 겹치지 않게
    full = dict(item, spent=300000)
    assert not budget._SpendDialog(full, [], w).quick_buttons['남은 돈 전부'].isEnabled()


def test_reimport_keeps_expenses_and_done(win, monkeypatch):
    """다시 가져오기(같은 이름은 숫자만 새로): 사용 내역 · 완료 표시는 그대로, 집행액 · 예산만 맞춤."""
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item(db, '강사 수당')['id']
    db.add_budget_expense(iid, 100, '3월 강사비', '2026-03-31')
    db.toggle_budget_item_completed(_item(db, '교재')['id'])
    w.refresh()
    new = [{'name': 'A', 'items': [{'cost_type': '운영수당', 'description': '강사 수당', 'budget': 1200, 'spent': 450},
                                   {'cost_type': '일반수용비', 'description': '교재', 'budget': 500, 'spent': 500},
                                   {'cost_type': '여비', 'description': '출장', 'budget': 90, 'spent': 0}]}]
    changes = w._import_changes(new)
    assert ('강사 수당', '집행액 +150', 'up') in changes and ('강사 수당', '예산 1,000 → 1,200', 'budget') in changes
    assert ('출장', '새 항목', 'new') in changes and len(changes) == 3
    w._apply_import(new, 'update')
    it = _item(db, '강사 수당')
    assert it['id'] == iid and it['budget'] == 1200 and it['spent'] == 450
    assert [e['purpose'] for e in db.get_budget_expenses() if e['item_id'] == iid] == ['3월 강사비']
    assert _item(db, '교재')['completed'] and _item(db, '출장')['budget'] == 90
    assert _item(db, '간식')['budget'] == 500                                   # B는 그대로
    w._apply_import([{'name': 'A', 'items': [new[0]['items'][0]]}], 'update')            # 자료에 없는 항목은 지우지 않음
    assert _item(db, '교재') and _item(db, '출장')
    dlg = budget._ImportChangesDialog(changes, w)
    assert len(dlg.rows) == 3
    assert budget._ImportChangesDialog([], w).rows == []


def test_reimport_shows_changes_dialog_before_apply(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    seen = []

    class FakePreview:
        def __init__(self, *a, **k):
            self.categories, self.mode = [{'name': 'A', 'items': [CATS[0]['items'][0]]}], 'update'

        def exec_(self):
            return budget.QDialog.Accepted
    monkeypatch.setattr(budget, '_ImportPreviewDialog', FakePreview)
    monkeypatch.setattr(budget._ImportChangesDialog, 'exec_', lambda self: seen.append(self.rows) or budget.QDialog.Rejected)
    w._import_parsed(CATS, '테스트')
    assert seen and ('교재', '자료에 없음 · 그대로 둬요', 'gone') in seen[0]
    assert _item(db, '교재')                                                   # 취소하면 그대로


def test_ctrl_z_undoes_spend_and_delete(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item(db, '강사 수당')['id']

    def fake_exec(self):
        self.result_data = {'amount': 300, 'purpose': '4월', 'spent_on': '2026-04-30'}
        return budget.QDialog.Accepted
    monkeypatch.setattr(budget._SpendDialog, 'exec_', fake_exec)
    w._open_spend(iid)
    assert _item(db, '강사 수당')['spent'] == 500 and 'Ctrl+Z' in w.ask_status.text()
    w._do_undo()
    assert _item(db, '강사 수당')['spent'] == 200 and not db.get_budget_expenses()
    w._do_undo()                                                               # 두 번 눌러도 그대로
    assert _item(db, '강사 수당')['spent'] == 200
    eid = db.add_budget_expense(iid, 50, '간식', '2026-05-01')
    w.refresh()
    monkeypatch.setattr(budget.QMessageBox, 'question', lambda *a, **k: budget.QMessageBox.Yes)
    w._delete_expense(eid)
    assert _item(db, '강사 수당')['spent'] == 200
    w._do_undo()
    assert _item(db, '강사 수당')['spent'] == 250 and [e['purpose'] for e in db.get_budget_expenses()] == ['간식']


# ── 2단계: 찾기 · 빠른 보기 · 남은 돈 많은 순 · 엑셀 내보내기 ──────────────
CATS2 = [{'name': 'A', 'items': [{'cost_type': '운영수당', 'description': '강사 수당', 'budget': 1000, 'spent': 950},
                                 {'cost_type': '일반수용비', 'description': '교재', 'budget': 500, 'spent': 600},
                                 {'cost_type': '여비', 'description': '출장 여비', 'budget': 900, 'spent': 0}]},
         {'name': 'B', 'items': [{'cost_type': '교육운영비', 'description': '간식', 'budget': 500, 'spent': 100}]}]


def _visible(w):
    out = []
    for cid, refs in w._sec_refs.items():
        if not refs['card'].isVisible():
            continue
        for t in refs['tables']:
            for i in range(len(t.row_costs)):
                if not t.isRowHidden(2 * i):
                    out.append(w._items_by_id[t.row_items[2 * i]]['description'])
    return sorted(out)


def test_search_filters_rows_and_hides_empty_sections(win):
    w, db, budget, app = win
    w._apply_import(CATS2, 'append')
    db.add_budget_expense(next(i['id'] for i in db.get_budget_items() if i['description'] == '간식'), 50, '체육대회 음료', '2026-05-01')
    w.refresh()
    w.show()
    app.processEvents()
    w._set_tab(w.categories[1]['id'])
    w.search.setText('여비')                                       # 찾기를 시작하면 전체로
    app.processEvents()
    assert w.active_tab == 'ALL' and _visible(w) == ['출장 여비'] and w.lbl_found.text() == '1개'
    assert not w._sec_refs[w.categories[1]['id']]['card'].isVisible()   # 하나도 안 남은 세부항목은 숨김
    w.search.setText('체육 대회')                                   # 쓴 곳 · 띄어쓰기 무시
    assert _visible(w) == ['간식']
    w.search.setText('없는말')
    assert _visible(w) == [] and w.lbl_found.text() == '찾는 항목이 없어요'
    w.search.setFocus()
    from PyQt5.QtTest import QTest
    from PyQt5.QtCore import Qt
    QTest.keyClick(w.search, Qt.Key_Escape)                         # Esc → 지우기
    assert w.search.text() == '' and len(_visible(w)) == 4 and w.lbl_found.text() == ''


def test_quick_views_and_sort(win):
    w, db, budget, app = win
    w._apply_import(CATS2, 'append')
    w.show()
    app.processEvents()
    w.quick_btns['warn'].click()
    assert _visible(w) == ['강사 수당']                              # 95%
    w.quick_btns['over'].click()                                    # 하나만 고름
    assert not w.quick_btns['warn'].isChecked() and _visible(w) == ['교재']
    w.quick_btns['over'].click()
    assert w._quick is None and len(_visible(w)) == 4
    w.btn_sort.click()                                              # 남은 돈 많은 순
    assert db.get_setting('budget_sort_rem') == '1'
    t = w._sec_refs[w.categories[0]['id']]['tables'][0]
    order = [w._items_by_id[t.row_items[2 * i]]['description'] for i in range(len(t.row_costs))]
    assert order == ['출장 여비', '강사 수당', '교재']


def test_export_excel(win, tmp_path, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS2, 'append')
    iid = next(i['id'] for i in db.get_budget_items() if i['description'] == '간식')
    db.add_budget_expense(iid, 50, '음료', '2026-05-01')
    w.refresh()
    monkeypatch.setattr(budget.QMessageBox, "warning", lambda *a, **k: (_ for _ in ()).throw(AssertionError(a[2])))
    import sys
    if not isinstance(sys.modules.get('winreg'), type(sys)):     # 다른 테스트가 가짜 winreg를 넣었으면
        monkeypatch.delitem(sys.modules, 'winreg', raising=False)  # (openpyxl → mimetypes가 레지스트리를 읽다 멈춤)
    import mimetypes
    if getattr(mimetypes, '_winreg', None) is not None:
        monkeypatch.setattr(mimetypes, '_winreg', None)
    path = w._export_excel(str(tmp_path / "out.xlsx"))
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb['예산']
    rows = [[c.value for c in r] for r in ws.iter_rows()]
    assert rows[0][:6] == ['세부항목', '원가통계비목', '산출내역', '예산현액', '집행액', '잔액']
    assert ['A', '운영수당', '강사 수당', 1000, 950, 50] == rows[1][:6]
    assert rows[4][0] == 'A 합계' and rows[4][3] == 2400 and rows[4][4] == 1550
    assert rows[-1][0] == '전체 합계' and rows[-1][3] == 2900 and rows[-1][4] == 1700
    ex = [[c.value for c in r] for r in wb['사용 내역'].iter_rows()]
    assert ex[1] == ['B', '간식', '2026-05-01', '음료', 50]
    assert '엑셀로 저장' in w.ask_status.text()
    assert any('엑셀로 내보내기' in a.text() for a in w._more_menu().actions())


# ── 3단계: 남은 돈 정리 ─────────────────────────────────────────
def test_fiscal_year_end():
    import budget
    from datetime import date
    assert budget.fiscal_year_end(date(2026, 11, 28)) == date(2027, 2, 28)
    assert budget.fiscal_year_end(date(2028, 1, 10)) == date(2028, 2, 29)     # 윤년
    assert budget.fiscal_year_end(date(2026, 3, 1)) == date(2027, 2, 28)


def test_leftover_dialog_lists_and_acts(win, monkeypatch):
    w, db, budget, app = win
    cats = [{'name': 'A', 'items': [{'cost_type': 'x', 'description': '많이 남음', 'budget': 900000, 'spent': 100000},
                                    {'cost_type': 'x', 'description': '조금 남음', 'budget': 50000, 'spent': 45000},
                                    {'cost_type': 'x', 'description': '다 씀', 'budget': 1000, 'spent': 1000}]},
            {'name': 'B', 'items': [{'cost_type': 'x', 'description': '중간', 'budget': 300000, 'spent': 0}]}]
    w._apply_import(cats, 'append')
    monkeypatch.setattr(budget._LeftoverDialog, 'exec_', lambda self: 0)
    dlg = w._open_leftover()
    names = [dlg.table.item(r, 1).text() for r in range(dlg.table.rowCount())]
    assert names == ['많이 남음', '중간']                                  # 남은 돈 많은 순, 1만 원 미만은 뺌
    assert dlg.table.item(0, 0).text() == 'A' and dlg.table.item(0, 3).text() == '800,000'
    assert dlg.table.item(0, 4).text() == '89%'
    assert '1,105,000원' in dlg.lbl_sum.text() and '3</b>개' in dlg.lbl_sum.text()
    dlg.chk_small.setChecked(False)
    assert dlg.table.rowCount() == 3
    iid = dlg.rows[1][0]                                                   # 중간
    dlg._done(iid)
    assert db.get_budget_items() and next(i for i in db.get_budget_items() if i['id'] == iid)['completed']
    assert [dlg.table.item(r, 1).text() for r in range(dlg.table.rowCount())] == ['많이 남음', '조금 남음']

    def fake_exec(self):
        self.result_data = {'amount': 800000, 'purpose': '물품', 'spent_on': '2026-12-01'}
        return budget.QDialog.Accepted
    monkeypatch.setattr(budget._SpendDialog, 'exec_', fake_exec)
    dlg.rows[0][1].click()                                                 # [＋ 지출 추가] → 다 써서 빠짐
    assert [dlg.table.item(r, 1).text() for r in range(dlg.table.rowCount())] == ['조금 남음']
    assert '남았어요' in dlg.lbl_days.text()
    assert w.btn_leftover.text() == '남은 돈 정리'
