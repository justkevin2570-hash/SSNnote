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
    assert seen and ('교재', '빠짐', 'gone') in seen[0]
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
