"""사용 내역(명목별 지출 기록): DB 기록·삭제·연쇄 삭제, AI 쓴 곳, 산출내역 칸 펼치기·더 보기·삭제, 되돌리기."""
import json

import pytest

pytest.importorskip('PyQt5')

_APP = None


@pytest.fixture
def db(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    return db


def _item(db, desc='협의회비', budget=1150000, spent=0):
    cid = db.add_budget_category(f'[목적]{desc}')
    iid = db.add_budget_item(cid, '운영수당', desc, budget, spent)
    return cid, iid


def _spent(db, iid):
    return next(i for i in db.get_budget_items() if i['id'] == iid)['spent']


# ── DB ─────────────────────────────────────────────────────────────

def test_add_expense_records_purpose_and_adds_to_spent(db):
    _, iid = _item(db)
    db.add_budget_expense(iid, 150000, '체육대회 협의회비', '2026-09-12')
    db.add_budget_expense(iid, 150000, ' 교육반성회 식사비 ', '2026-10-02')
    assert _spent(db, iid) == 300000
    rows = [e for e in db.get_budget_expenses() if e['item_id'] == iid]
    assert [e['purpose'] for e in rows] == ['교육반성회 식사비', '체육대회 협의회비']   # 최근 날짜 먼저, 앞뒤 공백 제거


def test_delete_expense_adjusts_spent_unless_told_not_to(db):
    _, iid = _item(db)
    eid = db.add_budget_expense(iid, 150000, '간식', '2026-11-01')
    db.delete_budget_expense(eid, adjust_spent=False)            # 되돌리기 때: 기록만 지움
    assert _spent(db, iid) == 150000 and db.get_budget_expenses() == []
    eid = db.add_budget_expense(iid, 50000, '간식', '2026-11-02')
    db.delete_budget_expense(eid)                                 # 화면에서 지울 때: 집행액도 줄임
    assert _spent(db, iid) == 150000


def test_save_item_edit_with_expense_is_one_step(db):
    _, iid = _item(db, spent=100000)
    db.save_budget_item_edit(iid, _item_cat(db, iid), '운영수당', '협의회비', 1150000, 250000,
                             {'amount': 150000, 'purpose': '체육대회', 'spent_on': '2026-11-10'})
    assert _spent(db, iid) == 250000                              # 집행액은 저장값 그대로 (다시 더하지 않음)
    assert [e['purpose'] for e in db.get_budget_expenses()] == ['체육대회']


def _item_cat(db, iid):
    return next(i for i in db.get_budget_items() if i['id'] == iid)['category_id']


def test_deleting_item_or_category_or_all_removes_expenses(db):
    cid, iid = _item(db)
    db.add_budget_expense(iid, 1, '가', '2026-01-01')
    db.delete_budget_item(iid)
    assert db.get_budget_expenses() == []
    cid, iid = _item(db, desc='다른 것')
    db.add_budget_expense(iid, 2, '나', '2026-01-01')
    db.delete_budget_category(cid)
    assert db.get_budget_expenses() == []
    cid, iid = _item(db, desc='마지막')
    db.add_budget_expense(iid, 3, '다', '2026-01-01')
    db.clear_budget_data()
    assert db.get_budget_expenses() == []


# ── AI가 쓴 곳 ───────────────────────────────────────────────────────

BY_ID = {11: {'id': 11, 'category_id': 1, 'description': '간식', 'cost_type': '교육운영비',
              'budget': 400000, 'spent': 115500, 'completed': 0}}


def test_ai_spend_keeps_purpose_and_trims_it():
    import budget_ai as ba
    long = '가' * 80
    raw = json.dumps({'answer': 'x', 'spend': [{'item_id': 11, 'amount': 50000, 'purpose': long}]}, ensure_ascii=False)
    c = ba.parse_reply(raw, BY_ID)['changes'][0]
    assert c['purpose'] == '가' * ba.PURPOSE_MAX and c['kind'] == 'spend'


def test_ai_purpose_defaults_to_blank_and_only_spend_has_it():
    import budget_ai as ba
    raw = json.dumps({'answer': 'x', 'spend': [{'item_id': 11, 'amount': 5000}],
                      'set_spent': [{'item_id': 11, 'spent': 120000}]}, ensure_ascii=False)
    r = ba.parse_reply(raw, BY_ID)
    assert [c['purpose'] for c in r['changes']] == ['']         # 같은 항목이라 두 번째는 제외되지만 쓴 곳은 빈 값
    raw = json.dumps({'answer': 'x', 'set_spent': [{'item_id': 11, 'spent': 120000}]})
    assert ba.parse_reply(raw, BY_ID)['changes'][0]['purpose'] == ''


def test_prompt_format_and_rules_mention_purpose():
    import budget_ai as ba
    assert '"purpose"' in ba._FORMAT
    assert 'purpose' in ba.SYSTEM and '쓴 곳' in ba.SYSTEM


# ── 화면 ───────────────────────────────────────────────────────────

@pytest.fixture
def win(tmp_path, monkeypatch, db):
    monkeypatch.setenv('APPDATA', str(tmp_path))
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])
    import budget
    monkeypatch.setattr(budget, '_STATE_FILE', str(tmp_path / 'budget_state.json'))
    w = budget.BudgetWindow()
    yield w, db, budget, _APP
    w.close()


def _open_one_item(w, db, desc='협의회비', spent=0, budget_amt=1150000):
    cid, iid = _item(db, desc=desc, budget=budget_amt, spent=spent)
    w.refresh()
    return iid


def test_header_has_manage_button_and_no_add_button(win):
    w, db, budget, app = win
    _open_one_item(w, db)
    assert not hasattr(w, 'btn_add_item')
    assert w.btn_manage.isHidden()                                    # 머리줄은 [가져오기] [⋯]만
    texts = [a.text() for a in w._more_menu().actions()]
    assert texts[0] == '세부항목 관리…'                                 # ⋯ 맨 위로 옮김
    assert any('원가통계비목 목록' in t for t in texts) and any('전체 지우기' in t for t in texts)


def test_delete_category_asks_then_removes_items_and_records(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtWidgets import QMessageBox
    asked = []
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: asked.append(a[2]) or QMessageBox.No)
    iid = _open_one_item(w, db, spent=0)
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    w.refresh()
    cid = db.get_budget_items()[0]['category_id']
    w._delete_category(cid)                                              # 아니요 → 그대로
    assert '항목 1개' in asked[0] and '150,000원' in asked[0] and '되돌릴 수 없습니다' in asked[0]
    assert db.get_budget_items() and db.get_budget_expenses()
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: QMessageBox.Yes)
    w.active_tab = cid
    w._delete_category(cid)                                              # 예 → 세부항목·항목·사용 내역 모두 삭제
    assert db.get_budget_categories() == [] and db.get_budget_items() == [] and db.get_budget_expenses() == []
    assert w.active_tab == 'ALL'


def _cell(w, iid):
    return w._row_refs[iid]['desc']


def test_row_is_collapsed_by_default_without_summary(win):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=150000)                    # 먼저 있던 150,000원 (명목 없음)
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')  # 집행액 300,000원
    w.refresh()
    cell = _cell(w, iid)
    assert cell._table.isRowHidden(cell._row + 1)                    # 접힌 상태가 기본
    assert cell.lbl_sum.text() == '' and cell.lbl_name.text() == '협의회비'
    assert cell._table.rowHeight(cell._row) == budget._sizes()['row_h']   # 항목 줄 높이는 그대로
    cell._state['open'] = True
    cell.set_expenses(cell._expenses)
    first = cell._dv.itemAt(0).widget().layout()                 # 맨 위 줄 = 명목 없음 (날짜 없이)
    first_labels = [first.itemAt(i).widget().text() for i in range(first.count())
                    if isinstance(first.itemAt(i).widget(), type(cell.lbl_sum))]
    assert first_labels == ['', '명목 없음', '150,000원']


def test_summary_warns_only_when_records_exceed_spent(win):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=115500)                  # 예전 자료: 쓴 곳 기록 없음
    assert _cell(w, iid).lbl_sum.text() == ''
    iid2 = _open_one_item(w, db, desc='빈 항목', spent=0)
    assert _cell(w, iid2).lbl_sum.text() == ''
    db.add_budget_expense(iid2, 5000, '잘못 적음', '2026-09-12')
    with db._connect() as conn:                                # 집행액을 기록보다 작게 낮춘다 (기록 초과 상황)
        conn.execute('UPDATE budget_items SET spent=0 WHERE id=?', (iid2,))
    w.refresh()
    assert _cell(w, iid2).lbl_sum.text() == '⚠ 기록 초과'


def test_toggle_expands_and_grows_row_then_folds_back(win):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=150000)
    db.add_budget_expense(iid, 150000, '체육대회 협의회비', '2026-09-12')
    w.refresh()
    cell = _cell(w, iid)
    table, row = cell._table, cell._row
    base = table.rowHeight(row)
    cell.btn_toggle.click()
    assert not table.isRowHidden(row + 1)
    assert table.rowHeight(row + 1) > 0 and table.rowHeight(row) == base     # 항목 줄은 그대로, 아래 줄이 펼쳐짐
    cell.btn_toggle.click()
    assert table.isRowHidden(row + 1)


def test_shows_five_then_more_and_unassigned_amount(win):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=0)
    for n in range(7):
        db.add_budget_expense(iid, 100000, f'명목{n}', f'2026-10-{n + 1:02d}')
    w.refresh()
    cell = _cell(w, iid)
    cell._state['open'] = True
    cell.set_expenses(cell._expenses)
    texts = ' '.join(lb.text() for lb in cell.detail.findChildren(type(cell.lbl_sum)))
    buttons = [b.text() for b in cell.detail.findChildren(type(cell.btn_toggle))]
    assert '더 보기 (2건 더)' in buttons
    assert '명목 없음' not in texts                              # 기록 합계와 집행액이 같으면 명목 없음 줄은 없음
    assert texts.count('100,000원') == 5                         # 먼저 5건만
    cell._toggle_more()
    assert cell._state.get('more') is True
    texts = ' '.join(lb.text() for lb in cell.detail.findChildren(type(cell.lbl_sum)))
    assert texts.count('100,000원') == 7


def test_unassigned_warning_when_records_exceed_spent(win):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=0)
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    db.set_budget_item_spent(iid, 100000)                        # 집행액을 직접 낮춘 것과 같은 상황
    w.refresh()
    cell = _cell(w, iid)
    cell._state['open'] = True
    cell.set_expenses(cell._expenses)
    labels = ' '.join(lb.text() for lb in cell.detail.findChildren(type(cell.lbl_sum)))
    assert '기록 합계가 집행액보다 큽니다' in labels


def test_delete_expense_lowers_spent_and_keeps_row_open(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtWidgets import QMessageBox
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: QMessageBox.Yes)
    iid = _open_one_item(w, db, spent=0)
    eid = db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    db.add_budget_expense(iid, 150000, '교육반성회', '2026-10-02')
    w._expense_ui[iid] = {'open': True}
    w.refresh()
    w._delete_expense(eid)
    assert _spent(db, iid) == 150000
    cell = _cell(w, iid)
    assert w._expense_ui[iid]['open'] is True and not cell._table.isRowHidden(cell._row + 1)


def test_spend_dialog_records_purpose_with_the_spend(win, monkeypatch):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=100000)

    def fake_exec(self):
        self.edit_spend_now.setText('150000')
        self.edit_purpose.setText('체육대회 협의회비')
        self.edit_spend_date.setText('2026-11-10')
        self._save()
        return budget.QDialog.Accepted if self.result_data else budget.QDialog.Rejected
    monkeypatch.setattr(budget._SpendDialog, 'exec_', fake_exec)
    w._open_spend(iid)
    assert _spent(db, iid) == 250000
    ex = db.get_budget_expenses()
    assert [(e['purpose'], e['amount'], e['spent_on']) for e in ex] == [('체육대회 협의회비', 150000, '2026-11-10')]


def test_spend_dialog_rejects_bad_date_and_keeps_data(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtWidgets import QMessageBox
    warned = []
    monkeypatch.setattr(QMessageBox, 'warning', lambda *a, **k: warned.append(a[2]))
    iid = _open_one_item(w, db, spent=100000)
    monkeypatch.setattr(budget._SpendDialog, 'exec_', lambda self: (
        self.edit_spend_now.setText('5000'), self.edit_spend_date.setText('11월 10일'), self._save(),
        budget.QDialog.Rejected)[-1])
    w._open_spend(iid)
    assert warned and '날짜' in warned[0]
    assert _spent(db, iid) == 100000 and db.get_budget_expenses() == []


# ── AI 넣기 · 되돌리기 ───────────────────────────────────────────────

class _FakeAdapter:
    def __init__(self, reply):
        self._reply = reply

    @staticmethod
    def get_api_key():
        return 'k'

    def complete(self, system, prompt, image_b64=None, max_tokens=0):
        return self._reply


def test_ai_spend_with_purpose_is_recorded_and_undone(win, monkeypatch):
    w, db, budget, app = win
    import ai_client
    import school_profile
    iid = _open_one_item(w, db, spent=200000)
    reply = json.dumps({'answer': '체육대회에 쓴 것으로 넣을게요.',
                        'actions': [{'do': 'budget.spend', 'item_id': iid, 'amount': 50000, 'purpose': '체육대회'}]},
                       ensure_ascii=False)
    import ai_actions
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: _FakeAdapter(reply))
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'claude')
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: False)
    monkeypatch.setattr(ai_actions.AskWorker, 'start', lambda self: self.run())
    w.refresh()
    w._ask('체육대회에 5만원 썼어')
    assert '체육대회' in w.ask_confirm.lbl_proposal.text()
    w.ask_confirm.apply()
    assert _spent(db, iid) == 250000
    assert [e['purpose'] for e in db.get_budget_expenses()] == ['체육대회']
    w._undo_apply()
    assert _spent(db, iid) == 200000 and db.get_budget_expenses() == []   # 기록도 함께 지워짐


# ── 사용 내역 고치기 · 산출내역 글자 복사 ─────────────────────────────────

def test_update_expense_adjusts_spent_by_difference(db):
    _, iid = _item(db, spent=100000)
    eid = db.add_budget_expense(iid, 50000, '간식', '2026-11-01')          # 집행액 150,000
    assert db.update_budget_expense(eid, 80000, ' 간식비 ', '2026-11-02')
    ex = db.get_budget_expenses()[0]
    assert (ex['amount'], ex['purpose'], ex['spent_on']) == (80000, '간식비', '2026-11-02')
    assert _spent(db, iid) == 180000
    assert db.update_budget_expense(eid, 30000, '간식비', '2026-11-02') and _spent(db, iid) == 130000
    assert db.update_budget_expense(9999, 1, '', '2026-11-02') is False


def test_record_unassigned_expense(db):
    _, iid = _item(db, spent=115500)                                       # 기록 없이 집행액만 (명목 없음)
    db.record_unassigned_expense(iid, 115500, 115500, '협의회 다과', '2026-10-01')
    assert _spent(db, iid) == 115500                                      # 같은 금액 → 집행액 그대로
    assert [(e['purpose'], e['amount']) for e in db.get_budget_expenses()] == [('협의회 다과', 115500)]


def test_edit_dialog_prefills_and_saves_difference(win, monkeypatch):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=0)
    eid = db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    w._expense_ui[iid] = {'open': True}
    w.refresh()
    seen = {}

    def fake_exec(self):
        seen['title'] = self.windowTitle()
        seen['fields'] = (self.edit_spend_now.text(), self.edit_purpose.text(), self.edit_spend_date.text())
        seen['btn'] = self.btn_save.text()
        self.edit_spend_now.setText('200000')
        seen['preview'] = self.lbl_spend_preview.text()
        self.edit_purpose.setText('체육대회 간식')
        self._save()
        return budget.QDialog.Accepted
    monkeypatch.setattr(budget._SpendDialog, 'exec_', fake_exec)
    ex = next(e for e in w._expenses_by_item[iid] if e['id'] == eid)
    w._edit_expense(iid, ex)
    assert seen['title'] == '사용 내역 수정' and seen['btn'] == '저장'
    assert seen['fields'] == ('150,000', '체육대회', '2026-09-12')
    assert '1,000,000' in seen['preview'] and '950,000' in seen['preview']   # 잔액 1,000,000 → 950,000
    assert _spent(db, iid) == 200000
    assert db.get_budget_expenses()[0]['purpose'] == '체육대회 간식'
    assert w._expense_ui[iid]['open'] is True                              # 펼친 줄은 그대로


def test_edit_cancel_changes_nothing(win, monkeypatch):
    w, db, budget, app = win
    iid = _open_one_item(w, db, spent=0)
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    w.refresh()
    monkeypatch.setattr(budget._SpendDialog, 'exec_', lambda self: budget.QDialog.Rejected)
    w._edit_expense(iid, w._expenses_by_item[iid][0])
    assert _spent(db, iid) == 150000 and db.get_budget_expenses()[0]['purpose'] == '체육대회'


def test_expense_rows_have_edit_button_and_double_click(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtCore import Qt, QEvent, QPointF
    from PyQt5.QtGui import QMouseEvent
    iid = _open_one_item(w, db, spent=50000)                               # 명목 없음 50,000
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    w._expense_ui[iid] = {'open': True}
    w.refresh()
    cell = _cell(w, iid)
    edits = [b for b in cell.detail.findChildren(budget._RowIconButton) if '고치기' in b.toolTip() or '적기' in b.toolTip()]
    assert len(edits) == 2                                                 # 명목 없음 줄 + 기록 줄
    calls = []
    monkeypatch.setattr(w, '_edit_expense', lambda i, ex: calls.append((i, ex['id'], ex['amount'])))
    w.refresh()
    cell = _cell(w, iid)
    edits = [b for b in cell.detail.findChildren(budget._RowIconButton) if '고치기' in b.toolTip() or '적기' in b.toolTip()]
    edits[0].click()
    assert calls[-1] == (iid, None, 50000)                                 # 명목 없음 → 기록 남기기
    lbl = next(lb for lb in cell.detail.findChildren(type(cell.lbl_sum)) if lb.text() == '체육대회')
    assert lbl.textInteractionFlags() & Qt.TextSelectableByMouse
    ev = QMouseEvent(QEvent.MouseButtonDblClick, QPointF(3, 3), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
    app.sendEvent(lbl, ev)
    assert calls[-1][0] == iid and calls[-1][2] == 150000


def test_name_is_selectable_click_toggles_and_copy(win):
    w, db, budget, app = win
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    iid = _open_one_item(w, db, desc='2학기 방과후 강사 수당', spent=0)
    w.show()
    app.processEvents()
    cell = _cell(w, iid)
    name = cell.lbl_name
    assert name.textInteractionFlags() & Qt.TextSelectableByMouse
    assert name.contextMenuPolicy() == Qt.NoContextMenu                   # 우클릭은 줄 메뉴로
    QTest.mouseClick(name, Qt.LeftButton, pos=name.rect().center())        # 그냥 누르면 펼치기
    assert w._expense_ui[iid].get('open') is True
    y = name.height() // 2                                                 # 드래그 → 글자 고르기, 펼침은 그대로
    QTest.mousePress(name, Qt.LeftButton, pos=budget.QPoint(1, y))
    from PyQt5.QtCore import QEvent, QPointF
    from PyQt5.QtGui import QMouseEvent
    app.sendEvent(name, QMouseEvent(QEvent.MouseMove, QPointF(name.width() - 2, y), Qt.NoButton,
                                    Qt.LeftButton, Qt.NoModifier))     # 누른 채로 끌기
    QTest.mouseRelease(name, Qt.LeftButton, pos=budget.QPoint(name.width() - 2, y))
    assert name.hasSelectedText() and w._expense_ui[iid].get('open') is True
    app.clipboard().clear()
    w._copy_desc(iid, selected_only=True)
    assert app.clipboard().text() == name.selectedText() != ''
    from PyQt5.QtGui import QMouseEvent
    ev = QMouseEvent(QEvent.MouseButtonDblClick, QPointF(3, y), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
    app.sendEvent(name, ev)                                                # 더블클릭 → 이름 전체 선택
    assert name.selectedText() == '2학기 방과후 강사 수당'
    app.clipboard().clear()
    QTest.keyClick(name, Qt.Key_C, Qt.ControlModifier)                     # Ctrl+C → 복사
    assert app.clipboard().text() == '2학기 방과후 강사 수당'


def test_expense_row_buttons_next_to_purpose_and_only_on_hover(win):
    w, db, budget, app = win
    from PyQt5.QtCore import QEvent
    iid = _open_one_item(w, db, spent=0)
    db.add_budget_expense(iid, 150000, '체육대회', '2026-09-12')
    w._expense_ui[iid] = {'open': True}
    w.refresh()
    w.show()
    app.processEvents()
    cell = _cell(w, iid)
    row = next(r for r in cell.detail.findChildren(budget._EntryRow))
    assert row.buttons and not any(b.isVisible() for b in row.buttons)          # 평소엔 숨김
    row.enterEvent(QEvent(QEvent.Enter))
    assert all(b.isVisible() for b in row.buttons)                              # 마우스를 올리면
    lay = row.layout()
    order = [lay.itemAt(i) for i in range(lay.count())]
    purpose_i = next(i for i, it in enumerate(order) if it.widget() and it.widget().text() == '체육대회')
    amount_i = next(i for i, it in enumerate(order) if it.widget() and it.widget().text() == '150,000원')
    assert order[purpose_i + 1].layout() is not None and amount_i > purpose_i + 1   # 쓴 곳 바로 오른쪽, 금액은 끝


def test_edit_dialog_labels(win):
    w, db, budget, app = win
    from PyQt5.QtWidgets import QLabel
    item = {'id': 1, 'category_id': 1, 'cost_type': '여비', 'description': '인솔 여비',
            'budget': 1000, 'spent': 500, 'completed': 0}
    dlg = budget._SpendDialog(item, [], w, expense={'id': 3, 'amount': 500, 'purpose': '버스', 'spent_on': '2026-11-01'})
    texts = [lb.text() for lb in dlg.findChildren(QLabel)]
    assert '사용 금액' in texts and '쓴 곳' in texts and '날짜' in texts and '쓴 곳 · 날짜' not in texts
    dlg.close()
