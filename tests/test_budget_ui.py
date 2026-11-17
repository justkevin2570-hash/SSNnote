"""예산 창: 집행액 바로 고치기 · 가져오기 방식 · 캡처 붙여넣기 연결."""
import json
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


CATS = [{'name': 'A', 'items': [{'cost_type': '운영수당', 'description': '강사 수당', 'budget': 1000, 'spent': 200}]},
        {'name': 'B', 'items': [{'cost_type': '교육운영비', 'description': '간식', 'budget': 500, 'spent': 0}]}]


def test_import_modes(win):
    w, db, budget, _ = win
    w._apply_import(CATS, 'append')
    assert len(db.get_budget_items()) == 2
    new_a = [{'name': 'A', 'items': [{'cost_type': 'x', 'description': '새 내용', 'budget': 9, 'spent': 1}]}]
    w._apply_import(new_a, 'update')                     # A만 새 내용으로, B는 그대로
    descs = sorted(i['description'] for i in db.get_budget_items())
    assert descs == ['간식', '새 내용']
    w._apply_import(new_a, 'append')
    assert len(db.get_budget_items()) == 3
    w._apply_import(new_a, 'replace')
    assert [c['name'] for c in db.get_budget_categories()] == ['A'] and len(db.get_budget_items()) == 1


def test_columns_name_first_no_spent_and_total_row(win):
    """표 순서: 산출내역(+비목 꼬리표) · 예산현액 · 잔액 · 집행률 (집행액 칸 없음). 세부항목 맨 아래 합계 줄."""
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    table = [t for t in w.findChildren(budget._BudgetTable)][0]
    headers = [table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
    assert headers[:4] == ['산출내역', '예산현액', '잔액', '집행률'] and '집행액' not in headers
    item = next(i for i in db.get_budget_items() if i['description'] == '강사 수당')
    db.set_budget_item_spent(item['id'], 700)
    w.refresh()
    table = [t for t in w.findChildren(budget._BudgetTable)][0]
    assert w.stat['spent'].text() == '700원' and w.stat['remaining'].text() == '800원'
    assert table.item(0, 2).text() == '300' and '700' in table.item(0, 2).toolTip()   # 집행액은 잔액에 마우스를 올리면
    cell = table.cellWidget(0, 0)
    from budget_types import short_name
    assert cell.lbl_name.text() == '강사 수당' and cell.lbl_cost.text() == short_name(item['cost_type'])
    total = w._sec_refs[item['category_id']]['total']
    assert total.item(0, 0).text() == '합계'
    assert total.item(0, 2).text() == _fmt_sum(db, item['category_id'])
    for c in range(3):                                             # 합계 줄 칸이 위 표 칸과 같은 폭
        assert total.columnWidth(c) == table.columnWidth(c)


def _fmt_sum(db, cat_id):
    its = [i for i in db.get_budget_items() if i['category_id'] == cat_id]
    return f"{sum(i['budget'] for i in its) - sum(i['spent'] for i in its):,}"


def test_completed_items_always_shown_below_without_toggle(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    first = db.get_budget_items()[0]
    db.toggle_budget_item_completed(first['id'])
    w.refresh()
    w.show()
    app.processEvents()
    done = [t for t in w.findChildren(budget._BudgetTable)
            if t.horizontalHeader().isHidden() and t.row_items]
    assert len(done) == 1 and done[0].isVisible()                 # 완료 항목은 늘 보임 (접기 단추 없음)
    from PyQt5.QtWidgets import QPushButton
    assert not [b for b in w.findChildren(QPushButton) if '완료한 항목' in b.text()]
    assert done[0].cellWidget(0, 3).findChild(budget.QLabel).text() == '완료'   # 집행률 대신 「완료」


def test_no_row_buttons_hover_tints_row(win):
    """줄 끝 단추([쓴 돈 추가] · 수정 · 완료)는 없음 — 우클릭 메뉴로. 마우스를 올린 줄은 바탕만 옅게."""
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    t = w.findChildren(budget._BudgetTable)[0]
    assert t.columnCount() == 4
    from PyQt5.QtWidgets import QPushButton
    assert not [b for b in w.findChildren(QPushButton) if b.text() == '쓴 돈 추가']
    btns = t.action_buttons[0]
    assert btns and not any(b.isVisible() for b in btns)            # 평소엔 숨김
    t.set_hover_row(0)
    assert t._hover_row == 0 and all(b.isVisible() for b in btns)    # 마우스를 올린 줄에만
    t.set_hover_row(-1)
    assert t._hover_row == -1 and not any(b.isVisible() for b in btns)
    cell = t.cellWidget(0, 0)
    assert '우클릭' in cell.lbl_name.toolTip()


def test_window_blank_right_click_has_no_menu(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtCore import QPoint
    from PyQt5.QtGui import QContextMenuEvent
    shown = []
    monkeypatch.setattr(budget.QMenu, 'exec_', lambda self, *a: shown.append(1))
    w.contextMenuEvent(QContextMenuEvent(QContextMenuEvent.Mouse, QPoint(5, 5), QPoint(5, 5)))
    assert not shown


def test_spend_button_adds_expense(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = next(i['id'] for i in db.get_budget_items() if i['description'] == '강사 수당')

    def fake_exec(self):
        self.edit_spend_now.setText('300')
        self.edit_purpose.setText('11월 강사료')
        self._save()
        return budget.QDialog.Accepted
    monkeypatch.setattr(budget._SpendDialog, 'exec_', fake_exec)
    w._open_spend(iid)
    it = next(i for i in db.get_budget_items() if i['id'] == iid)
    assert it['spent'] == 500 and [(e['purpose'], e['amount']) for e in db.get_budget_expenses()] == [('11월 강사료', 300)]


def test_row_right_click_menu(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    t = w.findChildren(budget._BudgetTable)[0]
    iid = t.row_items[0]
    shown = {}
    monkeypatch.setattr(budget.QMenu, 'exec_', lambda self, *a: shown.setdefault('acts', [a.text() for a in self.actions()]))
    w._row_menu(iid, budget.QPoint(0, 0))
    assert [x for x in shown['acts'] if x] == ['쓴 돈 추가', '수정', '완료', '삭제']
    monkeypatch.setattr(budget.QMessageBox, 'question', lambda *a, **k: budget.QMessageBox.Yes)
    w._delete_item(iid)
    assert iid not in [i['id'] for i in db.get_budget_items()]


def test_pct_label_fits_text(win):
    w, db, budget, app = win
    from PyQt5.QtGui import QFontMetrics
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    lbl = next(iter(w._row_refs.values()))['pct']
    lbl.setText('100.0%')
    assert QFontMetrics(lbl.font()).horizontalAdvance('100.0%') <= lbl.width()


def test_delete_from_edit_dialog(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtWidgets import QMessageBox
    w._apply_import(CATS, 'append')
    item = db.get_budget_items()[0]
    add = budget._ItemDialog(w.categories, None, w.categories[0]['id'], w)
    assert not hasattr(add, 'btn_delete')                  # 추가 창에는 삭제 없음
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: QMessageBox.Yes)
    monkeypatch.setattr(budget._ItemDialog, 'exec_', lambda self: (self.btn_delete.click(), self.result())[1])
    w._open_edit_item(item['id'])
    assert item['id'] not in [i['id'] for i in db.get_budget_items()]


def test_preview_reads_edits_and_default_mode(win):
    w, db, budget, app = win
    dlg = budget._ImportPreviewDialog('test', [dict(c, items=[dict(i) for i in c['items']]) for c in CATS],
                                      True, {'A'}, w)
    dlg.table.item(0, 0).setText('A 바꾼 이름')            # 세부항목 이름 고치기
    dlg.table.item(1, 3).setText('1,100')                 # 집행액 고치기 → 잔액 다시 계산
    assert dlg.table.item(1, 4).text() == '-100'
    dlg._accept()
    assert dlg.mode == 'update'
    assert dlg.categories[0]['name'] == 'A 바꾼 이름' and dlg.categories[0]['items'][0]['spent'] == 1100


def test_paste_routes_image_and_text(win, monkeypatch):
    w, db, budget, app = win
    from PyQt5.QtGui import QImage, QColor
    calls = []
    monkeypatch.setattr(w, '_import_image', lambda img: calls.append(('image', img.width())))
    monkeypatch.setattr(w, '_import_parsed', lambda parsed, src, parent=None: calls.append(('text', len(parsed))))
    img = QImage(40, 20, QImage.Format_RGB32)
    img.fill(QColor('white'))
    app.clipboard().setImage(img)
    w._paste()
    app.clipboard().setText('A열\t산출내역\t1000\t\t200\n')
    from PyQt5.QtWidgets import QMessageBox
    monkeypatch.setattr(QMessageBox, 'warning', lambda *a, **k: calls.append(('warn',)))
    w._paste()
    assert calls[0] == ('image', 40)
    assert calls[1][0] in ('text', 'warn')


def test_spend_dialog_preview_and_result(win):
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': '교육운영비', 'description': '간식비',
            'budget': 400000, 'spent': 115500, 'completed': 0}
    dlg = budget._SpendDialog(item, [], w)
    assert '더해지고' not in dlg.lbl_spend_preview.text() and '284,500원' in dlg.lbl_spend_preview.text()  # 지금 잔액만
    dlg.edit_spend_now.setText('50000')
    assert dlg.edit_spend_now.text() == '50,000'
    text = dlg.lbl_spend_preview.text()
    assert '284,500' in text and '234,500원' in text and '집행' not in text      # 잔액만 (집행액 없음)
    dlg.edit_spend_now.setText('400000')                  # 잔액 마이너스 → 빨강
    assert budget.C.NEG in dlg.lbl_spend_preview.text()
    dlg.edit_spend_now.setText('50000')
    dlg.edit_purpose.setText('간식')
    dlg._save()
    assert dlg.result_data['amount'] == 50000 and dlg.result_data['purpose'] == '간식'


def test_spend_dialog_needs_amount_and_shows_recent(win, monkeypatch):
    from PyQt5.QtWidgets import QMessageBox, QLabel
    w, db, budget, app = win
    warned = []
    monkeypatch.setattr(QMessageBox, 'warning', lambda *a, **k: warned.append(a[2]))
    item = {'id': 1, 'category_id': 1, 'cost_type': 'x', 'description': 'y', 'budget': 1000, 'spent': 300, 'completed': 0}
    recent = [{'spent_on': '2026-11-0%d' % i, 'purpose': f'p{i}', 'amount': 10 * i} for i in range(5, 0, -1)]
    dlg = budget._SpendDialog(item, recent, w)
    dlg._save()
    assert dlg.result_data is None and warned == ['쓴 금액을 적어 주세요.']
    texts = [l.text() for l in dlg.findChildren(QLabel)]
    assert '최근 쓴 돈' in texts and 'p5' in texts and 'p3' in texts and 'p2' not in texts   # 최근 3건만


def test_edit_dialog_has_no_spend_fields_and_keeps_spent(win):
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': 'x', 'description': 'y',
            'budget': 1000, 'spent': 300, 'completed': 0}
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], item, 1, w)
    assert dlg.edit_spend_now is None and dlg.edit_purpose is None       # 쓴 돈은 [쓴 돈 추가] 창에서
    dlg._save()
    assert dlg.result_data['spent'] == 300 and dlg.result_data['expense'] is None


def test_add_dialog_has_no_spend_now(win):
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    assert dlg.edit_spend_now is None


def test_cost_chips_totals_and_filter(win):
    w, db, budget, app = win
    cats = [{'name': 'A', 'items': [
        {'cost_type': '교육운영비', 'description': '간식', 'budget': 1000, 'spent': 400},
        {'cost_type': '운영수당', 'description': '강사', 'budget': 2000, 'spent': 0},
        {'cost_type': '교육운영비', 'description': '자료', 'budget': 500, 'spent': 500}]}]
    w._apply_import(cats, 'append')
    w.show()
    app.processEvents()
    assert '교육운영비' in w._overview_chips                      # '전체'일 때 맨 위 한 줄
    assert w._overview_chips['교육운영비'].isVisible()
    assert w._overview_chips['교육운영비'].text().endswith('교육운영  60%')   # 900 / 1500 (줄임말 + 아이콘)
    assert '<img' in w._overview_chips['교육운영비'].text()
    cid = w.categories[0]['id']
    chips = w._sec_refs[cid]['chips']
    assert list(chips) == ['교육운영비', '운영수당']
    assert chips['교육운영비'].text() == '교육운영  60%' and not chips['교육운영비'].icon().isNull() and '잔액 600원' in chips['교육운영비'].toolTip()
    table = w._sec_refs[cid]['tables'][0]
    h_all = table.height()
    chips['운영수당'].click()                                    # 운영수당 줄만
    assert [table.isRowHidden(2 * i) for i in range(3)] == [True, False, True]   # 항목 줄 (2칸씩 — 사이에 사용 내역 줄)
    assert table.height() < h_all
    chips['운영수당'].click()                                    # 다시 누르면 전부
    assert not any(table.isRowHidden(2 * i) for i in range(3))
    iid = next(i['id'] for i in db.get_budget_items() if i['description'] == '간식')
    db.set_budget_item_spent(iid, 900)                          # 집행액이 바뀌면 칩도 다시 계산
    w.refresh()
    assert '잔액 100원' in w._sec_refs[cid]['chips']['교육운영비'].toolTip()


def test_item_dialog_lists_cost_types(win):
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    names = [dlg.edit_cost.itemText(i) for i in range(dlg.edit_cost.count())]
    assert '교육운영비' in names and '목적사업업무추진비' in names


class _FakeAdapter:
    def __init__(self, reply):
        self.reply = reply
        self.prompts = []

    @staticmethod
    def get_api_key():
        return 'k'

    def complete(self, system, prompt, image_b64=None, max_tokens=0):
        self.prompts.append(prompt)
        return self.reply(prompt)


def _ai(monkeypatch, reply):
    import ai_client
    import budget_ai
    import school_profile
    fake = _FakeAdapter(reply)
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: fake)
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'claude')
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: False)
    monkeypatch.setattr(budget_ai.BudgetAskWorker, 'start', lambda self: self.run())   # 스레드 없이
    import ai_actions
    monkeypatch.setattr(ai_actions.AskWorker, 'start', lambda self: self.run())
    return fake


def test_ask_answers_with_budget_context(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    fake = _ai(monkeypatch, lambda p: '{"answer": "A의 잔액은 800원이에요.", "actions": []}')
    w.refresh()
    assert w.ask_card.isVisibleTo(w) and w.ask_input.isEnabled() and w.quick_buttons == []
    w.ask_input.setText('A 얼마 남았어?')
    w.ask_send.click()
    assert '강사 수당' in fake.prompts[0] and 'task.delete' in fake.prompts[0]   # 예산 표 + 앱 전체 기능 목록
    assert 'A의 잔액은 800원' in w.ask_answer.toPlainText()
    assert w.ask_confirm.isHidden()
    w._ask('그중 제일 큰 건?')
    assert len(fake.prompts) == 2 and '이전 질문: A 얼마 남았어?' in fake.prompts[1]


def test_spoken_spending_needs_confirm(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    item = next(i for i in db.get_budget_items() if i['description'] == '간식')
    _ai(monkeypatch, lambda p: json.dumps({'answer': '간식에 300원을 넣을까요?',
                                           'actions': [{'do': 'budget.spend', 'item_id': item['id'], 'amount': 300}]},
                                          ensure_ascii=False))
    w.refresh()
    w._ask('간식 300원 썼어')
    assert not w.ask_confirm.proposal.isHidden() and '300원 썼어요' in w.ask_confirm.lbl_proposal.text()
    assert next(i for i in db.get_budget_items() if i['id'] == item['id'])['spent'] == 0   # 아직 안 넣음
    w.ask_confirm.apply()
    assert next(i for i in db.get_budget_items() if i['id'] == item['id'])['spent'] == 300
    assert w.ask_confirm.proposal.isHidden() and '1건을 했어요' in w.ask_status.text()


def test_ask_disabled_when_ai_off(win, monkeypatch):
    w, db, budget, app = win
    import ai_client
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    w._apply_import(CATS, 'append')
    assert not w.ask_send.isEnabled() and w.ask_input.isEnabled() and 'AI가 꺼져' in w.ask_input.placeholderText()


# ── 항목 창: Enter는 저장, 저장 왼쪽 · 취소 오른쪽 ──

_ITEM = {'id': 1, 'category_id': 1, 'cost_type': '교육운영비', 'description': '간식비',
         'budget': 400000, 'spent': 115500, 'completed': 0}


def _press_enter(app, dlg, field):
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    dlg.show()
    app.processEvents()
    field.setFocus()
    app.processEvents()
    QTest.keyClick(field, Qt.Key_Return)
    app.processEvents()


@pytest.mark.parametrize('name', ['edit_desc', 'edit_budget', 'edit_spent', 'edit_cost'])
def test_enter_saves_item_dialog(win, name):
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], dict(_ITEM), 1, w)
    _press_enter(app, dlg, getattr(dlg, name))
    assert dlg.result() == 1 and dlg.result_data and dlg.result_data['description'] == '간식비'   # 취소가 아니라 저장
    assert not dlg.deleted
    dlg.close()


def test_enter_in_add_dialog_saves_only_with_description(win, monkeypatch):
    from PyQt5.QtWidgets import QMessageBox
    w, db, budget, app = win
    warned = []
    monkeypatch.setattr(QMessageBox, 'warning', lambda *a, **k: warned.append(a[2]))
    empty = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    _press_enter(app, empty, empty.edit_budget)
    assert empty.result() == 0 and empty.result_data is None and warned == ['산출내역(내용)을 입력해주세요.']
    empty.close()
    filled = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    filled.edit_desc.setText('새 항목')
    _press_enter(app, filled, filled.edit_desc)
    assert filled.result() == 1 and filled.result_data['description'] == '새 항목'
    filled.close()


def test_enter_never_deletes_or_cancels(win, monkeypatch):
    from PyQt5.QtWidgets import QMessageBox
    w, db, budget, app = win
    asked = []
    monkeypatch.setattr(QMessageBox, 'question', lambda *a, **k: asked.append(1) or QMessageBox.No)
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], dict(_ITEM), 1, w)
    assert dlg.btn_save.isDefault()
    assert not dlg.btn_cancel.autoDefault() and not dlg.btn_delete.autoDefault()
    _press_enter(app, dlg, dlg.edit_desc)
    assert asked == [] and not dlg.deleted
    dlg.close()


def test_cost_type_is_pick_only_from_list(win, monkeypatch):
    """원가통계비목은 목록에서 고르기만: (비목 없음) · 자주 쓰는 비목 · 나머지 · 지금 쓰는 비목 · 목록 관리…"""
    from PyQt5.QtWidgets import QComboBox
    import budget_types as bt
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], dict(_ITEM, cost_type='시설장비유지비'), 1, w)
    cb = dlg.edit_cost
    assert isinstance(cb, QComboBox) and not cb.isEditable()
    texts = [cb.itemText(i) for i in range(cb.count())]
    assert texts[0] == '(비목 없음)' and texts[1:1 + len(bt.PRIORITY)] == list(bt.PRIORITY)
    assert texts[-1] == '목록 관리…' and '기간제근로자법정부담금' in texts
    assert cb.currentText() == '시설장비유지비'                     # 예전 비목도 그대로 골라져 있음
    cb.setCurrentIndex(cb.findData('여비'))
    dlg._save()
    assert dlg.result_data['cost_type'] == '여비'
    dlg2 = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    assert dlg2.edit_cost.currentIndex() == 0
    opened = []
    monkeypatch.setattr(budget._CostTypeDialog, 'exec_', lambda self: opened.append(1) or 0)
    dlg2.edit_cost.setCurrentIndex(dlg2.edit_cost.findData('운영수당'))
    dlg2._on_cost_picked(dlg2.edit_cost.currentIndex())
    manage = dlg2.edit_cost.findData(budget._ItemDialog._MANAGE)
    dlg2._on_cost_picked(manage)
    assert opened and dlg2.cost_value() == '운영수당'                 # 목록 관리 뒤에도 고른 비목 그대로
    dlg.close(); dlg2.close()


def test_save_left_cancel_right(win):
    w, db, budget, app = win
    pairs = []
    item_dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], dict(_ITEM), 1, w)
    pairs.append((item_dlg, item_dlg.btn_save, item_dlg.btn_cancel))
    cats = [{'name': 'A', 'items': [{'cost_type': 'x', 'description': 'd', 'budget': 1, 'spent': 0}]}]
    preview = budget._ImportPreviewDialog('test', cats, False, (), w)
    pairs.append((preview, preview.btn_import, preview.btn_cancel))
    for dlg, left, right in pairs:
        dlg.show()
        app.processEvents()
        assert left.mapTo(dlg, left.rect().topLeft()).x() < right.mapTo(dlg, right.rect().topLeft()).x()
        assert left.isDefault() and not right.autoDefault()
        dlg.close()


def _item_id(db, desc):
    return next(i['id'] for i in db.get_budget_items() if i['description'] == desc)


def test_target_balance_apply_and_undo(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item_id(db, '강사 수당')                          # 예산 1,000 · 집행 200
    _ai(monkeypatch, lambda p: json.dumps({'answer': '강사 수당 잔액을 300원으로 맞출게요.',
                                           'actions': [{'do': 'budget.set_remaining', 'item_id': iid, 'remaining': 300}]},
                                          ensure_ascii=False))
    w.refresh()
    w._ask('강사 수당 잔액 300원으로 맞춰줘')
    text = w.ask_confirm.lbl_proposal.text()
    assert '이렇게 할까요' in text and '집행 700원' in text and '잔액 300원' in text   # 앱이 계산
    w.ask_confirm.apply()
    assert _spent(db, iid) == 700 and not w.btn_undo_apply.isHidden()
    w._undo_apply()
    assert _spent(db, iid) == 200 and w.btn_undo_apply.isHidden()


def test_rejected_suggestion_is_grey_and_cannot_be_applied(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item_id(db, '강사 수당')
    _ai(monkeypatch, lambda p: json.dumps({'answer': '잔액을 맞출게요.',
                                           'actions': [{'do': 'budget.set_remaining', 'item_id': iid,
                                                        'remaining': 99999}]}))
    w.refresh()
    w._ask('잔액 99999원으로')
    text = w.ask_confirm.lbl_proposal.text()
    assert '강사 수당' in text and '예산보다' in text and '할 수 없어요' in text
    assert w.ask_confirm.btn_apply.isHidden()
    w.ask_confirm.apply()
    assert _spent(db, iid) == 200


def test_budget_window_can_delete_and_add_tasks_too(win, monkeypatch):
    """예산 창에서도 앱 전체 기능: 예산 항목 지우기는 빨간 칸 [지우기]."""
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item_id(db, '간식')
    _ai(monkeypatch, lambda p: json.dumps({'answer': '간식 항목을 지울게요.',
                                           'actions': [{'do': 'budget.item_delete', 'item_id': iid}]},
                                          ensure_ascii=False))
    w.refresh()
    w._ask('간식 항목 지워줘')
    assert not w.ask_confirm.proposal_del.isHidden() and w.ask_confirm.proposal.isHidden()
    assert any(i['id'] == iid for i in db.get_budget_items())
    w.ask_confirm.apply_deletes()
    assert not any(i['description'] == '간식' for i in db.get_budget_items())
    w._undo_apply()
    assert any(i['description'] == '간식' for i in db.get_budget_items())


def _spent(db, iid):
    return next(i for i in db.get_budget_items() if i['id'] == iid)['spent']


# ── 전체 합계에서 빼기 · 세부항목 그룹 ─────────────────────────────
def _cat_id(db, name):
    return next(c['id'] for c in db.get_budget_categories() if c['name'] == name)


def test_excluded_category_left_out_of_totals(win):
    w, db, budget, _ = win
    w._apply_import(CATS, 'replace')
    w._set_excluded([_cat_id(db, 'B')], True)
    assert w.stat['budget'].text() == '1,000원'               # B(500)는 빠짐
    assert w.stat['spent'].text() == '200원'
    assert '합계 제외 1개' in w.lbl_meta.text()
    assert sum(i['budget'] for i in db.get_budget_items_in_total()) == 1000
    import dashboard
    s = dashboard.build_summary([], db.get_budget_items_in_total())
    assert s['budget']['allocated'] == 1000
    w._set_excluded([_cat_id(db, 'B')], False)
    assert w.stat['budget'].text() == '1,500원'


def test_ai_context_marks_excluded_and_groups(win):
    w, db, budget, _ = win
    import budget_ai
    w._apply_import(CATS, 'replace')
    gid = db.add_budget_group('묶음')
    db.set_budget_category_group(_cat_id(db, 'A'), gid)
    db.set_budget_category_group(_cat_id(db, 'B'), gid)
    db.set_budget_category_excluded(_cat_id(db, 'B'), True)
    ctx = budget_ai.build_context(db.get_budget_categories(), db.get_budget_items(), db.get_budget_groups())
    assert '[세부항목] B (전체 합계에서 제외)' in ctx
    assert '[그룹] 묶음 (A, B) — 예산 1,500 · 집행 200' in ctx      # 그룹 합은 제외 여부와 상관없이
    assert ctx.splitlines()[-1].endswith('예산 1,000 · 집행 200 · 잔액 800')


def test_group_nav_totals_and_tab(win):
    w, db, budget, _ = win
    w._apply_import(CATS + [{'name': 'C', 'items': [{'cost_type': 'x', 'description': 'y',
                                                    'budget': 100, 'spent': 100}]}], 'replace')
    gid = db.add_budget_group('묶음')
    w._move_to_group(_cat_id(db, 'A'), gid)
    w._move_to_group(_cat_id(db, 'B'), gid)
    key = ('G', gid)
    assert key in w._nav_refs
    assert '잔액 1,300원' in w._nav_refs[key].lbl_meta.text()
    w._set_tab(key)
    assert set(w._sec_refs) == {_cat_id(db, 'A'), _cat_id(db, 'B')}
    w._fold_group(gid, True)                                   # 접으면 안의 세부항목 줄은 숨김
    assert _cat_id(db, 'A') not in w._nav_refs and key in w._nav_refs
    w._delete_group(gid)                                       # 그룹을 풀어도 세부항목은 남음
    assert len(db.get_budget_categories()) == 3 and w.active_tab == 'ALL'
    assert all(c['group_id'] is None for c in db.get_budget_categories())


def test_reimport_keeps_exclusion_and_group(win):
    w, db, budget, _ = win
    w._apply_import(CATS, 'replace')
    gid = db.add_budget_group('묶음')
    db.set_budget_category_group(_cat_id(db, 'A'), gid)
    db.set_budget_category_excluded(_cat_id(db, 'B'), True)
    w._apply_import(CATS, 'replace')
    cats = {c['name']: c for c in db.get_budget_categories()}
    assert cats['A']['group_id'] == gid and cats['B']['excluded'] == 1
    db.clear_budget_data()
    assert db.get_budget_groups() == []


def test_add_item_from_group_tab_defaults_to_group_member(win, monkeypatch):
    w, db, budget, _ = win
    w._apply_import(CATS, 'replace')
    gid = db.add_budget_group('묶음')
    w._move_to_group(_cat_id(db, 'B'), gid)
    w._set_tab(('G', gid))
    seen = {}

    class Fake:
        def __init__(self, cats, item, default_cat_id, parent=None):
            seen['cat'] = default_cat_id

        def exec_(self):
            return 0
    monkeypatch.setattr(budget, '_ItemDialog', Fake)
    w._open_add_item()
    assert seen['cat'] == _cat_id(db, 'B')


# ── 예산 창 AI 연결 버튼 (공문 작성 창과 같은 메뉴 · 같은 설정) ──────────
def test_ai_button_shows_status_and_follows_changes(win, monkeypatch, tmp_path):
    w, db, budget, _ = win
    import ai_client
    from ui import ai_connect
    monkeypatch.setattr(ai_client, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai_client, 'AI_MODE_FILE', str(tmp_path / 'ai_mode.txt'))
    monkeypatch.setattr(ai_client, '_read_mode_file', lambda: open(tmp_path / 'ai_mode.txt').read().strip()
                        if (tmp_path / 'ai_mode.txt').exists() else 'none')
    monkeypatch.setattr(ai_connect, 'load_external_model_name', lambda: 'sonnet')
    ai_connect.turn_off()
    assert w.btn_ai.text() == 'AI 연결하기'                    # 꺼져 있으면 누를 수 있는 단추로
    assert not w.ask_send.isEnabled()

    ai_client.save_ai_mode('claude_code')                     # 다른 창(공문 작성)에서 연결했다고 치고
    ai_connect.notify_changed()
    assert w.btn_ai.text() == 'Claude 구독 · sonnet'
    assert w.ask_send.isEnabled()

    monkeypatch.setattr(ai_connect, 'show_menu', lambda *a, **k: 'none')   # 메뉴에서 [AI 끄기]
    w._ai_menu()
    assert ai_client.load_ai_mode() == 'none' and w.btn_ai.text() == 'AI 연결하기'

    seen = {}
    monkeypatch.setattr(ai_connect, 'show_menu', lambda *a, **k: 'openai')
    monkeypatch.setattr(ai_connect, 'connect', lambda parent, mode, say=None: seen.setdefault('mode', mode))
    w._ai_menu()
    assert seen['mode'] == 'openai'


# ── 화면 배치 기억 (보던 칸 · 펼침 · 비목 · 스크롤 · 창 크기) ─────────────
def test_budget_view_state_remembered_and_reset(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'replace')
    db.add_budget_item(_cat_id(db, 'A'), '교육운영비', '자료', 300, 0)   # 비목이 둘 이상이어야 비목 칩이 생김
    gid = db.add_budget_group('묶음')
    db.set_budget_category_group(_cat_id(db, 'A'), gid)
    w.refresh()
    w._state_ready = True
    w._set_tab(('G', gid))
    w._cost_filter = {_cat_id(db, 'A'): '운영수당'}
    w._show_done = {_cat_id(db, 'B')}
    w._save_state()
    saved = json.load(open(budget._STATE_FILE, encoding='utf-8'))
    assert saved['view']['tab'] == ['G', gid] and saved['view']['show_done'] == [_cat_id(db, 'B')]
    w2 = budget.BudgetWindow()                                     # 다음에 켰을 때
    assert w2.active_tab == ('G', gid)
    assert w2._cost_filter == {_cat_id(db, 'A'): '운영수당'} and w2._show_done == {_cat_id(db, 'B')}
    w2._state_ready = True
    w2.reset_layout()                                              # 기본값으로
    assert w2.active_tab == 'ALL' and not w2._show_done and not any(w2._cost_filter.values())
    assert json.load(open(budget._STATE_FILE, encoding='utf-8'))['view']['tab'] == 'ALL'
    w2.close()


def test_budget_state_not_saved_before_restore_and_deleted_tab_falls_back(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'replace')
    json.dump({'x': 10, 'y': 10, 'width': 900, 'height': 600,
               'view': {'tab': 999999, 'overview_open': True}}, open(budget._STATE_FILE, 'w'))
    w2 = budget.BudgetWindow()
    assert w2.active_tab == 'ALL'                                  # 지운 세부항목이면 전체로
    w2._save_state()                                               # 아직 다 그리기 전: 덮어쓰지 않음
    assert json.load(open(budget._STATE_FILE))['view']['overview_open'] is True
    w2.close()


def test_budget_ask_clear_command(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'replace')
    w._ask_history = [('질문', '답')]
    w._set_answer('Q. 질문\n\n답')
    started = []
    import budget_ai
    monkeypatch.setattr(budget_ai.BudgetAskWorker, 'start', lambda self: started.append(1))
    w._ask('/CLEAR ')
    assert w._ask_history == [] and w.ask_answer.isHidden() and not started
    w._ask('/모름')
    assert '모르는 명령' in w.ask_status.text() and not started


def test_opening_window_shows_no_stray_small_windows(win):
    """예산 창을 열 때 칸 속 글자(비목 꼬리표 등)가 따로 작은 창으로 뜨지 않는다."""
    w, db, budget, app = win
    from PyQt5.QtCore import QObject, QEvent
    shown = []

    class Spy(QObject):
        def eventFilter(self, o, e):
            if e.type() == QEvent.Show and o.isWidgetType() and o.isWindow():
                shown.append(type(o).__name__)
            return False
    spy = Spy()
    app.installEventFilter(spy)
    try:
        w._apply_import(CATS, 'append')
        first = db.get_budget_items()[0]
        db.toggle_budget_item_completed(first['id'])
        w._show_done = {first['category_id']}                 # 완료 항목을 펼쳐 둔 세부항목도
        w.refresh()
        w.show()
        app.processEvents()
        w.refresh()                                           # 열린 창에서 다시 그려도
        app.processEvents()
    finally:
        app.removeEventFilter(spy)
    assert set(shown) <= {'BudgetWindow'}


def test_spend_preview_balance_only_without_growing(win):
    """쓴 돈 미리보기는 잔액만 한 줄로, 금액을 적어도 창 높이가 바뀌지 않는다."""
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': '운영비', 'description': '강사 수당',
            'budget': 12000000, 'spent': 9800000, 'completed': 0}
    dlg = budget._SpendDialog(item, [], w)
    dlg.show()
    app.processEvents()
    h0 = dlg.height()
    dlg.edit_spend_now.setText('2500000')
    app.processEvents()
    lbl = dlg.lbl_spend_preview
    assert '-300,000원' in lbl.text() and '집행' not in lbl.text()
    assert lbl.height() >= lbl.sizeHint().height() and lbl.width() >= lbl.sizeHint().width()   # 잘리지 않음
    assert dlg.height() == h0
    dlg.close()


def test_amount_cursor_stays_where_you_edit(win):
    """금액 칸 가운데를 고쳐도 커서가 끝으로 튀지 않는다 (쉼표가 생기거나 없어져도)."""
    w, db, budget, app = win
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    item = {'id': 1, 'category_id': 1, 'cost_type': 'x', 'description': 'y', 'budget': 1000, 'spent': 0, 'completed': 0}
    dlg = budget._SpendDialog(item, [], w)
    e = dlg.edit_spend_now
    assert not (e.alignment() & Qt.AlignRight)                      # 왼쪽 정렬
    QTest.keyClicks(e, '300000')
    assert e.text() == '300,000' and e.cursorPosition() == len('300,000')
    e.setCursorPosition(1)                                          # 3|00,000 에서
    QTest.keyClick(e, Qt.Key_5)                                     # 5 넣기 → 3,500,000
    assert e.text() == '3,500,000' and e.text()[:e.cursorPosition()] == '3,5'
    QTest.keyClick(e, Qt.Key_Backspace)                             # 다시 지우면 300,000
    assert e.text() == '300,000' and e.text()[:e.cursorPosition()] == '3'
    dlg.close()


def test_row_icon_buttons_spend_and_done_only(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    first = db.get_budget_items()[0]
    db.toggle_budget_item_completed(first['id'])
    w.refresh()
    w.show()
    app.processEvents()
    open_items = [i for i in db.get_budget_items() if not i['completed']]
    cell = w._row_refs[open_items[0]['id']]['desc']
    assert list(cell.row_buttons) == ['spend', 'done'] and not hasattr(cell, 'btn_copy')
    assert [cell.row_buttons[k].toolTip() for k in ('spend', 'done')] == ['쓴 돈 추가', '완료']
    calls = []
    monkeypatch.setattr(w, '_open_spend', lambda iid: calls.append(('spend', iid)))
    monkeypatch.setattr(w, '_toggle_complete', lambda iid: calls.append(('done', iid)))
    w.refresh()
    cell = w._row_refs[open_items[0]['id']]['desc']
    cell.row_buttons['spend'].click()
    cell.row_buttons['done'].click()
    assert calls == [('spend', open_items[0]['id']), ('done', open_items[0]['id'])]
    done_cell = w._row_refs[first['id']]['desc']
    assert list(done_cell.row_buttons) == ['done'] and done_cell.row_buttons['done'].toolTip() == '완료 취소'
