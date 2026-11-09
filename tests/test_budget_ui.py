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


def test_spent_cell_edit_saves_and_updates_totals(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    table = [t for t in w.findChildren(budget._BudgetTable)][0]
    cell = table.item(0, budget._COL_SPENT)
    cell.setText('700')                                   # 셀에서 고친 것처럼
    item = next(i for i in db.get_budget_items() if i['description'] == '강사 수당')
    assert item['spent'] == 700
    assert cell.text() == '700'
    assert w.stat['spent'].text() == '700원' and w.stat['remaining'].text() == '800원'
    assert table.item(0, 4).text() == '300'


def test_completed_items_fold_away(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    db.toggle_budget_item_completed(db.get_budget_items()[0]['id'])
    w.refresh()
    w.show()
    app.processEvents()
    hidden = [t for t in w.findChildren(budget._BudgetTable) if t.horizontalHeader().isHidden()]
    assert len(hidden) == 1 and not hidden[0].isVisible()   # 완료 표는 접혀 있음


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


def test_item_dialog_spend_now_adds_to_spent(win):
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': '교육운영비', 'description': '간식비',
            'budget': 400000, 'spent': 115500, 'completed': 0}
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], item, 1, w)
    assert '더해지고' in dlg.lbl_spend_preview.text()
    dlg.edit_spend_now.setText('50000')
    assert dlg.edit_spend_now.text() == '50,000'
    assert '165,500원' in dlg.lbl_spend_preview.text() and '234,500원' in dlg.lbl_spend_preview.text()
    dlg.edit_spend_now.setText('400000')                  # 잔액 마이너스 → 빨강
    assert budget.C.NEG in dlg.lbl_spend_preview.text()
    dlg.edit_spend_now.setText('50000')
    dlg._save()
    assert dlg.result_data['spent'] == 165500 and dlg.result_data['budget'] == 400000


def test_item_dialog_spend_now_blank_keeps_spent(win):
    w, db, budget, app = win
    item = {'id': 1, 'category_id': 1, 'cost_type': 'x', 'description': 'y',
            'budget': 1000, 'spent': 300, 'completed': 0}
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], item, 1, w)
    dlg._save()
    assert dlg.result_data['spent'] == 300


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
    cid = w.categories[0]['id']
    chips = w._sec_refs[cid]['chips']
    assert list(chips) == ['교육운영비', '운영수당']
    assert '잔액 600원' in chips['교육운영비'].text()              # 1500 - 900
    assert '교육운영비' in w._overview_chips                      # '전체'일 때 맨 위 카드
    table = w._sec_refs[cid]['tables'][0]
    h_all = table.height()
    chips['운영수당'].click()                                    # 운영수당 줄만
    assert [table.isRowHidden(r) for r in range(3)] == [True, False, True]
    assert table.height() < h_all
    chips['운영수당'].click()                                    # 다시 누르면 전부
    assert not any(table.isRowHidden(r) for r in range(3))
    table.item(0, budget._COL_SPENT).setText('900')             # 집행액 고치면 칩도 다시 계산
    assert '잔액 100원' in chips['교육운영비'].text()


def test_item_dialog_has_cost_completer(win):
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], None, 1, w)
    model = dlg._cost_completer.model()
    names = [model.index(i, 0).data() for i in range(model.rowCount())]
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
    return fake


def test_ask_answers_with_budget_context(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    fake = _ai(monkeypatch, lambda p: '{"answer": "A의 잔액은 800원이에요.", "spend": []}')
    w.refresh()
    assert w.ask_card.isVisibleTo(w) and w.ask_input.isEnabled()
    w.ask_input.setText('A 얼마 남았어?')
    w.ask_send.click()
    assert '강사 수당' in fake.prompts[0]                     # 예산 표를 함께 보냄
    assert 'A의 잔액은 800원' in w.ask_answer.toPlainText()
    assert w.ask_proposal.isHidden()
    w.quick_buttons[0].click()                               # 빠른 질문 버튼
    assert len(fake.prompts) == 2 and '이전 질문: A 얼마 남았어?' in fake.prompts[1]


def test_spoken_spending_needs_confirm(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    item = next(i for i in db.get_budget_items() if i['description'] == '간식')
    _ai(monkeypatch, lambda p: json.dumps({'answer': '간식에 300원을 넣을까요?',
                                           'spend': [{'item_id': item['id'], 'amount': 300}]}, ensure_ascii=False))
    w.refresh()
    w._ask('간식 300원 썼어')
    assert not w.ask_proposal.isHidden() and '+300원' in w.lbl_proposal.text()
    assert next(i for i in db.get_budget_items() if i['id'] == item['id'])['spent'] == 0   # 아직 안 넣음
    w._apply_spend()
    assert next(i for i in db.get_budget_items() if i['id'] == item['id'])['spent'] == 300
    assert w.ask_proposal.isHidden() and '넣었어요' in w.ask_status.text()


def test_ask_disabled_when_ai_off(win, monkeypatch):
    w, db, budget, app = win
    import ai_client
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    w._apply_import(CATS, 'append')
    assert not w.ask_input.isEnabled() and 'AI가 꺼져' in w.ask_input.placeholderText()
