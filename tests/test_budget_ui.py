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


def test_spent_column_hidden_and_totals_follow(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    table = [t for t in w.findChildren(budget._BudgetTable)][0]
    headers = [table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
    assert '집행액' not in headers and headers[3] == '잔액'      # 집행액 열은 없음
    item = next(i for i in db.get_budget_items() if i['description'] == '강사 수당')
    db.set_budget_item_spent(item['id'], 700)
    w.refresh()
    table = [t for t in w.findChildren(budget._BudgetTable)][0]
    assert w.stat['spent'].text() == '700원' and w.stat['remaining'].text() == '800원'
    assert table.item(0, 3).text() == '300'


def test_completed_items_shown_by_default_and_can_fold(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    db.toggle_budget_item_completed(db.get_budget_items()[0]['id'])
    w.refresh()
    w.show()
    app.processEvents()
    done = [t for t in w.findChildren(budget._BudgetTable) if t.horizontalHeader().isHidden()]
    assert len(done) == 1 and done[0].isVisible()            # 완료 항목은 펼친 것이 기본
    from PyQt5.QtWidgets import QPushButton
    toggle = next(b for b in w.findChildren(QPushButton) if '완료 1건' in b.text())
    assert toggle.text().startswith('▾')
    toggle.click()
    assert not done[0].isVisible() and toggle.text().startswith('▸')


def test_row_buttons_edit_left_done_right_always_visible(win):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    w.show()
    app.processEvents()
    t = w.findChildren(budget._BudgetTable)[0]
    edit, done = t.action_buttons[0]
    assert (edit.text(), done.text()) == ('수정', '완료')
    assert edit.isVisible() and done.isVisible()
    assert edit.mapTo(t, edit.rect().topLeft()).x() < done.mapTo(t, done.rect().topLeft()).x()
    assert edit.height() >= 28 and edit.width() >= 64


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
    assert '교육운영비' in w._overview_chips                      # '전체'일 때 맨 위 카드
    assert not w._overview_chips['교육운영비'].isVisible()        # 기본은 접힘
    w._overview_open = True
    w._rebuild_sections()
    app.processEvents()
    assert w._overview_chips['교육운영비'].isVisible()            # 펼치면 보임
    w._overview_open = False
    w._rebuild_sections()
    cid = w.categories[0]['id']
    chips = w._sec_refs[cid]['chips']
    assert list(chips) == ['교육운영비', '운영수당']
    assert '잔액 600원' in chips['교육운영비'].text()              # 1500 - 900
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
    assert '잔액 100원' in w._sec_refs[cid]['chips']['교육운영비'].text()


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


@pytest.mark.parametrize('name', ['edit_spend_now', 'edit_desc', 'edit_budget', 'edit_spent', 'edit_cost'])
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


def test_completer_enter_picks_cost_type_without_saving(win):
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    w, db, budget, app = win
    dlg = budget._ItemDialog([{'id': 1, 'name': 'A'}], dict(_ITEM, cost_type=''), 1, w)
    dlg.show()
    app.processEvents()
    dlg.edit_cost.setFocus()
    dlg._cost_completer.complete()
    app.processEvents()
    pop = dlg._cost_completer.popup()
    pop.setCurrentIndex(pop.model().index(0, 0))
    QTest.keyClick(pop, Qt.Key_Return)
    app.processEvents()
    assert dlg.edit_cost.text() and dlg.result() == 0 and dlg.result_data is None   # 목록에서 고르기만, 저장은 안 됨
    dlg.close()


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


def test_target_balance_shows_table_then_apply_and_undo(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item_id(db, '강사 수당')                          # 예산 1,000 · 집행 200
    _ai(monkeypatch, lambda p: json.dumps({'answer': '강사 수당 잔액을 300원으로 맞출게요.',
                                           'set_remaining': [{'item_id': iid, 'remaining': 300}]}, ensure_ascii=False))
    w.refresh()
    w._ask('강사 수당 잔액 300원으로 맞춰줘')
    text = w.lbl_proposal.text()
    assert not w.ask_proposal.isHidden() and '이렇게 바꿀까요' in text and '실제 집행 내역과 맞는지' in text
    assert '200' in text and '700' in text                  # 집행액 200 → 700 (앱이 계산)
    w._apply_spend()
    assert _spent(db, iid) == 700
    assert not w.btn_undo_apply.isHidden() and '1개 항목' in w.ask_status.text()
    w._undo_apply()
    assert _spent(db, iid) == 200 and w.btn_undo_apply.isHidden()


def test_rejected_suggestion_is_grey_and_cannot_be_applied(win, monkeypatch):
    w, db, budget, app = win
    w._apply_import(CATS, 'append')
    iid = _item_id(db, '강사 수당')
    _ai(monkeypatch, lambda p: json.dumps({'answer': '잔액을 맞출게요.',
                                           'set_remaining': [{'item_id': iid, 'remaining': 99999}]}))
    w.refresh()
    w._ask('잔액 99999원으로')
    assert '제외: 강사 수당' in w.lbl_proposal.text() and '예산보다' in w.lbl_proposal.text()
    assert not w.btn_apply_spend.isEnabled()
    w._apply_spend()
    assert _spent(db, iid) == 200


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
    assert w.btn_ai.text() == 'AI 꺼짐 · 연결하기'
    assert not w.ask_input.isEnabled()

    ai_client.save_ai_mode('claude_code')                     # 다른 창(공문 작성)에서 연결했다고 치고
    ai_connect.notify_changed()
    assert w.btn_ai.text() == 'Claude 구독 · sonnet'
    assert w.ask_input.isEnabled()

    monkeypatch.setattr(ai_connect, 'show_menu', lambda *a, **k: 'none')   # 메뉴에서 [AI 끄기]
    w._ai_menu()
    assert ai_client.load_ai_mode() == 'none' and w.btn_ai.text() == 'AI 꺼짐 · 연결하기'

    seen = {}
    monkeypatch.setattr(ai_connect, 'show_menu', lambda *a, **k: 'openai')
    monkeypatch.setattr(ai_connect, 'connect', lambda parent, mode, say=None: seen.setdefault('mode', mode))
    w._ai_menu()
    assert seen['mode'] == 'openai'
