"""예산 창: 집행액 바로 고치기 · 가져오기 방식 · 캡처 붙여넣기 연결."""
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
