"""원가통계비목 목록과 교정 (budget_types)."""
import pytest

import budget_types as bt

NAMES = list(bt.DEFAULTS)


@pytest.mark.parametrize('text, expected, changed, ok', [
    ('교육운영비', '교육운영비', False, True),
    ('교육 운영비', '교육운영비', True, True),             # 띄어쓰기만 다름
    ('교육운염비', '교육운영비', True, True),              # 한 글자 오타
    ('운영수딩', '운영수당', True, True),
    ('기간제근로자법정부담굼', '기간제근로자법정부담금', True, True),
    ('기간제교원법정부담금', '기간제교원법정부담금', False, True),   # 목록에 있는 다른 비목은 그대로
    ('시설장비유지비', '시설장비유지비', False, False),   # 목록에 없음 → 그대로
    ('여', '여', False, False),
])
def test_correct(text, expected, changed, ok):
    assert bt.correct(text, NAMES) == (expected, changed, ok)


def test_similar_but_different_types_are_not_merged():
    # 기간제교원 ↔ 기간제근로자: 여러 글자가 달라 다른 비목
    assert bt.correct('기간제교원법정부담금', ['기간제근로자법정부담금'])[0] == '기간제교원법정부담금'


def test_defaults_have_no_duplicates():
    assert len(set(bt.DEFAULTS)) == len(bt.DEFAULTS)
    assert '목적사업업무추진비' in bt.DEFAULTS and '기간제교원법정부담금' in bt.DEFAULTS


def test_apply_to_categories_marks_changes():
    cats = [{'name': 'A', 'items': [
        {'cost_type': '교육운염비', 'description': 'a', 'budget': 1, 'spent': 0, 'warn': []},
        {'cost_type': '시설장비유지비', 'description': 'b', 'budget': 1, 'spent': 0, 'warn': []},
        {'cost_type': '운영수당', 'description': 'c', 'budget': 1, 'spent': 0, 'warn': []}]}]
    items = bt.apply_to_categories(cats, NAMES)[0]['items']
    assert items[0]['cost_type'] == '교육운영비' and '고쳤어요' in items[0]['warn'][0]
    assert '목록에 없는' in items[1]['warn'][0]
    assert items[2]['warn'] == []


def test_saved_list_and_in_use(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    assert bt.load_saved() == list(bt.DEFAULTS)          # 처음엔 기본 목록
    bt.save(['여비', ' 여비 ', '업무추진비'])
    assert bt.load_saved() == ['여비', '업무추진비']        # 중복 정리
    cid = db.add_budget_category('A')
    db.add_budget_item(cid, '시설장비유지비', 'x', 10, 0)
    assert bt.known() == ['여비', '업무추진비', '시설장비유지비']   # 지금 쓰는 비목도 함께
