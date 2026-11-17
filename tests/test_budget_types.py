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


def test_old_saved_list_is_reordered_once(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    import json
    db.set_setting(bt.SETTING_KEY, json.dumps(['여비', '운영수당', '내가 더한 비목'], ensure_ascii=False))
    names = bt.load_saved()
    assert names[:len(bt.PRIORITY)] == list(bt.PRIORITY) and names[-1] == '내가 더한 비목'
    assert len(names) == len(bt.DEFAULTS) + 1
    bt.save(['여비'])                                       # 그다음부터는 사용자가 고친 목록 그대로
    assert bt.load_saved() == ['여비']


def test_defaults_from_edufine_list():
    assert bt.DEFAULTS[:len(bt.PRIORITY)] == bt.PRIORITY and len(bt.PRIORITY) == 11
    for n in ('학습준비물', '도서구입비', '기간제직원법정부담금', '일반회계전출금', '기간제근로자법정부담금'):
        assert n in bt.DEFAULTS
    assert not any('(' in n for n in bt.DEFAULTS)


def test_short_names_and_icons_for_frequent_costs():
    assert set(bt.PRIORITY) <= set(bt.COST_STYLE) and set(bt.COST_STYLE) - set(bt.PRIORITY) == {'기간제교원법정부담금'}   # 자주 쓰는 11개 + 기간제부담금
    assert bt.short_name('일반업무추진비') == '업추비' and bt.short_name('교육운영비') == '교육운영'
    assert bt.short_name(' 일반 수용비 ') == '수용비'                   # 띄어쓰기가 달라도
    assert bt.short_name('전기요금') == '전기요금' and bt.cost_style('전기요금') is None
    assert bt.cost_style('여비')[1] == 'mdi6.bus-side'                 # 버스 옆모습
