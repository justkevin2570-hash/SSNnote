"""예산 창 AI에게 묻기 (budget_ai) — 표 만들기 · 응답 읽기 · 지출 제안 걸러내기."""
import json
from datetime import date

import budget_ai as ba

CATS = [{'id': 1, 'name': '[목적]고교학점제운영'}]
ITEMS = [
    {'id': 11, 'category_id': 1, 'cost_type': '교육운영비', 'description': '교과융합캠프 간식비',
     'budget': 400000, 'spent': 115500, 'completed': 0},
    {'id': 12, 'category_id': 1, 'cost_type': '운영수당', 'description': '강사수당',
     'budget': 2520000, 'spent': 2520000, 'completed': 1},
]
BY_ID = {it['id']: it for it in ITEMS}


def test_context_has_ids_totals_and_done():
    ctx = ba.build_context(CATS, ITEMS)
    assert 'id=11 | 교육운영비 | 교과융합캠프 간식비 | 예산 400,000 | 집행 115,500 | 잔액 284,500' in ctx
    assert 'id=12' in ctx and '완료' in ctx
    assert '[전체] 예산 2,920,000 · 집행 2,635,500 · 잔액 284,500' in ctx


def test_prompt_has_today_history_and_format():
    p = ba.build_prompt('얼마 남았어?', 'CTX', [('q1', 'a1'), ('q2', 'a2'), ('q3', 'a3')], date(2026, 11, 9))
    assert '오늘: 2026-11-09' in p and 'CTX' in p and '"spend"' in p
    assert 'q1' not in p and 'q2' in p and 'q3' in p          # 최근 두 개만


def test_parse_reply_filters_spend():
    raw = json.dumps({'answer': '간식비에 50,000원을 넣을게요.', 'spend': [
        {'item_id': 11, 'amount': 50000},
        {'item_id': 12, 'amount': 1000},        # 완료 항목 → 뺌
        {'item_id': 99, 'amount': 1000},        # 없는 항목 → 뺌
        {'item_id': 11, 'amount': 1},           # 같은 항목 두 번 → 뺌
        {'item_id': '11', 'amount': 0},         # 0원 → 뺌
    ]}, ensure_ascii=False)
    r = ba.parse_reply('```json\n' + raw + '\n```', BY_ID)
    assert r['answer'].startswith('간식비')
    assert [(c['item_id'], c['kind'], c['value'], c['new_spent']) for c in r['changes']] == [(11, 'spend', 50000, 165500)]
    assert len(r['rejected']) == 4                          # 완료 · 없는 항목 · 같은 항목 중복 2건


def test_parse_reply_plain_text_and_string_amount():
    assert ba.parse_reply('그냥 글로 답함', BY_ID) == {'answer': '그냥 글로 답함', 'changes': [], 'rejected': []}
    r = ba.parse_reply('{"answer": "a", "spend": [{"item_id": "11", "amount": "50,000원"}]}', BY_ID)
    assert r['changes'][0]['value'] == 50000


BIG = {'id': 21, 'category_id': 1, 'cost_type': '협의회비', 'description': '협의회비',
       'budget': 1150000, 'spent': 0, 'completed': 0}
BY_ID2 = {**BY_ID, 21: BIG}


def test_set_remaining_turns_target_balance_into_spent():
    raw = json.dumps({'answer': '협의회비 잔액을 30만원으로 맞출게요.',
                      'set_remaining': [{'item_id': 21, 'remaining': 300000}]}, ensure_ascii=False)
    r = ba.parse_reply(raw, BY_ID2)
    c = r['changes'][0]
    assert (c['kind'], c['old_spent'], c['new_spent']) == ('set_remaining', 0, 850000)
    assert (c['old_remaining'], c['new_remaining']) == (1150000, 300000)
    assert r['rejected'] == []


def test_set_remaining_above_budget_is_rejected():
    raw = '{"answer": "x", "set_remaining": [{"item_id": 21, "remaining": 2000000}]}'
    r = ba.parse_reply(raw, BY_ID2)
    assert r['changes'] == [] and '예산보다' in r['rejected'][0]['reason']


def test_set_spent_sets_exactly_and_rejects_negative():
    r = ba.parse_reply('{"answer": "x", "set_spent": [{"item_id": 11, "spent": 120000}]}', BY_ID2)
    assert (r['changes'][0]['new_spent'], r['changes'][0]['new_remaining']) == (120000, 280000)
    r = ba.parse_reply('{"answer": "x", "set_spent": [{"item_id": 11, "spent": -5}]}', BY_ID2)
    assert r['changes'] == [] and '0원보다' in r['rejected'][0]['reason']


def test_no_change_and_zero_are_dropped_quietly():
    r = ba.parse_reply('{"answer": "x", "set_spent": [{"item_id": 11, "spent": 115500}],'
                       ' "spend": [{"item_id": 11, "amount": 0}]}', BY_ID2)
    assert r['changes'] == [] and r['rejected'] == []


def test_missing_and_completed_items_have_reasons():
    raw = ('{"answer": "x", "set_spent": [{"item_id": 99, "spent": 1}, {"item_id": 12, "spent": 1}],'
           ' "set_remaining": [{"item_id": 11, "remaining": 1}, {"item_id": 11, "remaining": 2}]}')
    r = ba.parse_reply(raw, BY_ID2)
    reasons = [x['reason'] for x in r['rejected']]
    assert '없는 항목이에요' in reasons and '완료된 항목이에요' in reasons
    assert [c['item_id'] for c in r['changes']] == [11]                # 같은 항목은 첫 번째만
    assert any('첫 번째만' in x for x in reasons)


def test_target_spent_helper():
    assert ba.target_spent('spend', 50000, {'spent': 115500, 'budget': 400000}) == 165500
    assert ba.target_spent('set_remaining', 300000, {'spent': 0, 'budget': 1150000}) == 850000
    assert ba.target_spent('set_spent', 120000, {'spent': 1, 'budget': 9}) == 120000


def test_system_prompt_teaches_three_kinds_and_no_arithmetic():
    assert 'set_remaining' in ba.SYSTEM and 'set_spent' in ba.SYSTEM and '계산은 하지 마세요' in ba.SYSTEM
    assert 'set_remaining' in ba._FORMAT and 'set_spent' in ba._FORMAT


def test_quick_questions():
    assert len(ba.QUICK) == 4 and all(label and q for label, q in ba.QUICK)


def test_system_prompt_keeps_ids_out_of_answers():
    assert 'answer에는 id를 절대 쓰지' in ba.SYSTEM            # 실제 EXAONE 시험에서 답에 'id=12'가 그대로 나왔음
