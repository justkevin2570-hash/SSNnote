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
    assert r['answer'].startswith('간식비') and [(s['item_id'], s['amount']) for s in r['spend']] == [(11, 50000)]


def test_parse_reply_plain_text_and_string_amount():
    assert ba.parse_reply('그냥 글로 답함', BY_ID) == {'answer': '그냥 글로 답함', 'spend': []}
    r = ba.parse_reply('{"answer": "a", "spend": [{"item_id": "11", "amount": "50,000원"}]}', BY_ID)
    assert r['spend'][0]['amount'] == 50000


def test_quick_questions():
    assert len(ba.QUICK) == 4 and all(label and q for label, q in ba.QUICK)
