"""캡처한 예산 표 → 세부항목별 정리 (budget_image)."""
import json

import budget_image as bi

CAT1 = '[목적]고교학점제 선택과목'
CAT2 = '[목적]고교학점제운영'

# 사용자가 보낸 화면 그대로 (AI가 돌려줄 모양)
SCREEN = [
    (CAT1, '교육운영비', '일과운영 S/W 사용료', '440,000', '0'),
    ('', '기간제교원법정부담금', '선택과목 강사 보험료', '316,000', '85,160'),
    ('', '목적사업업무추진비', '협의회비', '460,000', '298,000'),
    ('', '운영수당', '교원대상연수 외부강사 수당', '920,000', '560,000'),
    ('', '운영수당', '선택과목 강사 수당', '18,000,000', '9,800,000'),
    ('', '일반수용비', '전산소모품 구입', '920,000', '602,000'),
    (CAT2, '교육운영비', '교과융합캠프 간식비', '400,000', '284,500'),
    ('', '교육운영비', '교육과정박람회 운영물품', '1,350,000', '-14,000'),
    ('', '교육운영비', '자료제작비', '2,430,000', '930,000'),
    ('', '교육운영비', '자율적 교육과정 운영비', '3,000,000', '46,240'),
    ('', '교육운영비', '추가교육과정 운영', '1,650,000', '1,650,000'),
    ('', '교육운영비', '특색진로프로그램 간식비', '800,000', '218,910'),
    ('', '교육운영비', '특색진로프로그램 체험비', '7,200,000', '-386,520'),
    ('', '목적사업업무추진비', '협의회비', '1,150,000', '1,150,000'),
    ('', '운영수당', '교과융합캠프 강사수당', '2,520,000', '1,940,000'),
    ('', '운영수당', '교육과정박람회 강사수당', '2,500,000', '2,150,000'),
]


def ai_json(rows=SCREEN):
    return json.dumps({'rows': [dict(category=c, cost_type=t, description=d, budget=b, remaining=r)
                                for c, t, d, b, r in rows]}, ensure_ascii=False)


def test_screen_splits_into_two_categories():
    cats = bi.group_rows(bi.parse_rows(ai_json()))
    assert [c['name'] for c in cats] == [CAT1, CAT2]
    assert [len(c['items']) for c in cats] == [6, 10]
    first = cats[0]['items'][0]
    assert first == {'cost_type': '교육운영비', 'description': '일과운영 S/W 사용료',
                     'budget': 440000, 'spent': 440000, 'warn': []}
    over = cats[1]['items'][1]                      # 잔액 -14,000 → 초과 집행
    assert over['budget'] == 1350000 and over['spent'] == 1364000 and over['warn'] == []
    assert sum(i['budget'] for i in cats[1]['items']) == 23000000


def test_parse_handles_code_fence_and_strings():
    raw = '```json\n' + ai_json(SCREEN[:2]) + '\n```'
    rows = bi.parse_rows(raw)
    assert rows[1]['budget'] == 316000 and rows[1]['remaining'] == 85160


def test_totals_and_blank_rows_are_skipped():
    rows = bi.normalize_rows([
        {'category': '', 'cost_type': '합계', 'description': '', 'budget': 5, 'remaining': 1},
        {'category': '', 'cost_type': '', 'description': '소계', 'budget': 5, 'remaining': 1},
        {'category': '', 'cost_type': '', 'description': '', 'budget': None, 'remaining': None},
        {'category': '', 'cost_type': '운영수당', 'description': '강사 수당', 'budget': 10, 'remaining': 4},
    ])
    assert [r['description'] for r in rows] == ['강사 수당']


def test_warnings():
    cats = bi.group_rows(bi.normalize_rows([
        {'category': 'A', 'cost_type': 'x', 'description': '잔액이 큼', 'budget': 100, 'remaining': 300},
        {'category': '', 'cost_type': 'x', 'description': '잔액 없음', 'budget': 100, 'remaining': None},
    ]))
    a, b = cats[0]['items']
    assert a['warn'] and '잔액이 예산보다' in a['warn'][0]
    assert b['spent'] == 0 and '읽지 못해' in b['warn'][0]


def test_rows_before_first_category_go_to_unclassified():
    cats = bi.group_rows([{'category': '', 'cost_type': 'x', 'description': 'd', 'budget': 1, 'remaining': 0}])
    assert cats[0]['name'] == '미분류'


def test_to_amount():
    assert bi.to_amount('1,350,000') == 1350000
    assert bi.to_amount('-14,000') == -14000
    assert bi.to_amount('△386,520') == -386520
    assert bi.to_amount('(5,000)') == -5000
    assert bi.to_amount('1.350.000') == 1350000      # 쉼표를 마침표로 잘못 읽어도
    assert bi.to_amount('운영수당') is None
    assert bi.to_amount('') is None


# ── Windows OCR 단어 위치로 나누기 (AI 없음) ──

def _words(table, x_cols, y0=0, row_h=30, char_w=12):
    """[줄][칸 글자] → OCR 단어 [(글자, x, y, w, h)]. 칸 안의 띄어쓰기는 단어로 나눈다."""
    words = []
    for r, cells in enumerate(table):
        y = y0 + r * row_h
        for x, text in zip(x_cols, cells):
            cx = x
            for part in str(text).split():
                w = len(part) * char_w
                words.append((part, cx, y, w, 16))
                cx += w + 6
    return words


X = [10, 200, 380, 450, 760, 880]


def test_rows_from_words_with_header():
    header = ['세 부 항 목', '원가통계비목', '선택', '산출내역', '예 산 현 액', '예산잔액']
    body = [[c, t, '□', d, b, r] for c, t, d, b, r in SCREEN[:8]]
    rows = bi.rows_from_words(_words([header] + body, X))
    cats = bi.group_rows(rows)
    assert [c['name'] for c in cats] == [CAT1, CAT2]
    assert [len(c['items']) for c in cats] == [6, 2]
    assert cats[0]['items'][4]['description'] == '선택과목 강사 수당'
    assert cats[0]['items'][4]['spent'] == 18000000 - 9800000
    assert cats[1]['items'][1]['spent'] == 1364000


def test_rows_from_words_without_header():
    body = [[c, t, d, b, r] for c, t, d, b, r in SCREEN[:3]]
    rows = bi.rows_from_words(_words(body, [10, 200, 450, 760, 880]))
    cats = bi.group_rows(rows)
    assert cats[0]['name'] == CAT1 and len(cats[0]['items']) == 3
    assert cats[0]['items'][2]['description'] == '협의회비'


def test_words_to_text_keeps_columns():
    text = bi.words_to_text(_words([[CAT1, '교육운영비', '일과운영 S/W 사용료', '440,000', '0']],
                                   [10, 200, 450, 760, 880]))
    assert text == f'{CAT1} | 교육운영비 | 일과운영 S/W 사용료 | 440,000 | 0'


def test_prompt_mentions_rules():
    p = bi.build_prompt()
    assert '예산잔액' in p and '"rows"' in p
    assert '칸 사이' in bi.build_prompt('a | b')


# ── 읽는 방법 고르기 (가짜 AI·가짜 OCR) ──

def _run_worker(monkeypatch, adapter, privacy, words):
    import pytest
    pytest.importorskip('PyQt5')
    import ai_client
    import capture
    import school_profile
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: adapter)
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: privacy)
    monkeypatch.setattr(capture, 'ocr_words', lambda png: words)
    w = bi.BudgetImageWorker('b64', b'png')
    out = {}
    w.done.connect(lambda cats, method: out.update(cats=cats, method=method))
    w.failed.connect(lambda msg: out.update(error=msg))
    w.run()                                   # 스레드 없이 바로 실행
    return out


class _FakeAI:
    def __init__(self):
        self.calls = []

    @staticmethod
    def get_api_key():
        return 'k'

    def complete(self, system, prompt, image_b64=None, max_tokens=0):
        self.calls.append('image' if image_b64 else 'text')
        return ai_json()


def test_worker_ai_image(monkeypatch):
    ai = _FakeAI()
    out = _run_worker(monkeypatch, ai, privacy=False, words=[])
    assert out['method'] == 'ai-image' and ai.calls == ['image']
    assert [len(c['items']) for c in out['cats']] == [6, 10]


def test_worker_privacy_sends_text_only(monkeypatch):
    ai = _FakeAI()
    words = _words([[c, t, d, b, r] for c, t, d, b, r in SCREEN[:2]], [10, 200, 450, 760, 880])
    out = _run_worker(monkeypatch, ai, privacy=True, words=words)
    assert out['method'] == 'ai-text' and ai.calls == ['text']


def test_worker_without_ai_uses_ocr_layout(monkeypatch):
    header = ['세부항목', '원가통계비목', '선택', '산출내역', '예산현액', '예산잔액']
    words = _words([header] + [[c, t, '□', d, b, r] for c, t, d, b, r in SCREEN], X)
    out = _run_worker(monkeypatch, None, privacy=False, words=words)
    assert out['method'] == 'ocr'
    assert [c['name'] for c in out['cats']] == [CAT1, CAT2]
    assert [len(c['items']) for c in out['cats']] == [6, 10]


def test_worker_reports_when_nothing_read(monkeypatch):
    out = _run_worker(monkeypatch, None, privacy=False, words=[])
    assert '글자를 읽지 못했' in out['error']
