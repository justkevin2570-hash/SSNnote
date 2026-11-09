"""budget_ai.py — 예산 창에서 AI에게 묻고 답 듣기 (UI 없음 + 뒤에서 실행할 일꾼).

  · 지금 예산 표(세부항목·원가통계비목·산출내역·금액)를 함께 보내 질문에 답한다.
  · "간식비 5만원 썼어"처럼 쓴 돈을 말하면, 어느 항목에 얼마를 넣을지 '제안'만 돌려준다.
    실제 반영은 사용자가 [넣기]를 눌러야 한다 (AI가 예산 데이터를 직접 고치지 않음).
  · 개인정보 보호 모드면 보내는 글의 개인정보를 가린다 (school_profile.mask_for_ai).

응답 형식 (JSON): {"answer": "답", "spend": [{"item_id": 12, "amount": 50000}]}
"""

import json
import re
from datetime import date

SYSTEM = (
    "당신은 한국 학교 예산 담당 선생님을 돕는 도우미입니다. 아래에 주어진 예산 표만 근거로 답합니다.\n"
    "- 금액은 '1,234,000원'처럼 쉼표와 '원'을 붙이고, 계산은 정확히 합니다. 표에 없는 내용은 지어내지 않습니다.\n"
    "- 답은 짧고 분명하게, 목록이 필요하면 '- '로 시작하는 줄로. 마크다운(**, ##, 표)은 쓰지 않습니다.\n"
    "- 학교회계연도는 3월 1일부터 다음 해 2월 말일까지입니다.\n"
    "- 사용자가 돈을 썼다고 말하면(예: '간식비 5만원 썼어', '강사수당 32만원 지급') spend에 항목 id와 금액을 넣고,\n"
    "  answer에는 어느 항목에 얼마를 넣을지 한 줄로 확인하듯 적습니다. 실제 반영은 사용자가 확인한 뒤에 합니다.\n"
    "- 어느 항목인지 애매하면 spend는 비우고 answer에서 후보를 보여 주며 되묻습니다.\n"
    "- 완료된 항목에는 지출을 넣지 않습니다.\n"
    "반드시 JSON만 출력합니다."
)

_FORMAT = '{"answer": "답", "spend": [{"item_id": 12, "amount": 50000}]}'

# 자주 묻는 질문 (버튼 글자, 실제 질문)
QUICK = (
    ('초과 집행 항목', '잔액이 마이너스인(초과 집행된) 항목을 세부항목과 함께 알려줘. 없으면 없다고 해줘.'),
    ('잔액 많은 순', '잔액이 많은 항목 5개를 세부항목·원가통계비목과 함께 알려줘.'),
    ('비목별 요약', '원가통계비목별로 예산·집행액·잔액·집행률을 요약해줘.'),
    ('남은 돈 점검', '회계연도가 끝나기 전에 써야 할 잔액이 큰 세부항목과 항목을 알려주고, 집행률이 낮은 곳을 짚어줘.'),
)


def build_context(categories, items) -> str:
    """예산 표 → AI에 함께 보낼 글. 항목마다 id를 붙여 지출 제안 때 가리킬 수 있게."""
    lines = []
    tb = ts = 0
    for cat in categories:
        cat_items = [it for it in items if it['category_id'] == cat['id']]
        b = sum(it['budget'] for it in cat_items)
        s = sum(it['spent'] for it in cat_items)
        tb, ts = tb + b, ts + s
        lines.append(f"[세부항목] {cat['name']} — 예산 {b:,} · 집행 {s:,} · 잔액 {b - s:,}")
        for it in cat_items:
            done = ' · 완료' if it.get('completed') else ''
            lines.append(f"  id={it['id']} | {it['cost_type']} | {it['description']} | 예산 {it['budget']:,} | "
                         f"집행 {it['spent']:,} | 잔액 {it['budget'] - it['spent']:,}{done}")
    lines.append(f"[전체] 예산 {tb:,} · 집행 {ts:,} · 잔액 {tb - ts:,}")
    return '\n'.join(lines)


def build_prompt(question: str, context: str, history=(), today: date = None) -> str:
    today = today or date.today()
    past = ''.join(f"\n이전 질문: {q}\n이전 답: {a}" for q, a in history[-2:])
    return (f"오늘: {today.isoformat()}\n\n예산 표 (금액 단위: 원)\n{context}\n"
            f"{past}\n\n질문: {question}\n\n다음 형식의 JSON만 출력하세요:\n{_FORMAT}")


def _extract_json(raw: str):
    raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', (raw or '').strip())
    try:
        return json.loads(raw)
    except Exception:
        a, b = raw.find('{'), raw.rfind('}')
        if a != -1 and b > a:
            try:
                return json.loads(raw[a:b + 1])
            except Exception:
                pass
    return None


def _to_int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    digits = re.sub(r'[^\d]', '', str(v or ''))
    return int(digits) if digits else None


def parse_reply(raw: str, items_by_id: dict) -> dict:
    """AI 응답 → {'answer': 글, 'spend': [{'item_id', 'amount', 'item'}]}.

    JSON이 아니면 글 전체를 답으로 쓴다. 지출 제안은 있는 항목·완료 안 된 항목·0보다 큰 금액만 남긴다.
    """
    data = _extract_json(raw)
    if not isinstance(data, dict):
        return {'answer': (raw or '').strip(), 'spend': []}
    answer = str(data.get('answer') or '').strip()
    spend, seen = [], set()
    for s in data.get('spend') or []:
        if not isinstance(s, dict):
            continue
        iid, amount = _to_int(s.get('item_id')), _to_int(s.get('amount'))
        item = items_by_id.get(iid)
        if not item or item.get('completed') or not amount or amount <= 0 or iid in seen:
            continue
        seen.add(iid)
        spend.append({'item_id': iid, 'amount': amount, 'item': item})
    return {'answer': answer, 'spend': spend}


# ── 실행 (뒤에서) ────────────────────────────────────────────────

try:
    from PyQt5.QtCore import QThread, pyqtSignal
except ImportError:          # 테스트 환경 등
    QThread = object

    def pyqtSignal(*_a):     # noqa: N802
        return None


class BudgetAskWorker(QThread):
    """질문 → AI 답.  done(reply: dict, masked: dict) / failed(message: str)"""
    done = pyqtSignal(dict, dict)
    failed = pyqtSignal(str)

    def __init__(self, question, categories, items, history=(), parent=None):
        super().__init__(parent)
        self._q = question
        self._cats = [dict(c) for c in categories]
        self._items = [dict(i) for i in items]
        self._history = list(history)

    def run(self):
        try:
            from ai_client import get_adapter, friendly_error  # noqa: F401
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있어요. 공문 작성 창 → [AI 설정]에서 OpenCode·Claude·ChatGPT·EXAONE 중 하나를 연결하세요.')
                return
            if not adapter.get_api_key():
                self.failed.emit('AI API 키가 없어요. 공문 작성 창 → [AI 설정]에서 키를 넣어 주세요.')
                return
            prompt = build_prompt(self._q, build_context(self._cats, self._items), self._history)
            masked = {}
            try:
                from ai_client import masking_required
                from school_profile import mask_for_ai
                if masking_required():
                    prompt, masked = mask_for_ai(prompt)
            except Exception:
                pass
            raw = adapter.complete(SYSTEM, prompt, max_tokens=2000)
            reply = parse_reply(raw, {it['id']: it for it in self._items})
            if not reply['answer'] and not reply['spend']:
                self.failed.emit('AI가 빈 답을 보냈어요. 다시 물어봐 주세요.')
                return
            self.done.emit(reply, masked or {})
        except Exception as e:
            try:
                from ai_client import friendly_error
                msg = friendly_error(e) if getattr(e, 'status_code', None) else str(e)
            except Exception:
                msg = str(e)
            self.failed.emit(msg)
