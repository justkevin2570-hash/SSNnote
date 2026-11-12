"""budget_ai.py — 예산 창에서 AI에게 묻고 답 듣기 (UI 없음 + 뒤에서 실행할 일꾼).

  · 지금 예산 표(세부항목·원가통계비목·산출내역·금액)를 함께 보내 질문에 답한다.
  · "간식비 5만원 썼어"처럼 쓴 돈, "협의회비 잔액 30만원으로 맞춰줘"처럼 목표 잔액,
    "간식비 집행액을 12만원으로 고쳐줘"처럼 집행액을 말하면, AI는 항목과 목표값만 정한다.
    잔액·집행액 계산과 검증은 앱이 한다 (parse_reply). 실제 반영은 사용자가 [넣기]를 눌러야 한다.
  · 개인정보 보호 모드면 보내는 글의 개인정보를 가린다 (school_profile.mask_for_ai).

응답 형식 (JSON):
  {"answer": "답", "spend": [{"item_id": 12, "amount": 50000}],
   "set_remaining": [{"item_id": 12, "remaining": 300000}], "set_spent": [{"item_id": 7, "spent": 120000}]}
"""

import json
import re
from datetime import date

SYSTEM = (
    "당신은 한국 학교 예산 담당 선생님을 돕는 도우미입니다. 아래에 주어진 예산 표만 근거로 답합니다.\n"
    "- 금액은 '1,234,000원'처럼 쉼표와 '원'을 붙이고, 계산은 정확히 합니다. 표에 없는 내용은 지어내지 않습니다.\n"
    "- 답은 짧고 분명하게, 목록이 필요하면 '- '로 시작하는 줄로. 마크다운(**, ##, 표)은 쓰지 않습니다.\n"
    "- 학교회계연도는 3월 1일부터 다음 해 2월 말일까지입니다.\n"
    "- '전체 합계에서 제외'로 표시된 세부항목은 전체 금액·전체 집행률에 넣지 않습니다. 그 세부항목을 직접 물으면 답합니다.\n"
    "- [그룹]은 묶인 세부항목들의 합입니다. 그룹 이름으로 물으면 그 합으로 답합니다.\n"
    "- 사용자가 돈을 썼다고 말하면(예: '간식비 5만원 썼어', '강사수당 32만원 지급') spend에 항목 id, 쓴 금액(amount),\n"
    "  쓴 곳(purpose: 무엇에 썼는지 짧게, 예: '체육대회', '교육반성회 식사비')을 넣습니다. 쓴 곳을 말하지 않았으면 purpose는 빈 글로 둡니다.\n"
    "- 잔액을 특정 금액으로 맞추라고 하면(예: '협의회비 잔액 30만원으로 맞춰줘') set_remaining에 항목 id와 목표 잔액(remaining)을 넣습니다.\n"
    "- 집행액을 특정 금액으로 고치라고 하면(예: '간식비 집행액을 12만원으로 고쳐줘') set_spent에 항목 id와 집행액(spent)을 넣습니다.\n"
    "- 계산은 하지 마세요. 목표값(금액)만 정하면, 잔액·집행액 계산과 확인은 앱이 합니다.\n"
    "- answer에는 어느 항목에 얼마를 바꿀지 한 줄로 확인하듯 적습니다. 실제 반영은 사용자가 확인한 뒤에 합니다.\n"
    "- 어느 항목인지 애매하면 spend·set_remaining·set_spent를 모두 비우고 answer에서 후보를 보여 주며 되묻습니다.\n"
    "- 완료된 항목에는 지출을 넣지 않습니다.\n"
    "- 항목 id는 spend·set_remaining·set_spent에만 쓰고, answer에는 id를 절대 쓰지 마세요. 항목은 산출내역(이름)으로 말합니다.\n"
    "반드시 JSON만 출력합니다."
)

_FORMAT = ('{"answer": "답", "spend": [{"item_id": 12, "amount": 50000, "purpose": "쓴 곳"}], '
           '"set_remaining": [{"item_id": 12, "remaining": 300000}], '
           '"set_spent": [{"item_id": 7, "spent": 120000}]}')

PURPOSE_MAX = 40                 # 쓴 곳 글자 수 한도

# AI 응답의 세 종류: (응답 키, 값 키, 뜻)
_KINDS = (('spend', 'amount'), ('set_remaining', 'remaining'), ('set_spent', 'spent'))

# 자주 묻는 질문 (버튼 글자, 실제 질문)
QUICK = (
    ('초과 집행 항목', '잔액이 마이너스인(초과 집행된) 항목을 세부항목과 함께 알려줘. 없으면 없다고 해줘.'),
    ('잔액 많은 순', '잔액이 많은 항목 5개를 세부항목·원가통계비목과 함께 알려줘.'),
    ('비목별 요약', '원가통계비목별로 예산·집행액·잔액·집행률을 요약해줘.'),
    ('남은 돈 점검', '회계연도가 끝나기 전에 써야 할 잔액이 큰 세부항목과 항목을 알려주고, 집행률이 낮은 곳을 짚어줘.'),
)


def build_context(categories, items, groups=()) -> str:
    """예산 표 → AI에 함께 보낼 글. 항목마다 id를 붙여 지출 제안 때 가리킬 수 있게.
    전체 합계에서 뺀 세부항목은 '합계 제외'로 표시하고 [전체]에 더하지 않는다. 그룹은 묶인 세부항목의 합."""
    lines = []
    tb = ts = 0
    sums = {}
    for cat in categories:
        cat_items = [it for it in items if it['category_id'] == cat['id']]
        b = sum(it['budget'] for it in cat_items)
        s = sum(it['spent'] for it in cat_items)
        sums[cat['id']] = (b, s)
        excluded = bool(cat.get('excluded'))
        if not excluded:
            tb, ts = tb + b, ts + s
        mark = ' (전체 합계에서 제외)' if excluded else ''
        lines.append(f"[세부항목] {cat['name']}{mark} — 예산 {b:,} · 집행 {s:,} · 잔액 {b - s:,}")
        for it in cat_items:
            done = ' · 완료' if it.get('completed') else ''
            lines.append(f"  id={it['id']} | {it['cost_type']} | {it['description']} | 예산 {it['budget']:,} | "
                         f"집행 {it['spent']:,} | 잔액 {it['budget'] - it['spent']:,}{done}")
    for g in groups:
        members = [c for c in categories if c.get('group_id') == g['id']]
        if not members:
            continue
        b = sum(sums[c['id']][0] for c in members)
        s = sum(sums[c['id']][1] for c in members)
        names = ', '.join(c['name'] for c in members)
        lines.append(f"[그룹] {g['name']} ({names}) — 예산 {b:,} · 집행 {s:,} · 잔액 {b - s:,}")
    n_ex = sum(1 for c in categories if c.get('excluded'))
    note = f' (합계 제외 세부항목 {n_ex}개 뺌)' if n_ex else ''
    lines.append(f"[전체]{note} 예산 {tb:,} · 집행 {ts:,} · 잔액 {tb - ts:,}")
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


def target_spent(kind: str, value: int, item: dict) -> int:
    """바꿀 집행액. 계산은 여기서만 한다 (AI는 목표값만 정한다)."""
    if kind == 'spend':
        return int(item['spent']) + value                  # 쓴 금액을 더함
    if kind == 'set_remaining':
        return int(item['budget']) - value                 # 목표 잔액 → 집행액
    return value                                           # set_spent: 집행액 그대로


def _check(kind: str, value, item) -> tuple:
    """(바꿀 값 또는 None, 이유). 이유가 None이면 바꾸지 않아도 되는 경우(0원·변화 없음)."""
    if value is None:
        return None, '금액을 읽지 못했어요'
    if value < 0:
        return None, '금액이 0원보다 작아요'
    if kind == 'spend' and value == 0:
        return None, None
    if kind == 'set_remaining' and value > int(item['budget']):
        return None, '잔액 목표가 예산보다 커요'
    new = target_spent(kind, value, item)
    if new < 0:
        return None, '집행액이 0원보다 작아져요'
    if new == int(item['spent']):
        return None, None
    if kind == 'set_remaining' and int(item['budget']) - new != value:
        return None, '계산이 맞지 않아요'
    return new, None


def parse_reply(raw: str, items_by_id: dict) -> dict:
    """AI 응답 → {'answer': 글, 'changes': [...], 'rejected': [...]}.

    changes: 앱이 검증을 통과한 변경 (항목·종류·기존/변경 집행액·잔액). 중복·완료·없는 항목·음수 등은 rejected에 이유와 함께 남긴다.
    JSON이 아니면 글 전체를 답으로 쓴다.
    """
    data = _extract_json(raw)
    if not isinstance(data, dict):
        return {'answer': (raw or '').strip(), 'changes': [], 'rejected': []}
    answer = str(data.get('answer') or '').strip()
    changes, rejected, done_ids = [], [], set()
    for key, value_key in _KINDS:
        for s in data.get(key) or []:
            if not isinstance(s, dict):
                continue
            iid = _to_int(s.get('item_id'))
            value = _to_int(s.get(value_key))
            item = items_by_id.get(iid)
            if item is None:
                rejected.append({'item_id': iid, 'item': None, 'kind': key, 'reason': '없는 항목이에요'})
                continue
            if item.get('completed'):
                rejected.append({'item_id': iid, 'item': item, 'kind': key, 'reason': '완료된 항목이에요'})
                continue
            if iid in done_ids:
                rejected.append({'item_id': iid, 'item': item, 'kind': key,
                                 'reason': '같은 항목이 여러 번 나와 첫 번째만 써요'})
                continue
            new, reason = _check(key, value, item)
            if reason:
                rejected.append({'item_id': iid, 'item': item, 'kind': key, 'reason': reason})
                continue
            if new is None:
                continue                                    # 0원·변화 없음: 조용히 뺌
            done_ids.add(iid)
            budget, spent = int(item['budget']), int(item['spent'])
            purpose = str(s.get('purpose') or '').strip()[:PURPOSE_MAX] if key == 'spend' else ''
            changes.append({'item_id': iid, 'item': item, 'kind': key, 'value': value, 'purpose': purpose,
                            'old_spent': spent, 'new_spent': new,
                            'old_remaining': budget - spent, 'new_remaining': budget - new})
    return {'answer': answer, 'changes': changes, 'rejected': rejected}


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

    def __init__(self, question, categories, items, history=(), parent=None, groups=()):
        super().__init__(parent)
        self._q = question
        self._cats = [dict(c) for c in categories]
        self._groups = [dict(g) for g in groups]
        self._items = [dict(i) for i in items]
        self._history = list(history)

    def run(self):
        try:
            from ai_client import get_adapter, friendly_error  # noqa: F401
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있어요. 공문 작성 창 → [AI 설정]에서 OpenCode·Claude·ChatGPT·Claude 구독 중 하나를 연결하세요.')
                return
            if not adapter.get_api_key():
                self.failed.emit('AI API 키가 없어요. 공문 작성 창 → [AI 설정]에서 키를 넣어 주세요.')
                return
            prompt = build_prompt(self._q, build_context(self._cats, self._items, self._groups), self._history)
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
            if not reply['answer'] and not reply['changes'] and not reply['rejected']:
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
