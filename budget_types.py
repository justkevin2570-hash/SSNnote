"""budget_types.py — 원가통계비목 목록과 글자 교정 (UI 없음).

목록 = 저장해 둔 목록(app_settings 'budget_cost_types', 처음엔 DEFAULTS) ∪ 지금 예산에 들어 있는 비목.
코드에 고정된 목록이 아니라, 예산 창 [⋯] → 원가통계비목 목록에서 더하고 뺄 수 있고
새 비목을 한 번 가져오면 다음부터는 그것도 '아는 비목'이 된다.

캡처를 읽은 뒤(correct): 목록과 한두 글자만 다르면 목록의 이름으로 고친다 ('교육운염비' → '교육운영비').
여러 글자가 다르면('기간제교원…' ↔ '기간제근로자…') 다른 비목으로 보고 그대로 둔다.
"""

import json

SETTING_KEY = 'budget_cost_types'

# 처음 쓸 때의 목록 (설정에 저장된 뒤에는 사용자가 고친 목록을 쓴다)
DEFAULTS = ('여비', '운영수당', '일반수용비', '교육운영비', '기간제근로자법정부담금',
            '일반업무추진비', '기간제교원법정부담금', '목적사업업무추진비')


def _clean(name) -> str:
    return ' '.join(str(name or '').split())


def _dedupe(names) -> list:
    out, seen = [], set()
    for n in names:
        n = _clean(n)
        key = n.replace(' ', '')
        if n and key not in seen:
            seen.add(key)
            out.append(n)
    return out


def load_saved() -> list:
    """저장해 둔 목록 (없으면 기본 목록)."""
    try:
        from db import get_setting
        raw = get_setting(SETTING_KEY, '')
        if raw:
            data = json.loads(raw)
            if isinstance(data, list):
                return _dedupe(data)
    except Exception:
        pass
    return list(DEFAULTS)


def save(names):
    from db import set_setting
    set_setting(SETTING_KEY, json.dumps(_dedupe(names), ensure_ascii=False))


def in_use() -> list:
    """지금 예산 항목에 쓰이고 있는 비목."""
    try:
        from db import get_budget_items
        return _dedupe(it.get('cost_type') for it in get_budget_items()
                       if _clean(it.get('cost_type')) not in ('', '일반'))
    except Exception:
        return []


def known() -> list:
    """아는 비목 전체 = 저장한 목록 + 지금 쓰이는 비목."""
    return _dedupe(load_saved() + in_use())


# ── 교정 ───────────────────────────────────────────────────────

def edit_distance(a: str, b: str) -> int:
    """두 글자열 사이 편집 거리 (넣기·빼기·바꾸기 한 번 = 1)."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def correct(text, names) -> tuple:
    """(고친 이름, 고쳤는지, 목록에 있는지).

    띄어쓰기만 다르면 목록의 이름으로 맞춘다. 글자가 다르면 편집 거리가
    max(1, 길이 // 5) 이하이고 가장 가까운 후보가 하나뿐일 때만 고친다.
    한 글자는 고치지 않고, 세 글자 이하는 길이가 같은 이름(한 글자 바뀜)만 본다.
    """
    t = _clean(text)
    key = t.replace(' ', '')
    if not key:
        return t, False, False
    for n in names:
        if n.replace(' ', '') == key:
            return n, n != t, True
    if len(key) < 2:
        return t, False, False                      # 한 글자는 무엇인지 짐작할 수 없음
    limit = max(1, len(key) // 5)
    cands = [n for n in names if len(key) > 3 or len(n.replace(' ', '')) == len(key)]   # 짧은 이름은 글자 바뀜만
    scored = sorted((edit_distance(key, n.replace(' ', '')), n) for n in cands)
    if scored and scored[0][0] <= limit and (len(scored) == 1 or scored[1][0] > scored[0][0]):
        return scored[0][1], True, True
    return t, False, False


def apply_to_categories(cats, names) -> list:
    """가져올 묶음의 원가통계비목을 목록에 맞춰 고친다. 고친 칸·목록에 없는 비목은 warn에 적는다."""
    if not names:
        return cats
    for cat in cats:
        for it in cat['items']:
            original = it.get('cost_type') or ''
            if original in ('', '일반'):
                continue
            fixed, changed, ok = correct(original, names)
            if changed:
                it['cost_type'] = fixed
                if fixed.replace(' ', '') != original.replace(' ', ''):
                    it.setdefault('warn', []).append(f"원가통계비목 '{original}'을(를) '{fixed}'(으)로 고쳤어요")
            elif not ok:
                it.setdefault('warn', []).append(f"목록에 없는 원가통계비목 '{original}' — 맞는지 확인하세요")
    return cats
