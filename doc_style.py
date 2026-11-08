"""doc_style.py — 내 공문 스타일을 AI에 알려 주기 (UI 없음).

  1) 비슷한 공문 찾기: 공문 작성 창에서 저장한 공문 중 제목이 비슷한 것을 골라 AI에 예시로 보낸다.
     큰 모델 없이 제목의 글자 두 개 묶음(교무·무회·회의…)이 얼마나 겹치는지로 비교한다.
     흔한 말(계획·운영·2026 등)은 덜, 드문 말(경시대회·박람회)은 더 무겁게 친다.
     같은 종류(_detect_doc_type)를 먼저, 최근 공문을 조금 더 먼저 고른다.
     비슷한 게 모자라면 같은 종류의 최근 공문으로 채운다 — 내용이 달라도 문체는 배울 수 있으므로.
  2) 나만의 작성 규칙: 사용자가 적어 둔 메모(app_settings 'doc_style_notes')를 매번 함께 보낸다.

예시와 규칙은 AI에 보내는 글(prompt)에 넣으므로 개인정보 보호 모드의 가리기를 그대로 거친다.
"""

import math
import re

STYLE_NOTES_KEY = 'doc_style_notes'
STYLE_NOTES_MAX = 1000          # 규칙 메모 최대 글자 수
EXAMPLE_MAX_CHARS = 1500        # 예시 공문 하나당 보낼 최대 글자 수
MIN_SCORE = 0.2                 # 이보다 덜 비슷하면 '비슷한 공문'으로 치지 않음


# ── 나만의 작성 규칙 ─────────────────────────────────────────────

def load_style_notes() -> str:
    try:
        from db import get_setting
        return (get_setting(STYLE_NOTES_KEY, '') or '').strip()
    except Exception:
        return ''


def save_style_notes(text: str):
    from db import set_setting
    set_setting(STYLE_NOTES_KEY, (text or '').strip()[:STYLE_NOTES_MAX])


def style_notes_prompt(notes: str = None) -> str:
    notes = load_style_notes() if notes is None else (notes or '').strip()
    if not notes:
        return ''
    return ('\n[작성자의 작성 규칙 — 서식 규칙보다 우선해서 꼭 따를 것]\n'
            f'{notes}\n')


# ── 비슷한 공문 찾기 ─────────────────────────────────────────────

def _bigrams(text: str) -> set:
    """한글·영문만 남겨 단어별 글자 두 개 묶음. 한 글자 단어는 그대로."""
    out = set()
    for word in re.findall(r'[가-힣A-Za-z]+', text or ''):
        word = word.lower()
        if len(word) == 1:
            out.add(word)
        out.update(word[i:i + 2] for i in range(len(word) - 1))
    return out


def search_similar(title: str, docs: list, doc_type: str = None, top_k: int = 2) -> list:
    """제목과 비슷한 저장 공문 top_k개.

    docs: db.get_official_documents() 결과 (최근 순). 각 항목에 'doc_type'이 채워져 있어야 같은 종류를 우선한다.
    반환: [{'title', 'content', 'doc_type', 'score'}] — score가 0이면 '같은 종류 최근 공문'으로 채운 것.
    """
    # 같은 제목으로 여러 번 저장했으면 가장 최근 것만
    seen, pool = set(), []
    for d in docs or []:
        t = (d.get('title') or '').strip()
        if not t or not (d.get('content') or '').strip() or t in seen:
            continue
        seen.add(t)
        pool.append(d)
    if not pool:
        return []

    grams = [_bigrams(d['title']) for d in pool]
    n = len(pool)
    df = {}
    for g in grams:
        for b in g:
            df[b] = df.get(b, 0) + 1

    def idf(b):
        return math.log((n + 1) / (df.get(b, 0) + 1)) + 1.0

    q = _bigrams(title)
    q_norm = math.sqrt(sum(idf(b) ** 2 for b in q)) if q else 0.0

    scored = []
    for i, (d, g) in enumerate(zip(pool, grams)):
        if not q_norm or not g:
            continue
        common = q & g
        if not common:
            continue
        d_norm = math.sqrt(sum(idf(b) ** 2 for b in g))
        sim = sum(idf(b) ** 2 for b in common) / (q_norm * d_norm)
        if sim < MIN_SCORE:
            continue
        bonus = (0.1 if doc_type and d.get('doc_type') == doc_type else 0.0) + 0.05 * (1 - i / n)
        scored.append((sim + bonus, i))
    scored.sort(key=lambda x: -x[0])

    picked = [(s, i) for s, i in scored[:top_k]]
    used = {i for _, i in picked}
    if doc_type and len(picked) < top_k:                     # 모자라면 같은 종류 최근 공문
        for i, d in enumerate(pool):
            if len(picked) >= top_k:
                break
            if i not in used and d.get('doc_type') == doc_type:
                picked.append((0.0, i))
                used.add(i)

    return [{'title': pool[i]['title'], 'content': pool[i]['content'],
             'doc_type': pool[i].get('doc_type', ''), 'score': round(s, 3)} for s, i in picked]


def examples_prompt(examples: list) -> str:
    """찾은 공문 → AI에 함께 보낼 예시 글."""
    if not examples:
        return ''
    lines = ['\n[작성자가 예전에 쓴 공문 — 문체·표현·항목 구성·띄어쓰기를 그대로 따르세요. '
             '날짜·장소·수치 같은 내용은 베끼지 말고 이번 제목에 맞게 쓰세요.]']
    for i, d in enumerate(examples, 1):
        body = d['content'].strip()
        if len(body) > EXAMPLE_MAX_CHARS:
            body = body[:EXAMPLE_MAX_CHARS] + '\n…(이하 생략)'
        lines.append(f'\n[예시 {i}] 제목: {d["title"]}\n{body}')
    return '\n'.join(lines) + '\n'
