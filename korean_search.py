"""korean_search.py — 한글 초성·부분 일치 검색 점수 (명령 팔레트용, UI 의존 없음).

    score('ㄱㅁㅎㅇ', '교무회의 자료 제출')  → 양수 (초성 일치)
    score('회의', '교무회의 자료 제출')      → 양수 (부분 일치)
    score('xyz', '교무회의')                → 0 (불일치)
"""

_CHO = 'ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ'
_CHO_SET = set(_CHO)


def chosung(ch: str) -> str:
    """'교' → 'ㄱ'. 한글 음절이 아니면 그대로."""
    code = ord(ch) - 0xAC00
    if 0 <= code < 11172:
        return _CHO[code // 588]
    return ch


def _char_match(q: str, t: str) -> bool:
    return q == t or (q in _CHO_SET and chosung(t) == q)


def score(query: str, text: str) -> int:
    """일치 정도 (0 = 불일치). 앞부분 일치 > 부분 일치 > 초성 일치 > 글자 순서만 일치."""
    q = ''.join(query.lower().split())
    if not q:
        return 1
    t_full = text.lower()
    t = ''.join(t_full.split())  # 띄어쓰기 무시
    if not t:
        return 0
    if t.startswith(q):
        return 1000 - len(t)
    pos = t.find(q)
    if pos >= 0:
        return 800 - pos * 5 - len(t)
    # 초성 섞인 연속 일치 ('ㄱㅁ회의', 'ㅇㅅ')
    n = len(q)
    for i in range(len(t) - n + 1):
        if all(_char_match(q[j], t[i + j]) for j in range(n)):
            return 600 - i * 5 - len(t)
    # 글자 순서만 일치 (띄엄띄엄)
    i = 0
    gaps = 0
    for ch in t:
        if i < n and _char_match(q[i], ch):
            i += 1
        elif i:
            gaps += 1
    if i == n:
        return max(1, 300 - gaps * 10 - len(t))
    return 0


def best_score(query: str, *texts) -> int:
    return max((score(query, x) for x in texts if x), default=0)
