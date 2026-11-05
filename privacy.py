"""privacy.py — AI로 보내기 전 개인정보 가리기.

외부 AI(Gemini·Claude·ChatGPT 등)로 글을 보내기 전에 주민등록번호, 전화번호,
이메일, 사용자가 등록한 이름(학생 명단 등)을 가린다. UI·DB에 의존하지 않는 순수 로직.

    masked, counts = mask_pii(text, names=['홍길동'], exempt=['031-123-4567'])
"""

import re

# 주민등록번호 / 외국인등록번호: 6자리-7자리 (뒷자리 첫 숫자 1~8)
_RRN = re.compile(r'(?<!\d)(\d{6})\s?-\s?([1-8])\d{6}(?!\d)')
# 휴대전화: 010-1234-5678, 010 1234 5678, 01012345678
_MOBILE = re.compile(r'(?<!\d)(01[016789])[-.\s]?(\d{3,4})[-.\s]?(\d{4})(?!\d)')
# 일반 전화: 02-123-4567, 031-1234-5678 (구분자 필수 — 문서번호·금액과 혼동 방지)
_PHONE = re.compile(r'(?<!\d)(0\d{1,2})[-.)]\s?(\d{3,4})[-.]\s?(\d{4})(?!\d)')
_EMAIL = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')

_LABELS = {'rrn': '주민등록번호', 'phone': '전화번호', 'email': '이메일', 'name': '이름'}


def _mask_name(name: str) -> str:
    """'홍길동' → '홍○○', '남궁민수' → '남○○○'."""
    return name[0] + '○' * (len(name) - 1) if len(name) > 1 else '○'


def mask_pii(text: str, names=(), exempt=()):
    """개인정보를 가린 문자열과 종류별 가린 개수를 돌려준다.

    names:  가릴 이름 목록 (2글자 이상만 사용)
    exempt: 가리지 않을 문자열 (학교 대표번호·담당자 연락처처럼 공문에 들어가야 하는 값)
    """
    if not text:
        return text, {}
    counts = {}
    keep = {e.strip() for e in exempt if e and e.strip()}

    def sub(kind, pattern, repl):
        nonlocal text

        def _r(m):
            if m.group(0).strip() in keep:
                return m.group(0)
            counts[kind] = counts.get(kind, 0) + 1
            return repl(m)
        text = pattern.sub(_r, text)

    # 주민번호를 전화번호보다 먼저 (숫자 패턴 겹침 방지)
    sub('rrn', _RRN, lambda m: f'{m.group(1)}-{m.group(2)}******')
    sub('email', _EMAIL, lambda m: '***@' + m.group(0).split('@', 1)[1])
    sub('phone', _MOBILE, lambda m: f'{m.group(1)}-****-{m.group(3)}')
    sub('phone', _PHONE, lambda m: f'{m.group(1)}-****-{m.group(3)}')

    # 이름: 긴 이름부터 (예: '김민수'가 '김민수진'을 망가뜨리지 않도록)
    uniq = sorted({n.strip() for n in names if n and len(n.strip()) >= 2}, key=len, reverse=True)
    for n in uniq:
        if n in keep:
            continue
        hits = text.count(n)
        if hits:
            counts['name'] = counts.get('name', 0) + hits
            text = text.replace(n, _mask_name(n))
    return text, counts


def describe(counts: dict) -> str:
    """{'phone': 2, 'name': 1} → '전화번호 2건, 이름 1건'."""
    return ', '.join(f'{_LABELS[k]} {v}건' for k, v in counts.items() if v)


def parse_name_list(text: str) -> list:
    """쉼표·줄바꿈·공백으로 구분된 이름 목록을 리스트로."""
    return [n for n in re.split(r'[,\s]+', text or '') if len(n) >= 2]
