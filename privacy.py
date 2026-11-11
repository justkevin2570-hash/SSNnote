"""privacy.py — AI로 보내기 전 개인정보 가리기.

AI(OpenCode·Claude·ChatGPT)로 글을 보내기 전에 주민등록번호, 전화번호, 이메일,
계좌번호, 아파트 동·호수, 생년월일, 이름(등록한 명단 + '○○○ 학생'처럼 이름이 분명한 자리)을
가린다. UI·DB에 의존하지 않는 순수 로직.

    masked, counts = mask_pii(text, names=['홍길동'], exempt=['031-123-4567'])
"""

import re

# 주민등록번호 / 외국인등록번호: 앞 6자리(생년월일) + 뒤 7자리(첫 숫자 1~8). 하이픈 없어도 잡는다.
_RRN = re.compile(r'(?<!\d)(\d{2})(\d{2})(\d{2})(\s?-\s?|\s?)([1-8])\d{6}(?!\d)')
# 휴대전화: 010-1234-5678, 010 1234 5678, 01012345678, +82 10-1234-5678
_MOBILE = re.compile(r'(?<![\d+])(\+82[-.\s]?|0)(1[016789])[-.\s]?(\d{3,4})[-.\s]?(\d{4})(?!\d)')
# 일반 전화(구분자 있음): 02-123-4567, 031)123-4567, 070-1234-5678
_PHONE = re.compile(r'(?<!\d)(0\d{1,2})[-.)]\s?(\d{3,4})[-.]\s?(\d{4})(?!\d)')
# 일반 전화(구분자 없음): 지역번호(02, 031~064)나 070으로 시작하는 9~11자리 — 0으로 시작해 문서번호·금액과 안 겹침
_PHONE_PLAIN = re.compile(r'(?<!\d)(0(?:2|[3-6][1-5]|70))(\d{3,4})(\d{4})(?!\d)')
_EMAIL = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
# 계좌번호: 숫자 묶음 3~5개를 하이픈으로 이은 것 중 숫자가 10자리 이상 (날짜 2026-11-12 같은 건 제외)
_ACCOUNT = re.compile(r'(?<![\d-])\d{2,6}(?:-\d{1,8}){2,4}(?![\d-])')
# 아파트 동·호수: 101동 1203호
_APT = re.compile(r'(?<!\d)\d{1,4}\s?동\s?\d{1,5}\s?호(?![가-힣])')
# 생년월일: '생년월일', '생일', '출생' 뒤의 날짜
_BIRTH = re.compile(r'(생년월일|생일|출생일?)(\s*[:：]?\s*)'
                    r'(\d{2,4}\s?[.\-/년]\s?\d{1,2}\s?[.\-/월]\s?\d{1,2}\s?일?\.?|\d{8}|\d{6})(?!\d)')

# 흔한 성씨 (등록하지 않은 이름을 문맥으로 찾을 때 사용)
_SURNAMES_2 = ('남궁', '황보', '제갈', '선우', '독고', '사공', '서문', '동방')
_SURNAMES_1 = ('김이박최정강조윤장임한오서신권황안송전홍유고문양손배백허남심노하곽성차주우구민류'
               '나진지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용'
               '예경봉사부가복태목형피두감음빈동온호범좌팽승간상시갈단견당')
_NAME = (r'(?:(?:' + '|'.join(_SURNAMES_2) + r')[가-힣]{1,2}|[' + _SURNAMES_1 + r'][가-힣]{2})')
# 이름 바로 뒤에 붙을 수 있는 조사·호칭 (이것 말고 한글이 붙으면 이름이 아니라 다른 낱말의 일부)
_JOSA = ('에게서', '이에게', '이라고', '입니다', '이에요', '에게', '께서', '한테', '이랑', '하고', '학생', '이가', '이는',
         '이를', '이의', '이와', '이도', '이고', '이다', '에서', '으로', '이', '가', '은', '는', '을', '를', '의',
         '와', '과', '도', '만', '님', '씨', '군', '양', '랑', '아', '야', '로', '에')
_AFTER_NAME = r'(?=(?:' + '|'.join(_JOSA) + r')?(?![가-힣]))'
# 이름이 분명한 자리
_ROLE_AFTER = r'(?:학생|군|양|보호자|학부모|어머니|아버지|어머님|아버님)(?![가-힣])'
_CTX_PATTERNS = [
    # '홍길동 학생', '홍길동 학부모', '홍길동 군'
    re.compile(r'(?<![가-힣])(' + _NAME + r')(\s?)(?=' + _ROLE_AFTER + r')'),
    # 명렬표: '2학년 3반 5번 홍길동', '5번 홍길동, 6번 김철수'
    re.compile(r'((?:\d\s?반\s?)\d{1,2}\s?번\s*)(' + _NAME + r')' + _AFTER_NAME),
    re.compile(r'((?<![\d가-힣])\d{1,2}\s?번\s+)(' + _NAME + r')(?=\s*(?:[,)(/·]|$|\d))', re.M),
    # '성명: 홍길동', '보호자: 홍길동', '예금주 홍길동'
    re.compile(r'((?:성명|이름|학생명|학생 성명|보호자명|보호자|학부모명|학부모|신청자|피해학생|가해학생)'
               r'\s*[:：]\s*|예금주\s*[:：]?\s*)(' + _NAME + r')' + _AFTER_NAME),
]
# 꾸미는 말의 끝 글자 ('신청한 학생', '선발된 학생') — 이름이 아님
_NOT_NAME_ENDINGS = set('한된던할될는운')
# 조사로 끝나는 말 ('3번 문항을') — 이름이 아님
_JOSA_ENDINGS = set('을를에의로')
# 이름처럼 생겼지만 흔한 낱말 (예: '장애인 학생', '한부모 학생')
_NOT_NAMES = {
    '장애인', '한부모', '재학생', '졸업생', '신입생', '전학생', '편입생', '차상위', '수급자', '희망자', '대상자',
    '참가자', '지원자', '합격자', '선발자', '해당자', '우수자', '모범생', '고학년', '저학년', '전학년', '전교생',
    '다문화', '유학생', '기숙사', '공무원', '외국인', '교육생', '연수생', '실습생', '학교장', '이주민', '고등학',
    '중학교', '초등학', '고교생', '중학생', '대학생', '방과후', '기초생', '저소득', '특수교', '일반고', '자사고',
    '특성화', '마이스', '조손가', '북한이', '탈북민', '위기학', '부적응', '전입생', '전출생', '자퇴생', '휴학생',
    '복학생', '어려운', '모든', '전체', '소속', '여러', '해당', '우수', '신입', '이번',
}

_LABELS = {'rrn': '주민등록번호', 'phone': '전화번호', 'email': '이메일', 'account': '계좌번호',
           'address': '동·호수', 'birth': '생년월일', 'name': '이름'}


def _mask_name(name: str) -> str:
    """'홍길동' → '홍○○', '남궁민수' → '남○○○'."""
    return name[0] + '○' * (len(name) - 1) if len(name) > 1 else '○'


def _valid_ymd(mm: str, dd: str) -> bool:
    return 1 <= int(mm) <= 12 and 1 <= int(dd) <= 31


def mask_pii(text: str, names=(), exempt=()):
    """개인정보를 가린 문자열과 종류별 가린 개수를 돌려준다.

    names:  가릴 이름 목록 (2글자 이상만 사용)
    exempt: 가리지 않을 문자열 (학교 대표번호·담당자 연락처·담당자 이름처럼 공문에 들어가야 하는 값)
    """
    if not text:
        return text, {}
    counts = {}
    keep = {e.strip() for e in exempt if e and e.strip()}

    def add(kind, n=1):
        counts[kind] = counts.get(kind, 0) + n

    def sub(kind, pattern, repl, check=None):
        nonlocal text

        def _r(m):
            if m.group(0).strip() in keep or (check and not check(m)):
                return m.group(0)
            add(kind)
            return repl(m)
        text = pattern.sub(_r, text)

    # 숫자 패턴이 겹치므로 순서가 중요: 주민번호 → 생년월일 → 이메일 → 전화 → 계좌
    sub('rrn', _RRN, lambda m: f'{m.group(1)}{m.group(2)}{m.group(3)}-{m.group(5)}******',
        check=lambda m: _valid_ymd(m.group(2), m.group(3)))
    sub('birth', _BIRTH, lambda m: m.group(1) + m.group(2) + '○○○○. ○○. ○○.')
    sub('email', _EMAIL, lambda m: '***@' + m.group(0).split('@', 1)[1])
    sub('phone', _MOBILE, lambda m: f'{m.group(1)}{m.group(2)}-****-{m.group(4)}')
    sub('phone', _PHONE, lambda m: f'{m.group(1)}-****-{m.group(3)}')
    sub('phone', _PHONE_PLAIN, lambda m: f'{m.group(1)}-****-{m.group(3)}')
    sub('account', _ACCOUNT, lambda m: re.sub(r'\d', '*', m.group(0)),
        check=lambda m: sum(c.isdigit() for c in m.group(0)) >= 10)
    sub('address', _APT, lambda m: '○○○동 ○○○호')

    # 등록한 이름: 긴 이름부터 (예: '김민수'가 '김민수진'을 망가뜨리지 않도록).
    # 앞에 한글이 붙어 있거나, 뒤에 조사·호칭이 아닌 한글이 붙으면 다른 낱말의 일부로 보고 건너뛴다 (이수 ↔ 이수증).
    uniq = sorted({n.strip() for n in names if n and len(n.strip()) >= 2}, key=len, reverse=True)
    for n in uniq:
        if n in keep:
            continue
        pat = re.compile(r'(?<![가-힣])' + re.escape(n) + _AFTER_NAME)
        text, hits = pat.subn(_mask_name(n), text)
        if hits:
            add('name', hits)

    # 등록하지 않은 이름: 이름이 분명한 자리에서만 (성씨 + 두 글자)
    for i, pat in enumerate(_CTX_PATTERNS):
        def _r(m, first=(i == 0)):
            name = m.group(1) if first else m.group(2)
            if (name in keep or name in _NOT_NAMES or '○' in name
                    or name[-1] in _JOSA_ENDINGS or (first and name[-1] in _NOT_NAME_ENDINGS)):
                return m.group(0)
            add('name')
            if first:
                return _mask_name(name) + m.group(2)
            return m.group(1) + _mask_name(name)
        text = pat.sub(_r, text)
    return text, counts


MASK_MARKERS = ('○○', '****', '***@')


def looks_masked(text: str) -> bool:
    """가려진 표시(○○, ****)가 들어 있는지 — AI가 가린 글을 그대로 되돌려 쓰는 것을 막을 때 사용."""
    return bool(text) and any(m in text for m in MASK_MARKERS)


def describe(counts: dict) -> str:
    """{'phone': 2, 'name': 1} → '전화번호 2건, 이름 1건'."""
    return ', '.join(f'{_LABELS.get(k, k)} {v}건' for k, v in counts.items() if v)


def parse_name_list(text: str) -> list:
    """쉼표·줄바꿈·공백으로 구분된 이름 목록을 리스트로."""
    return [n for n in re.split(r'[,\s]+', text or '') if len(n) >= 2]
