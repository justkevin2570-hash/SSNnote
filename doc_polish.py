"""doc_polish.py — 공문 선택 부분 다듬기의 순수 로직 (UI 없음).

  · 프롬프트: 공문 문체 유지, 의미·수치·날짜 바꾸지 않기, 작성 규칙·예시 함께, 다듬은 글만 출력.
  · 개인정보 가리기 되살리기: 가린 글을 AI에 보내고, 가리기 전·후 단어를 1:1로 맞춰
    AI 답에 남은 가린 표시(○○·****)를 원래 값으로 바꾼다. 맞출 수 없으면 적용을 막는다.
  · 비교: 원문과 다듬은 글에서 바뀐 단어의 위치(글자 범위)를 구한다.
"""

import difflib
import re

POLISH_SYSTEM = (
    "당신은 한국 학교 행정 공문 문장을 다듬는 교정 전문가입니다.\n"
    "- 주어진 글을 공문 문체(개조식, '~함', '~바람', '~고자 함', '~하고자 합니다')로 자연스럽게 고칩니다.\n"
    "- 의미·수치·날짜·시간·금액·장소·인원·공문번호는 절대 바꾸지 않습니다. 없는 내용을 덧붙이거나 빼지 않습니다.\n"
    "- 가나다·1)·가. 같은 항목 번호, 들여쓰기, 줄바꿈 구조, <table> 태그는 그대로 둡니다.\n"
    "- ○○, **** 처럼 가려진 부분은 글자 그대로 두고 바꾸지 않습니다.\n"
    "- 다듬은 글만 출력합니다. 설명·머리말·따옴표·마크다운(**, ##) 사용 금지."
)

_TOKEN = re.compile(r'\s+|[^\s]+')
_MASK_MARK = re.compile(r'○|\*\*\*')


def build_polish_prompt(text: str, notes: str = '', examples: str = '') -> str:
    """다듬을 글 → AI에 보낼 질문. notes·examples는 doc_style이 만든 글(없으면 빈 문자열)."""
    return f"{notes}{examples}\n[다듬을 글]\n{text}\n\n다듬은 글만 출력하세요."


def _tokens(text: str) -> list:
    return _TOKEN.findall(text or '')


def _spans(tokens: list, indexes) -> list:
    """토큰 목록에서 주어진 토큰 번호들의 글자 범위 [(start, end)] (공백만 있는 토큰은 뺌)."""
    offsets, pos = [], 0
    for t in tokens:
        offsets.append((pos, pos + len(t)))
        pos += len(t)
    out = []
    for i in indexes:
        s, e = offsets[i]
        if tokens[i].strip():
            out.append((s, e))
    return out


def diff_spans(original: str, polished: str):
    """바뀐 단어의 글자 범위. 반환: (원문에서 빠진 부분, 다듬은 글에서 새로 들어간 부분)."""
    a, b = _tokens(original), _tokens(polished)
    left, right = [], []
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ('delete', 'replace'):
            left += _spans(a, range(i1, i2))
        if tag in ('insert', 'replace'):
            right += _spans(b, range(j1, j2))
    return left, right


def restore_masked(polished: str, original: str, masked: str):
    """AI 답의 가린 표시를 원래 값으로 바꾼다. 반환: (글, 되살렸는지).
    가리기 전(original)과 후(masked)의 단어 수가 같을 때만 되살린다 (단어 1:1 대응)."""
    if original == masked:
        return polished, True
    ot, mt = _tokens(original), _tokens(masked)
    if len(ot) != len(mt):
        return polished, False
    pairs = {m: o for m, o in zip(mt, ot) if m != o and m.strip()}
    text = polished
    for m in sorted(pairs, key=len, reverse=True):
        text = text.replace(m, pairs[m])
    return text, True


def leftover_masks(text: str) -> bool:
    """되살린 뒤에도 가린 표시가 남아 있는지 (AI가 새로 가린 경우)."""
    return bool(_MASK_MARK.search(text or ''))
