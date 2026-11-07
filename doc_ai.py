"""doc_ai.py — 캡처한 공문을 AI로 분석해 요약을 만든다.

흐름
  1) 캡처 이미지 → (개인정보 보호 모드) PC의 Windows OCR로 글자 추출 → 개인정보 가리기 → AI에 글자만 전송
                 → (보호 모드 꺼짐 / 로컬 AI) 이미지를 그대로 AI에 전송
  2) AI 응답(JSON) → normalize_analysis()로 날짜·목록 정리
  3) 제목·공문번호 → 공문 작성 칸(제목·관련 번호), 요약·기한 → 받은 공문 칸의 요약 (요약 복사)

UI 없는 순수 로직(build_*, normalize_*, parse_*)은 tests/에서 검증한다.
"""

import json
import re
from datetime import date, timedelta

from PyQt5.QtCore import QThread, pyqtSignal

ANALYZE_SYSTEM = (
    "당신은 한국 학교 행정 공문 분석 전문가입니다. "
    "공문을 읽고 핵심 내용을 정확하게 정리합니다. "
    "공문에 없는 내용은 지어내지 말고 빈 값으로 둡니다. 반드시 JSON만 출력합니다."
)

_ANALYZE_FORMAT = """{
  "title": "공문 제목 (번호·기관명 제외한 제목만)",
  "doc_number": "문서번호의 부서명-번호만 (기관명 제외, 예: 중등교육과-12345). 없으면 빈 문자열",
  "doc_date": "시행일자 YYYY-MM-DD (문서번호 옆이나 시행 칸의 날짜). 없으면 빈 문자열",
  "deadline": "학교가 제출·회신·완료해야 하는 기한 YYYY-MM-DD. 없으면 빈 문자열",
  "deadline_time": "기한 시각 HH:MM (예: 17:00). 없으면 빈 문자열",
  "summary": ["핵심 내용 한 줄", "최대 3줄"],
  "needs_reply": true
}"""


def build_analyze_prompt(text: str = '', today: date = None) -> str:
    """분석 프롬프트. text가 있으면 글자 기반, 없으면 첨부 이미지 기반."""
    today = today or date.today()
    src = f"[공문 내용]\n{text.strip()}\n\n" if text and text.strip() else "첨부한 이미지는 한국 학교 행정 공문입니다.\n\n"
    return (
        f"{src}"
        f"오늘 날짜: {today.isoformat()} (연도가 없는 날짜는 오늘 기준으로 가장 가까운 해로 판단)\n"
        "needs_reply: 학교가 공문으로 제출·보고·회신해야 하면 true, 안내·참고용이면 false.\n"
        "아래 JSON 형식으로만 답하세요. 설명·마크다운 금지.\n"
        f"{_ANALYZE_FORMAT}"
    )


# ── 응답 파싱 / 정리 ──────────────────────────────────────────────

def parse_json_object(raw: str) -> dict:
    """AI 응답에서 첫 JSON 객체를 꺼낸다. 실패하면 ValueError."""
    if not raw:
        raise ValueError('AI 응답이 비어 있습니다.')
    m = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', raw, re.DOTALL)
    candidates = [m.group(1)] if m else []
    start, end = raw.find('{'), raw.rfind('}')
    if start != -1 and end > start:
        candidates.append(raw[start:end + 1])
    for c in candidates:
        try:
            obj = json.loads(c)
            if isinstance(obj, dict):
                return obj
        except ValueError:
            continue
    raise ValueError('AI 응답에서 JSON을 찾지 못했습니다.')


_DATE_PATTERNS = [
    re.compile(r'(?P<y>\d{4})\s*[-./년]\s*(?P<m>\d{1,2})\s*[-./월]\s*(?P<d>\d{1,2})'),
    re.compile(r'(?<!\d)(?P<m>\d{1,2})\s*[./월]\s*(?P<d>\d{1,2})\s*[.일]?'),
]


def normalize_date(value, today: date = None) -> str:
    """'2026. 11. 12.(목)', '11월 12일', '2026-11-12' → '2026-11-12'. 못 읽으면 ''.

    연도가 없으면 올해로 보되, 오늘보다 60일 넘게 지난 날짜면 내년으로 본다
    (12월에 받은 '1월 10일까지' 공문).
    """
    if not value or not isinstance(value, str):
        return ''
    today = today or date.today()
    s = value.strip()
    for i, pat in enumerate(_DATE_PATTERNS):
        m = pat.search(s)
        if not m:
            continue
        try:
            month, day = int(m.group('m')), int(m.group('d'))
            if i == 0:
                return date(int(m.group('y')), month, day).isoformat()
            d = date(today.year, month, day)
            if d < today - timedelta(days=60):
                d = date(today.year + 1, month, day)
            return d.isoformat()
        except ValueError:
            continue
    return ''


def normalize_time(value) -> str:
    """'17:00', '17시', '오후 5시 30분' → 'HH:MM'. 못 읽으면 ''."""
    if not value or not isinstance(value, str):
        return ''
    s = value.strip()
    m = re.search(r'(\d{1,2})\s*[:시]\s*(\d{1,2})?', s)
    if not m:
        return ''
    h, mi = int(m.group(1)), int(m.group(2) or 0)
    if '오후' in s and h < 12:
        h += 12
    if not (0 <= h <= 23 and 0 <= mi <= 59):
        return ''
    return f'{h:02d}:{mi:02d}'


def _str_list(value, limit):
    if isinstance(value, str):
        value = [v for v in re.split(r'\n+', value)]
    if not isinstance(value, list):
        return []
    out = []
    for v in value:
        v = re.sub(r'^\s*(?:[-•·*]|\d+[.)])\s*', '', str(v)).strip()
        if v:
            out.append(v)
    return out[:limit]


def normalize_analysis(data: dict, today: date = None) -> dict:
    """AI가 준 dict를 화면에서 바로 쓸 수 있게 정리 (누락 키 채움, 날짜 형식 통일)."""
    data = data if isinstance(data, dict) else {}

    def text(k):
        v = data.get(k, '')
        return v.strip() if isinstance(v, str) else ''

    needs = data.get('needs_reply', False)
    if isinstance(needs, str):
        needs = needs.strip().lower() in ('true', 'yes', '1', '예', '필요')
    doc_date = normalize_date(text('doc_date'), today)
    return {
        'title': text('title'),
        'doc_number': format_doc_number(text('doc_number'), doc_date),   # '중등교육과-12345(2026. 11. 3.)'
        'doc_date': doc_date,
        'deadline': normalize_date(text('deadline'), today),
        'deadline_time': normalize_time(text('deadline_time')),
        'summary': _str_list(data.get('summary'), 3),
        'needs_reply': bool(needs),
    }


# ── 결과 활용 ────────────────────────────────────────────────────

def format_korean_date(iso: str) -> str:
    """'2026-11-05' → '2026. 11. 5.' (공문 날짜 표기)."""
    try:
        y, m, d = (int(x) for x in iso.split('-'))
        return f'{y}. {m}. {d}.'
    except (ValueError, AttributeError):
        return ''


_DOC_NO = re.compile(r'([가-힣A-Za-z0-9·]+?)\s*-\s*(\d[\d\s]*)')


def format_doc_number(num: str, iso_date: str = '') -> str:
    """공문번호를 '중등교육과-12345(2026. 11. 3.)' 형식으로 맞춘다.

    기관명('경기도교육청 ')은 빼고, 괄호 안 날짜가 있으면 그 날짜를, 없으면 iso_date(시행일)를 붙인다.
    날짜는 '2026. 11. 3.'(월·일 앞 0 없음, 마침표 뒤 한 칸). 번호 모양이 아니면 그대로 돌려준다.
    """
    s = (num or '').strip()
    if not s:
        return ''
    s = s.translate(str.maketrans('（）【】〔〕［］', '()()()()'))
    for ch in '—–‑−':
        s = s.replace(ch, '-')
    s = re.sub(r'\s+', ' ', s)
    date_part = ''
    m = re.search(r'\(([^)]*)\)', s)
    if m:
        date_part, s = m.group(1), (s[:m.start()] + s[m.end():]).strip()
    hits = list(_DOC_NO.finditer(s))
    if not hits:
        return (num or '').strip()
    dept, number = hits[-1].group(1), re.sub(r'\s', '', hits[-1].group(2))
    d = ''
    nums = re.findall(r'\d+', date_part)
    if len(nums) >= 3:
        y, mo, dd = int(nums[0]), int(nums[1]), int(nums[2])
        if y < 100:
            y += 2000
        if 1 <= mo <= 12 and 1 <= dd <= 31:
            d = f'{y}. {mo}. {dd}.'
    if not d:
        d = format_korean_date(iso_date)
    return f'{dept}-{number}({d})' if d else f'{dept}-{number}'


def related_line(a: dict) -> str:
    """'관련' 칸에 넣을 문구: '중등교육과-12345(2026. 11. 5.)'."""
    return format_doc_number(a.get('doc_number', ''), a.get('doc_date', ''))


def summary_body(a: dict) -> str:
    """받은 공문 칸의 요약 칸에 보일 글: 핵심 요약 · 기한."""
    lines = [f'· {x}' for x in a.get('summary', [])]
    if a.get('deadline'):
        lines.append(f"· 기한: {format_korean_date(a['deadline'])} {a.get('deadline_time', '')}".rstrip())
    return '\n'.join(lines)


def summary_text(title: str, rel: str, body: str) -> str:
    """클립보드용 요약 (메신저·회의 공유용): '제목 (관련 번호)' + 요약 칸 글."""
    head = (title or '').strip()
    rel = (rel or '').strip()
    if rel:
        head = f'{head} ({rel})' if head else rel
    return '\n'.join(x for x in (head, (body or '').strip()) if x).strip()


def context_block(title: str, rel: str, body: str) -> str:
    """AI 대화·자동 초안에 함께 보낼 '받은 공문' 정리 (공문 작성 창 왼쪽 칸)."""
    lines = [f'받은 공문 제목: {title}'] if title else []
    if rel:
        lines.append(f'받은 공문 번호: {rel}')
    if body.strip():
        lines.append('요약:\n' + body.strip())
    return '[참고: 받은 공문]\n' + '\n'.join(lines) + '\n'


# ── 분석 실행 (백그라운드 스레드) ─────────────────────────────────

class DocAnalysisWorker(QThread):
    """캡처 이미지(또는 이미 읽은 글자)를 분석한다.

    analyzed(result: dict, info: dict)  info = {'mode': 'text'|'image', 'masked': {...}, 'ocr_text': str}
    failed(message: str)
    progress(message: str)
    """
    analyzed = pyqtSignal(dict, dict)   # QThread.finished와 이름이 겹치지 않게
    failed = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, pixmap=None, ocr_text: str = '', parent=None):
        super().__init__(parent)
        self._image_b64 = None
        if pixmap is not None:
            from ai_client import _pixmap_to_base64
            self._image_b64 = _pixmap_to_base64(pixmap)
        self._ocr_text = ocr_text

    def run(self):
        try:
            from ai_client import get_adapter, friendly_error
            from school_profile import privacy_enabled, mask_for_ai
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있습니다. 메뉴 → 설정 → AI 연결(API 키)에서 사용할 AI를 고르세요.')
                return
            if not adapter.get_api_key():
                self.failed.emit('AI API 키가 설정되지 않았습니다. 메뉴 → 설정 → AI 연결(API 키)에서 키를 입력하세요.')
                return

            info = {'mode': 'image', 'masked': {}, 'ocr_text': self._ocr_text}
            use_text = bool(self._ocr_text) or (privacy_enabled())
            if use_text:
                text = self._ocr_text
                if not text:
                    self.failed.emit('글자를 읽지 못했습니다.')
                    return
                if privacy_enabled():
                    text, info['masked'] = mask_for_ai(text)
                info['mode'] = 'text'
                self.progress.emit('AI가 공문을 읽는 중…')
                raw = adapter.complete(ANALYZE_SYSTEM, build_analyze_prompt(text))
            else:
                self.progress.emit('AI가 이미지를 읽는 중…')
                raw = adapter.complete(ANALYZE_SYSTEM, build_analyze_prompt(), image_b64=self._image_b64)
            self.analyzed.emit(normalize_analysis(parse_json_object(raw)), info)
        except Exception as e:  # 네트워크·키·응답 형식 오류 모두 사용자에게 표시
            self.failed.emit(friendly_error(e) if getattr(e, 'status_code', None) else str(e))
