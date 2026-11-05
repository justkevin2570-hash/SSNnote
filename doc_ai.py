"""doc_ai.py — 캡처한 공문을 AI로 분석하고, 회신 공문 프롬프트를 만든다.

흐름
  1) 캡처 이미지 → (개인정보 보호 모드) PC의 Windows OCR로 글자 추출 → 개인정보 가리기 → AI에 글자만 전송
                 → (보호 모드 꺼짐 / 로컬 AI) 이미지를 그대로 AI에 전송
  2) AI 응답(JSON) → normalize_analysis()로 날짜·목록 정리
  3) 결과 창에서 업무 등록 / 회신 공문 작성 / 요약 복사

UI 없는 순수 로직(build_*, normalize_*, parse_*)은 tests/에서 검증한다.
"""

import json
import re
from datetime import date, timedelta

from PyQt5.QtCore import QThread, pyqtSignal

ANALYZE_SYSTEM = (
    "당신은 한국 학교 행정 공문 분석 전문가입니다. "
    "공문을 읽고 학교 담당 교사가 해야 할 일을 정확하게 정리합니다. "
    "공문에 없는 내용은 지어내지 말고 빈 값으로 둡니다. 반드시 JSON만 출력합니다."
)

_ANALYZE_FORMAT = """{
  "title": "공문 제목 (번호·기관명 제외한 제목만)",
  "doc_number": "문서번호 (예: 경기도교육청 중등교육과-12345). 없으면 빈 문자열",
  "doc_date": "시행일자 YYYY-MM-DD. 없으면 빈 문자열",
  "sender": "발신 기관/부서. 없으면 빈 문자열",
  "deadline": "학교가 제출·회신·완료해야 하는 기한 YYYY-MM-DD. 없으면 빈 문자열",
  "deadline_time": "기한 시각 HH:MM (예: 17:00). 없으면 빈 문자열",
  "summary": ["핵심 내용 한 줄", "최대 3줄"],
  "todos": ["학교(담당자)가 해야 할 일. 동사로 끝나게. 최대 5개"],
  "attachments": ["붙임 파일 이름"],
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
    return {
        'title': text('title'),
        'doc_number': text('doc_number'),
        'doc_date': normalize_date(text('doc_date'), today),
        'sender': text('sender'),
        'deadline': normalize_date(text('deadline'), today),
        'deadline_time': normalize_time(text('deadline_time')),
        'summary': _str_list(data.get('summary'), 3),
        'todos': _str_list(data.get('todos'), 5),
        'attachments': _str_list(data.get('attachments'), 10),
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


def related_line(a: dict) -> str:
    """'관련' 칸에 넣을 문구: '경기도교육청 중등교육과-12345(2026. 11. 5.)'."""
    num = a.get('doc_number', '')
    d = format_korean_date(a.get('doc_date', ''))
    if num and d:
        return f'{num}({d})'
    return num


def task_deadline(a: dict) -> str:
    """업무 마감 문자열 (db.tasks.deadline 형식: 'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM')."""
    if not a.get('deadline'):
        return ''
    return f"{a['deadline']} {a['deadline_time']}".strip() if a.get('deadline_time') else a['deadline']


def task_notes(a: dict) -> str:
    """업무 메모 칸에 넣을 정리본."""
    parts = []
    if a.get('sender'):
        parts.append(f"발신: {a['sender']}")
    if a.get('summary'):
        parts.append('요약\n' + '\n'.join(f'- {s}' for s in a['summary']))
    if a.get('todos'):
        parts.append('할 일\n' + '\n'.join(f'☐ {t}' for t in a['todos']))
    if a.get('attachments'):
        parts.append('붙임: ' + ', '.join(a['attachments']))
    return '\n\n'.join(parts)


def summary_text(a: dict) -> str:
    """클립보드용 요약 (메신저·회의 공유용)."""
    head = a.get('title', '')
    if a.get('doc_number'):
        head += f" ({a['doc_number']})"
    lines = [head] if head else []
    if a.get('deadline'):
        lines.append(f"기한: {format_korean_date(a['deadline'])} {a.get('deadline_time', '')}".rstrip())
    body = task_notes(a)
    if body:
        lines.append(body)
    return '\n'.join(lines).strip()


REPLY_SYSTEM = (
    "당신은 한국 학교 행정 공문 작성 전문가입니다. "
    "받은 공문에 대한 학교의 회신(제출·보고) 공문 본문을 작성합니다. "
    "공식 공문 문체(개조식, '~함', '~바람', '~하고자 합니다')를 쓰고, "
    "날짜는 '2026. 3. 12.' 형식, 마지막은 '끝.'으로 맺습니다. "
    "공문에 없는 수치·명단은 지어내지 말고 '○○'로 비워 둡니다. 마크다운 금지."
)


def build_reply_prompt(a: dict, profile_text: str = '', instruction: str = '') -> str:
    """회신 공문 본문 생성 프롬프트 (AiStreamThread structured 모드용: {"body": ...})."""
    rel = related_line(a)
    info = [f"받은 공문 제목: {a.get('title', '')}"]
    if rel:
        info.append(f'관련 문서: {rel}')
    if a.get('sender'):
        info.append(f"발신 기관: {a['sender']}")
    if a.get('deadline'):
        info.append(f"제출 기한: {format_korean_date(a['deadline'])} {a.get('deadline_time', '')}".rstrip())
    if a.get('summary'):
        info.append('받은 공문 요약:\n' + '\n'.join(f'- {s}' for s in a['summary']))
    if a.get('todos'):
        info.append('요구 사항:\n' + '\n'.join(f'- {t}' for t in a['todos']))
    if a.get('attachments'):
        info.append('받은 공문 붙임: ' + ', '.join(a['attachments']))
    parts = ['\n'.join(info)]
    if profile_text:
        parts.append(profile_text.strip())
    if instruction.strip():
        parts.append(f'추가 요청: {instruction.strip()}')
    parts.append(
        '본문 형식:\n'
        f"1. 관련: {rel or '[관련 문서]'}\n"
        '2. 위와 관련하여 [무엇]을 다음과 같이 [제출/보고]합니다.\n'
        '  가. ...\n'
        '붙임  [회신 붙임 파일명] 1부.  끝.\n\n'
        '반드시 JSON 형식({"body": "공문 본문"})으로만 출력하세요.'
    )
    return '\n\n'.join(parts)


def reply_title(a: dict) -> str:
    """회신 공문 제목 초안: 받은 제목 끝에 '제출'을 붙인다 (이미 제출/보고/회신이면 그대로)."""
    t = a.get('title', '').strip()
    if not t or re.search(r'(제출|보고|회신|결과)$', t):
        return t
    t = re.sub(r'\s*(안내|알림|요청|협조\s*요청|조사)$', '', t)
    return f'{t} 제출'


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
            from ai_client import get_adapter, is_local_adapter
            from school_profile import privacy_enabled, mask_for_ai
            adapter = get_adapter()
            if adapter is None:
                self.failed.emit('AI가 꺼져 있습니다. 메뉴 → 설정 → AI 연결(API 키)에서 사용할 AI를 고르세요.')
                return
            if getattr(adapter, 'requires_api_key', True) and not adapter.get_api_key():
                self.failed.emit('AI API 키가 설정되지 않았습니다. 메뉴 → 설정 → AI 연결(API 키)에서 키를 입력하세요.')
                return

            info = {'mode': 'image', 'masked': {}, 'ocr_text': self._ocr_text}
            use_text = bool(self._ocr_text) or (privacy_enabled() and not is_local_adapter(adapter))
            if use_text:
                text = self._ocr_text
                if not text:
                    self.failed.emit('글자를 읽지 못했습니다.')
                    return
                if privacy_enabled() and not is_local_adapter(adapter):
                    text, info['masked'] = mask_for_ai(text)
                info['mode'] = 'text'
                self.progress.emit('AI가 공문을 읽는 중…')
                raw = adapter.complete(ANALYZE_SYSTEM, build_analyze_prompt(text))
            else:
                self.progress.emit('AI가 이미지를 읽는 중…')
                raw = adapter.complete(ANALYZE_SYSTEM, build_analyze_prompt(), image_b64=self._image_b64)
            self.analyzed.emit(normalize_analysis(parse_json_object(raw)), info)
        except Exception as e:  # 네트워크·키·응답 형식 오류 모두 사용자에게 표시
            self.failed.emit(str(e))
