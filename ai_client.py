"""
ai_client.py — AI 제공자 추상화 레이어

지원 제공자: Gemini, Claude, OpenAI(ChatGPT), NVIDIA NIM, Ollama(로컬)
다른 제공자로 교체 시 이 파일의 어댑터만 수정하면 됩니다.
"""

import os
import base64
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QPixmap

APPDATA_DIR        = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
KEY_FILE           = os.path.join(APPDATA_DIR, 'gemini_key.txt')
AI_MODE_FILE       = os.path.join(APPDATA_DIR, 'ai_mode.txt')
OLLAMA_MODEL_FILE  = os.path.join(APPDATA_DIR, 'ollama_model.txt')
GEMINI_MODEL_FILE  = os.path.join(APPDATA_DIR, 'gemini_model.txt')
CLAUDE_KEY_FILE    = os.path.join(APPDATA_DIR, 'claude_key.txt')
CLAUDE_MODEL_FILE  = os.path.join(APPDATA_DIR, 'claude_model.txt')
OPENAI_KEY_FILE    = os.path.join(APPDATA_DIR, 'openai_key.txt')
OPENAI_MODEL_FILE  = os.path.join(APPDATA_DIR, 'openai_model.txt')
NVIDIA_KEY_FILE    = os.path.join(APPDATA_DIR, 'nvidia_key.txt')
NVIDIA_MODEL_FILE  = os.path.join(APPDATA_DIR, 'nvidia_model.txt')
OPENCODE_KEY_FILE   = os.path.join(APPDATA_DIR, 'opencode_key.txt')
OPENCODE_MODEL_FILE = os.path.join(APPDATA_DIR, 'opencode_model.txt')
OLLAMA_HOST        = 'http://localhost:11434'

# 세션 키 (API 키 저장 체크 해제 시 메모리에만 보관)
_session_keys: dict[str, str] = {}

# AI 모드: 'gemini'|'claude'|'openai'|'nvidia'|'opencode'|'internal'|'none'
# 'external'은 하위 호환 (= 'gemini')
_AI_MODES = ('external', 'internal', 'none', 'gemini', 'claude', 'openai', 'nvidia', 'opencode')

# 기본 모델 목록 (API 키 없이도 표시)
GEMINI_DEFAULT_MODELS = [
    "gemini-3-flash-preview",
    "gemini-2.0-flash",
    "gemini-2.5-pro-preview-05-06",
]

CLAUDE_DEFAULT_MODELS = [
    "claude-sonnet-4-6",
    "claude-opus-4-7",
    "claude-haiku-4-5-20251001",
]

OPENAI_DEFAULT_MODELS = [
    "gpt-4.1",
    "gpt-4o",
    "gpt-4o-mini",
]

NVIDIA_DEFAULT_MODELS = [
    "nvidia/llama-3.1-nemotron-ultra-253b-v1",
    "meta/llama-3.1-405b-instruct",
    "meta/llama-3.3-70b-instruct",
    "mistralai/mixtral-8x7b-instruct-v0.1",
]

OPENCODE_DEFAULT_MODELS = [
    "glm-5.1",
    "glm-5",
    "kimi-k2.5",
    "kimi-k2.6",
    "deepseek-v4-pro",
    "deepseek-v4-flash",
    "mimo-v2-pro",
    "mimo-v2-omni",
    "mimo-v2.5-pro",
    "mimo-v2.5",
    "qwen3.6-plus",
    "qwen3.5-plus",
]

# 외부 AI(대형 모델): 형식 제약 최소화, 모델 자율성 최대
EXTERNAL_DOC_SYSTEM_PROMPT = (
    "당신은 한국 학교 행정 공문 작성 전문가입니다. "
    "공식 공문 문체(개조식, '~함', '~바람', '~고자 함')로 작성하세요. "
    "불필요한 설명 없이 공문 본문만 작성하세요. "
    "마크다운(**·##) 사용 금지."
)

# 내부 AI(Ollama 소형 모델): 형식 고정, few-shot으로 틀 강제
INTERNAL_DOC_SYSTEM_PROMPT = """당신은 한국 학교 행정 공문 작성 전문가입니다.
제목과 문서유형이 주어지면 뼈대의 빈칸(: 뒤 공란)만 채워 공문 본문을 완성하세요.
형식(가~마 구조, 붙임 문구) 절대 변경 금지. 설명·인사말·마크다운(**·##) 절대 금지.

=== 유형별 작성 예시 ===

[일반 계획/운영]
1. 2026. 체육대회를 다음과 같이 운영하고자 합니다.

  가. 일시: 2026. 5. 15.(금) 09:00~17:00
  나. 장소: 학교 운동장
  다. 대상: 전교생
  라. 내용: 육상, 단체줄넘기, 응원전 등

붙임  체육대회 계획 1부.  끝.

[결과 보고]
1. 2025. 북세통 프로젝트 사전답사 결과를 다음과 같이 보고합니다.

  가. 일시: 2025. 9. 10.(수) 10:00~15:00
  나. 장소: 전남 순천시 일원
  다. 인원: 인솔교사 2명, 학생 25명
  라. 내용: 사전답사 경로 확인 및 안전 점검

붙임  북세통 프로젝트 사전답사 결과 1부.  끝.

[가정통신문 발송]
1. 2026. 청소년 미디어 이용습관 진단조사 가정통신문 발송을 붙임과 같이 발송하고자 합니다.

붙임  청소년 미디어 이용습관 진단조사 가정통신문 1부.  끝.

[수요조사 결과 제출]
1. 2026년도 공무원 맞춤형복지 단체보험 희망상품 수요조사 결과를 붙임과 같이 제출하고자 합니다.

붙임  기관별 공무원 단체보험 수요 현황 조사 결과(서식) 1부.  끝.

[채용 공고]
1. 2026학년도 시간강사(수학) 채용 계획을 다음과 같이 공고하고자 합니다.

  가. 채용예정인원 및 과목:
      구분 | 채용 과목 | 선발예정 인원 | 채용기간 | 비고
      학교자율사업 | 수학 | 1명 | 2026.3.17.~2026.5.29. | 주2일근무

  나. 접수기간: 2026.3.10.(화) 11:00 ~ 2026.3.12.(목) 14:00
  다. 접수장소: 고흥고등학교 창의예술관2층 교무실(전자우편, 우편, 방문접수)
  라. 공고방법: 인터넷 홈페이지(전라남도교육청)
  마. 공고일자: 2026.3.10.(목) 11:00

붙임  1. 시간강사(수학) 채용 공고문 1부.
      2. 서류평가 및 면접평가표 1부.(별첨)
      3. 서류 및 면접 전형 평가표 1부.(별첨)  끝.

[채용 결정]
1. 고교학점제(선택과목) 시간강사(윤리)를 다음과 같이 채용하고자 합니다.

  가. 채용 대상자:
      연번 | 학교명 | 성명 | 생년월일 | 자격증 | 운용요일 | 비고
      1 | 고흥고 | 홍길동 | 1980-01-01 | 중등2정 | 월,목,금 | -

  나. 강사 채용 기간: 2026.3.1. ~ 2026.12.31.(여름방학 제외)
  다. 강사 채용 과목: 윤리
  라. 강사 근무 시간: 주3일, 주당 9시간(월:3시간, 목:3시간, 금:3시간)
  마. 강사 수당: 1시간 40,000원

붙임  1. 서류(적부) 전형 평가표 1부.(별첨)
      2. 응시원서 및 자기소개서 1부.(별첨)
      3. 채용계약서 1부.(별첨)  끝.
=================

위 예시 형식 그대로 출력하되, 빈칸은 제목과 문서유형에 맞게 채우세요."""

# 하위 호환용 별칭
DOC_SYSTEM_PROMPT = EXTERNAL_DOC_SYSTEM_PROMPT

# ── 외부 AI 퓨샷 시스템 프롬프트 (유형별) ───────────────────────────

_FEWSHOT_RULES = """
들여쓰기 규칙:
- 가나다 항목: 앞에 공백 2칸
- 1)2)3) 항목: 앞에 공백 4칸, 첫 번째(1))도 동일

2번 문장 생성 규칙:
- 제목에서 마지막 명사(계획/요구/실시 등) 제거
- 남은 핵심 명사 + 를/을 다음과 같이 + 핵심동사 + 하고자 합니다.
- 예) 교육과정 박람회 운영 계획 → 교육과정 박람회를 다음과 같이 운영하고자 합니다.
- 예) 현장체험학습 실시 계획 → 현장체험학습을 다음과 같이 실시하고자 합니다.
- 예) 급식비 지출 요구 → 급식비를 다음과 같이 지출하고자 합니다.

줄 간격 규칙:
- 2번 문장 바로 다음 줄부터 가나다 항목 시작 (빈 줄 없음)

붙임 형식 규칙:
- 붙임 1개:  붙임  파일명 1부.  끝.
- 붙임 2개 이상: 번호를 붙이고 2번째 항목부터 앞에 공백 6칸, 끝.은 마지막 줄 끝에 인라인
  붙임  1. 파일명1 1부.
        2. 파일명2 1부.  끝.
"""

FEWSHOT_PLAN_SYSTEM = (
    "당신은 한국 학교 행정 공문 본문 작성 전문가입니다.\n"
    "아래 규칙과 예시를 따라 공문 본문만 작성하세요. 마크다운(**·##) 사용 금지.\n"
    + _FEWSHOT_RULES +
    "\n[예시]\n"
    "입력:\n"
    "제목: 2026. 교육과정 박람회 운영 계획\n"
    "관련: 2026. 학교교육계획\n"
    "\n출력:\n"
    "1. 관련: 2026. 학교교육계획\n"
    "2. 2026. 교육과정 박람회를 다음과 같이 운영하고자 합니다.\n"
    "  가. 일시: 2026. 5. 13.(수) 13:00 ~ 16:30\n"
    "  나. 장소: 본교 체육관 및 각 교과별 지정 교실\n"
    "  다. 대상: 전교생, 학부모 및 전교직원\n"
    "  라. 내용\n"
    "    1) 2022 개정 교육과정 및 선택 과목 이수 가이드 안내\n"
    "    2) 교과별 부스 운영을 통한 과목 상담 및 학습 설계 지원\n"
    "    3) 진로 연계 선택 과목 탐색 및 1:1 맞춤형 컨설팅 실시\n"
    "    4) 교육과정 관련 자료집 배부 및 홍보 영상 상영\n"
    "\n붙임  교육과정 박람회 계획 1부.  끝."
)

FEWSHOT_PURCHASE_SYSTEM = (
    "당신은 한국 학교 행정 공문 본문 작성 전문가입니다.\n"
    "아래 규칙과 예시를 따라 공문 본문만 작성하세요. 마크다운(**·##) 사용 금지.\n"
    + _FEWSHOT_RULES +
    "\n추가 규칙:\n"
    "- 끝.은 붙임 라인 없이 마지막 항목(예상금액) 뒤에 인라인으로 표시\n"
    "\n[예시]\n"
    "입력:\n"
    "제목: 2025. 독서 프로젝트 교재 구입 지출 요구\n"
    "관련: 한국고-12345(2025. 3. 12.)\n"
    "\n출력:\n"
    "1. 관련: 한국고-12345(2025. 3. 12.)\n"
    "2. 2025. 독서 프로젝트 교재 구입을 다음과 같이 지출하고자 합니다.\n"
    "  가. 일시: 2025. 12. 17.(수) ~ 12. 19.(금)\n"
    "  나. 장소: 수도권 일대\n"
    "  다. 대상: 독서 프로젝트 참가자\n"
    "  라. 내용: 독서 프로젝트 출판사 인터뷰 도서\n"
    "  마. 내역: 도서 외 3종\n"
    "  바. 예상금액: 350,000원.  끝."
)

FEWSHOT_DEFAULT_SYSTEM = (
    "당신은 한국 학교 행정 공문 본문 작성 전문가입니다.\n"
    "아래 규칙과 예시를 따라 공문 본문만 작성하세요. 마크다운(**·##) 사용 금지.\n"
    + _FEWSHOT_RULES +
    "\n[예시]\n"
    "입력:\n"
    "제목: 2026. 현장체험학습 실시 안내\n"
    "관련: 2026. 학교교육계획\n"
    "\n출력:\n"
    "1. 관련: 2026. 학교교육계획\n"
    "2. 2026. 현장체험학습을 다음과 같이 실시하고자 합니다.\n"
    "  가. 일시: 2026. 4. 10.(목)\n"
    "  나. 장소: 경주 일대\n"
    "  다. 대상: 2학년 전체\n"
    "  라. 내용\n"
    "    1) 불국사·석굴암 문화유산 탐방\n"
    "    2) 국립경주박물관 체험 학습\n"
    "\n붙임  현장체험학습 계획서 1부.  끝."
)

FEWSHOT_SUBMIT_SYSTEM = (
    "당신은 한국 학교 행정 공문 본문 작성 전문가입니다.\n"
    "아래 규칙과 예시를 따라 공문 본문만 작성하세요. 마크다운(**·##) 사용 금지.\n"
    + _FEWSHOT_RULES +
    "\n추가 규칙:\n"
    "- 제출 공문은 가나다 항목 없이 2번 문장 바로 다음에 붙임 줄만 작성\n"
    "- 2번 문장: 핵심명사(제목에서 '제출' 제거) + 를/을 붙임과 같이 제출하고자 합니다.\n"
    "\n[예시1 — 붙임 1개]\n"
    "입력:\n"
    "제목: 2026. 교원 연수 결과 제출\n"
    "관련: 전남교육청-11111(2026. 2. 10.)\n"
    "\n출력:\n"
    "1. 관련: 전남교육청-11111(2026. 2. 10.)\n"
    "2. 교원 연수 결과를 붙임과 같이 제출하고자 합니다.\n"
    "\n붙임  교원 연수 결과 보고서 1부.  끝."
)

FEWSHOT_ANNOUNCE_SYSTEM = (
    "당신은 한국 학교 행정 공문 본문 작성 전문가입니다.\n"
    "아래 규칙과 예시를 따라 공문 본문만 작성하세요. 마크다운(**·##) 사용 금지.\n"
    + _FEWSHOT_RULES +
    "\n추가 규칙 (안내 공문):\n"
    "- 2번 문장: 핵심 사업명 + 을/를 다음과 같이 안내하니, 관련 담당자는 협조해 주시기 바랍니다.\n"
    "- 표가 필요한 항목은 반드시 HTML <table> 태그로 작성. 표 외 나머지는 일반 텍스트.\n"
    "- 표 스타일: <table border=\"1\" style=\"border-collapse:collapse;width:100%\">\n"
    "- 헤더 행: <tr><th>...</th></tr>, 데이터 행: <tr><td>...</td></tr>\n"
    "\n[예시]\n"
    "입력:\n"
    "제목: 2026. 중등교과교육연구회 선정·지원 및 자료 제출 안내\n"
    "관련: 중등교육과-7106(2026. 2. 24.)「2026. 중등교과교육연구회 지원 계획 및 신청 안내」\n"
    "\n출력:\n"
    "1. 관련: 중등교육과-7106(2026. 2. 24.)「2026. 중등교과교육연구회 지원 계획 및 신청 안내」\n"
    "2. 2026년 중등교과교육연구회 선정·지원 내역을 다음과 같이 안내하니, 해당 연구회 대표자는 관련 자료를 기한 내 제출해 주시기 바랍니다.\n"
    "  가. 선정 및 지원 내역\n"
    "    1) 선정·지원팀: \n"
    "    2) 선정 기준: 수업공개 실적, 전년도 운영 이력, 예산 집행 적정성 등\n"
    "    3) 세부 내역: [붙임1] 참조\n"
    "  나. 자료 제출:\n"
    '<table border="1" style="border-collapse:collapse;width:100%">\n'
    "<tr><th>제출</th><th>방법</th><th>서식</th><th>기한</th></tr>\n"
    "<tr><td>운영계획서(수정)</td><td>업무포털 게시</td><td>[붙임4]-서식1</td><td>2026.4.17.(금) 17:00</td></tr>\n"
    "<tr><td>상반기 수업공개 계획</td><td>공문 제출</td><td>[붙임3]</td><td>2026.4.17.(금) 17:00</td></tr>\n"
    "<tr><td>하반기 수업공개 계획</td><td>공문 제출</td><td>[붙임3]</td><td>2026.8.21.(금) 17:00</td></tr>\n"
    "<tr><td>결과 보고서</td><td>업무포털 게시</td><td>[붙임4]-서식2</td><td>2026.12.29.(화) 17:00</td></tr>\n"
    "<tr><td>결과 보고서 및 지원금 정산서</td><td>공문 제출</td><td>[붙임4],[붙임5]</td><td></td></tr>\n"
    "</table>\n"
    "  다. 행정사항: 지원금 2026. 4. 15.(수) 이후 지급 예정\n"
    "3. 아울러, 2026. 중등교과교육연구회 운영 협의회를 다음과 같이 실시하니, 대상자는 반드시 참석해 주시기 바랍니다.\n"
    "  가. 일시: 2026. 4. 8.(수) 16:00~17:00\n"
    "  나. 대상: 2026. 중등교과교육연구회 대표(회장 및 총무)\n"
    "  다. 방법: 온라인(zoom)\n"
    "  라. 내용: 2026. 중등교과교육연구회 운영 방향 협의\n"
    "\n붙임  1. 2026. 중등교과교육연구회 선정 및 지원 내역 1부.\n"
    "      2. 2026. 중등교과교육연구회 운영 안내 1부.\n"
    "      3. 2026. 중등교과교육연구회 수업공개 계획(서식) 1부.\n"
    "      4. 2026. 중등교과교육연구회 제출(서식) 1부.\n"
    "      5. 2026. 중등교과교육연구회 지원금 집행 정산서(서식) 1부.  끝."
)

OCR_SYSTEM_PROMPT = (
    "당신은 한국 행정 공문서 이미지 분석 전문가입니다. "
    "이미지에서 텍스트를 정확하게 추출하세요."
)


# ── API 키 / AI 모드 관리 ────────────────────────────────────────

def load_api_key() -> str:
    key = _session_keys.get('gemini', '')
    if key:
        return key
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_api_key(key: str, persist: bool = True):
    _session_keys.pop('gemini', None)
    if persist:
        os.makedirs(APPDATA_DIR, exist_ok=True)
        with open(KEY_FILE, 'w', encoding='utf-8') as f:
            f.write(key.strip())
    else:
        _session_keys['gemini'] = key
    GeminiAdapter.invalidate_client()


def load_ai_mode() -> str:
    if os.path.exists(AI_MODE_FILE):
        with open(AI_MODE_FILE, 'r', encoding='utf-8') as f:
            mode = f.read().strip()
            if mode in _AI_MODES:
                return mode
    return 'none'


def save_ai_mode(mode: str):
    if mode not in _AI_MODES:
        raise ValueError(f'유효하지 않은 모드: {mode}')
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(AI_MODE_FILE, 'w', encoding='utf-8') as f:
        f.write(mode)


def load_ollama_model() -> str:
    if os.path.exists(OLLAMA_MODEL_FILE):
        with open(OLLAMA_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'gemma3:2b'


def save_ollama_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(OLLAMA_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_gemini_model() -> str:
    if os.path.exists(GEMINI_MODEL_FILE):
        with open(GEMINI_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'gemini-3-flash-preview'


def save_gemini_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(GEMINI_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_claude_key() -> str:
    key = _session_keys.get('claude', '')
    if key:
        return key
    if os.path.exists(CLAUDE_KEY_FILE):
        with open(CLAUDE_KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_claude_key(key: str, persist: bool = True):
    _session_keys.pop('claude', None)
    if persist:
        os.makedirs(APPDATA_DIR, exist_ok=True)
        with open(CLAUDE_KEY_FILE, 'w', encoding='utf-8') as f:
            f.write(key.strip())
    else:
        _session_keys['claude'] = key


def load_claude_model() -> str:
    if os.path.exists(CLAUDE_MODEL_FILE):
        with open(CLAUDE_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'claude-sonnet-4-6'


def save_claude_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(CLAUDE_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_openai_key() -> str:
    key = _session_keys.get('openai', '')
    if key:
        return key
    if os.path.exists(OPENAI_KEY_FILE):
        with open(OPENAI_KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_openai_key(key: str, persist: bool = True):
    _session_keys.pop('openai', None)
    if persist:
        os.makedirs(APPDATA_DIR, exist_ok=True)
        with open(OPENAI_KEY_FILE, 'w', encoding='utf-8') as f:
            f.write(key.strip())
    else:
        _session_keys['openai'] = key


def load_openai_model() -> str:
    if os.path.exists(OPENAI_MODEL_FILE):
        with open(OPENAI_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'gpt-4.1'


def save_openai_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(OPENAI_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_nvidia_key() -> str:
    key = _session_keys.get('nvidia', '')
    if key:
        return key
    if os.path.exists(NVIDIA_KEY_FILE):
        with open(NVIDIA_KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_nvidia_key(key: str, persist: bool = True):
    _session_keys.pop('nvidia', None)
    if persist:
        os.makedirs(APPDATA_DIR, exist_ok=True)
        with open(NVIDIA_KEY_FILE, 'w', encoding='utf-8') as f:
            f.write(key.strip())
    else:
        _session_keys['nvidia'] = key


def load_nvidia_model() -> str:
    if os.path.exists(NVIDIA_MODEL_FILE):
        with open(NVIDIA_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'nvidia/llama-3.1-nemotron-ultra-253b-v1'


def save_nvidia_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(NVIDIA_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_opencode_key() -> str:
    key = _session_keys.get('opencode', '')
    if key:
        return key
    if os.path.exists(OPENCODE_KEY_FILE):
        with open(OPENCODE_KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_opencode_key(key: str, persist: bool = True):
    _session_keys.pop('opencode', None)
    if persist:
        os.makedirs(APPDATA_DIR, exist_ok=True)
        with open(OPENCODE_KEY_FILE, 'w', encoding='utf-8') as f:
            f.write(key.strip())
    else:
        _session_keys['opencode'] = key


def load_opencode_model() -> str:
    if os.path.exists(OPENCODE_MODEL_FILE):
        with open(OPENCODE_MODEL_FILE, 'r', encoding='utf-8') as f:
            v = f.read().strip()
            if v:
                return v
    return 'deepseek-v4-flash'


def save_opencode_model(model: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(OPENCODE_MODEL_FILE, 'w', encoding='utf-8') as f:
        f.write(model.strip())


def load_external_model_name() -> str:
    """현재 외부 AI 제공자의 모델 표시명 반환."""
    mode = load_ai_mode()
    if mode == 'claude':
        return load_claude_model()
    elif mode == 'openai':
        return load_openai_model()
    elif mode == 'nvidia':
        return load_nvidia_model()
    elif mode == 'opencode':
        return load_opencode_model()
    return load_gemini_model()


# ── 내부 유틸 ────────────────────────────────────────────────────

def _pixmap_to_base64(pixmap: QPixmap) -> str:
    """QPixmap → PNG base64 문자열."""
    from PyQt5.QtCore import QBuffer, QIODevice
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    pixmap.save(buf, 'PNG')
    buf.close()
    return base64.b64encode(bytes(buf.data())).decode('utf-8')


def _parse_json_response(raw: str) -> dict:
    """단계별 JSON 추출: ①```json블록 ②raw json ③Pydantic fallback ④텍스트 그대로."""
    import json, re

    # Step 1: ```json ... ``` 블록 추출
    m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Step 2: raw 시도
    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        pass

    # Step 3: Pydantic fallback (body 전체를 raw text로)
    try:
        from schema import OfficialDocument
        doc = OfficialDocument(body=raw.strip())
        return doc.model_dump()
    except Exception:
        pass

    # Step 4: 최후의 수단
    return {"body": raw.strip()}


# ── GeminiAdapter ────────────────────────────────────────────────

_gemini_client = None  # 싱글톤 캐시


class GeminiAdapter:
    requires_api_key = True

    @staticmethod
    def get_api_key() -> str:
        return load_api_key()

    @staticmethod
    def _client():
        global _gemini_client
        from google import genai
        if _gemini_client is None:
            _gemini_client = genai.Client(api_key=load_api_key())
        return _gemini_client

    @staticmethod
    def invalidate_client():
        global _gemini_client
        _gemini_client = None

    @staticmethod
    def fetch_models() -> list:
        return GEMINI_DEFAULT_MODELS

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        try:
            from google import genai
            from google.genai import types
            client = cls._client()
            model = load_gemini_model()
            config = types.GenerateContentConfig(
                system_instruction=system,
                thinking_config=types.ThinkingConfig(thinking_budget=0),
                max_output_tokens=1500,
            )
            full = ""
            for chunk in client.models.generate_content_stream(
                model=model,
                contents=prompt,
                config=config,
            ):
                text = chunk.text or ""
                full += text
                if text:
                    on_chunk(text)
            on_done(full)
        except Exception as e:
            on_error(str(e))

    @classmethod
    def ask_with_image(cls, prompt: str, image_b64: str) -> dict:
        try:
            from google import genai
            from google.genai import types
            client = cls._client()
            model = load_gemini_model()
            image_part = types.Part.from_bytes(
                data=base64.b64decode(image_b64),
                mime_type="image/png",
            )
            config = types.GenerateContentConfig(system_instruction=OCR_SYSTEM_PROMPT)
            response = client.models.generate_content(
                model=model,
                contents=[prompt, image_part],
                config=config,
            )
            return _parse_json_response(response.text.strip())
        except Exception as e:
            return {"title": "", "doc_number": "", "error": str(e)}


# ── ClaudeAdapter ─────────────────────────────────────────────────

class ClaudeAdapter:
    requires_api_key = True

    @staticmethod
    def get_api_key() -> str:
        return load_claude_key()

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=cls.get_api_key())
            full = ""
            with client.messages.stream(
                model=load_claude_model(),
                max_tokens=1500,
                system=[{"type": "text", "text": system,
                         "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": prompt}]
            ) as stream:
                for text in stream.text_stream:
                    full += text
                    on_chunk(text)
            on_done(full)
        except Exception as e:
            on_error(str(e))

    @classmethod
    def ask_with_image(cls, prompt: str, image_b64: str) -> dict:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=cls.get_api_key())
            response = client.messages.create(
                model=load_claude_model(),
                max_tokens=256,
                system=[{"type": "text", "text": OCR_SYSTEM_PROMPT,
                         "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": "image/png", "data": image_b64
                    }},
                    {"type": "text", "text": prompt}
                ]}]
            )
            return _parse_json_response(response.content[0].text.strip())
        except Exception as e:
            return {"title": "", "doc_number": "", "error": str(e)}


# ── OpenAIAdapter ─────────────────────────────────────────────────

class OpenAIAdapter:
    requires_api_key = True
    _base_url = None  # None → 기본 OpenAI endpoint

    @classmethod
    def get_api_key(cls) -> str:
        return load_openai_key()

    @classmethod
    def _model(cls) -> str:
        return load_openai_model()

    @classmethod
    def _make_client(cls):
        from openai import OpenAI
        kwargs = {"api_key": cls.get_api_key()}
        if cls._base_url:
            kwargs["base_url"] = cls._base_url
        return OpenAI(**kwargs)

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        try:
            client = cls._make_client()
            full = ""
            with client.chat.completions.create(
                model=cls._model(),
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1500,
                stream=True
            ) as stream:
                for chunk in stream:
                    text = chunk.choices[0].delta.content or ""
                    full += text
                    if text:
                        on_chunk(text)
            on_done(full)
        except Exception as e:
            on_error(str(e))

    @classmethod
    def ask_with_image(cls, prompt: str, image_b64: str) -> dict:
        try:
            client = cls._make_client()
            response = client.chat.completions.create(
                model=cls._model(),
                messages=[{"role": "user", "content": [
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
                    {"type": "text", "text": prompt}
                ]}],
                max_tokens=256
            )
            return _parse_json_response(response.choices[0].message.content.strip())
        except Exception as e:
            return {"title": "", "doc_number": "", "error": str(e)}


# ── NvidiaAdapter (OpenAI 호환 API) ──────────────────────────────

class NvidiaAdapter(OpenAIAdapter):
    _base_url = 'https://integrate.api.nvidia.com/v1'

    @classmethod
    def get_api_key(cls) -> str:
        return load_nvidia_key()

    @classmethod
    def _model(cls) -> str:
        return load_nvidia_model()


# ── OpenCodeAdapter (OpenAI 호환 API) ────────────────────────────

class OpenCodeAdapter(OpenAIAdapter):
    _base_url = 'https://opencode.ai/zen/go/v1'

    @classmethod
    def get_api_key(cls) -> str:
        return load_opencode_key()

    @classmethod
    def _model(cls) -> str:
        return load_opencode_model()


# ── OllamaAdapter ────────────────────────────────────────────────

class OllamaAdapter:
    requires_api_key = False

    @staticmethod
    def get_api_key() -> str:
        return 'ollama'  # 키 불필요 — 항상 통과

    @staticmethod
    def _model():
        return load_ollama_model()

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=INTERNAL_DOC_SYSTEM_PROMPT):
        import json, requests
        url = f'{OLLAMA_HOST}/api/generate'
        model = cls._model()
        payload = {
            'model': model,
            'prompt': f'{system}\n\n{prompt}',
            'stream': True,
            'think': False,
        }
        try:
            with requests.post(url, json=payload, stream=True, timeout=120) as resp:
                if resp.status_code == 404:
                    on_error(f'모델 "{model}"을 찾을 수 없습니다.\nAI 설정에서 올바른 모델명을 입력하세요.\n(ollama list 명령으로 설치된 모델 확인)')
                    return
                resp.raise_for_status()
                full = ''
                for line in resp.iter_lines():
                    if not line:
                        continue
                    data = json.loads(line)
                    if 'error' in data:
                        on_error(data['error'])
                        return
                    text = data.get('response', '')
                    full += text
                    if text:
                        on_chunk(text)
                    if data.get('done'):
                        break
            on_done(full)
        except requests.exceptions.ConnectionError:
            on_error('Ollama 서버에 연결할 수 없습니다. Ollama가 실행 중인지 확인하세요.')
        except Exception as e:
            on_error(str(e))

    @classmethod
    def ask_with_image(cls, prompt: str, image_b64: str) -> dict:
        import json, requests
        url = f'{OLLAMA_HOST}/api/generate'
        payload = {
            'model': cls._model(),
            'prompt': prompt,
            'images': [image_b64],
            'stream': False,
        }
        try:
            resp = requests.post(url, json=payload, timeout=60)
            resp.raise_for_status()
            return _parse_json_response(resp.json().get('response', '').strip())
        except requests.exceptions.ConnectionError:
            return {'title': '', 'doc_number': '', 'error': 'Ollama 서버에 연결할 수 없습니다.'}
        except Exception as e:
            return {'title': '', 'doc_number': '', 'error': str(e)}


# ── 어댑터 선택 ──────────────────────────────────────────────────

def _get_adapter():
    """현재 AI 모드에 맞는 어댑터 클래스 반환."""
    mode = load_ai_mode()
    if mode in ('external', 'gemini'):
        return GeminiAdapter
    elif mode == 'claude':
        return ClaudeAdapter
    elif mode == 'openai':
        return OpenAIAdapter
    elif mode == 'nvidia':
        return NvidiaAdapter
    elif mode == 'opencode':
        return OpenCodeAdapter
    elif mode == 'internal':
        return OllamaAdapter
    return GeminiAdapter


# ── 공개 인터페이스 ──────────────────────────────────────────────

class AiStreamThread(QThread):
    new_text_signal  = pyqtSignal(str)
    finished_signal  = pyqtSignal(str)
    error_signal     = pyqtSignal(str)

    def __init__(self, prompt: str, system: str = None, structured: bool = False):
        super().__init__()
        self.prompt = prompt
        self.system = system
        self.structured = structured
        self._is_running = True

    def stop(self):
        self._is_running = False

    def _retry_prompt(self, error_msg: str) -> str:
        return (
            self.prompt
            + "\n\n[시스템 지시] 이전 답변이 형식에 맞지 않았습니다.\n"
            + f"오류: {error_msg}\n"
            + '반드시 {"body": "여기에 공문 본문 전체를 작성"} JSON 형식으로만 출력하세요.'
        )

    def run(self):
        adapter = _get_adapter()
        if getattr(adapter, 'requires_api_key', True) and not adapter.get_api_key():
            self.error_signal.emit("API 키가 설정되지 않았습니다. 설정 버튼을 눌러 키를 입력하세요.")
            return

        if not self.structured:
            # 일반 스트리밍 모드 (자유 채팅 등)
            def on_chunk(text):
                if self._is_running:
                    self.new_text_signal.emit(text)

            def on_done(full):
                if self._is_running:
                    self.finished_signal.emit(full)

            kwargs = {}
            if self.system is not None:
                kwargs['system'] = self.system
            adapter.stream_text(self.prompt, on_chunk, on_done, self.error_signal.emit, **kwargs)
            return

        # Structured 모드: JSON 수집 → 파싱 → 검증 → 재시도
        for attempt in range(2):
            if not self._is_running:
                return

            buffer = []
            err = [None]

            def on_chunk(text):
                if self._is_running:
                    buffer.append(text)

            def on_done(full):
                pass  # buffer에 이미 다 있음

            def on_error(msg):
                err[0] = msg

            kwargs = {}
            if self.system is not None:
                kwargs['system'] = self.system
            adapter.stream_text(self.prompt, on_chunk, on_done, on_error, **kwargs)

            if err[0]:
                if attempt == 0:
                    self.prompt = self._retry_prompt(err[0])
                    continue
                self.error_signal.emit(err[0])
                return

            full = ''.join(buffer)
            parsed = _parse_json_response(full)
            body = parsed.get('body', '').strip()

            if len(body) > 20:
                # 성공: 본문을 스트리밍처럼 한 번에 전송
                self.new_text_signal.emit(body)
                self.finished_signal.emit(body)
                return

            if attempt == 0:
                reason = '본문이 너무 짧음' if len(body) <= 20 else 'JSON 형식 아님'
                self.prompt = self._retry_prompt(reason)

        # 2회 실패
        self.error_signal.emit("AI 응답을 올바르게 받지 못했습니다. 다시 시도해 주세요.")


class AiImageWorker(QThread):
    """이미지 → 공문 제목/번호 추출 스레드."""
    finished = pyqtSignal(dict)
    error    = pyqtSignal(str)

    _PROMPT = (
        "이 이미지는 한국 행정 공문서의 일부입니다.\n"
        "이미지에서 공문 제목과 공문번호를 추출하세요.\n"
        "반드시 아래 JSON 형식으로만 답하세요. 설명 없이 JSON만 출력하세요.\n"
        '{"title": "공문 제목", "doc_number": "공문번호(없으면 빈 문자열)"}'
    )

    def __init__(self, pixmap: QPixmap):
        super().__init__()
        self._image_b64 = _pixmap_to_base64(pixmap)

    def run(self):
        adapter = _get_adapter()
        if getattr(adapter, 'requires_api_key', True) and not adapter.get_api_key():
            self.error.emit("API 키가 설정되지 않았습니다.")
            return
        result = adapter.ask_with_image(self._PROMPT, self._image_b64)
        if "error" in result:
            self.error.emit(result["error"])
        else:
            self.finished.emit(result)
