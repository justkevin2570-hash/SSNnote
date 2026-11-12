"""
ai_client.py — AI 제공자 추상화 레이어

지원 제공자: OpenCode(Go·Zen), Claude(Anthropic), ChatGPT(OpenAI) — API 키로 연결. Claude 구독은 내 PC의 Claude Code로 연결.
  · OpenCode  : 모델마다 OpenAI(chat-completions) 또는 Claude(messages) 방식
  · Claude    : Anthropic Messages API
  · ChatGPT   : OpenAI Responses API
다른 제공자로 교체 시 이 파일의 어댑터만 수정하면 됩니다.
"""

import os
import sys
import json
import base64
import atexit
import shutil
import subprocess
import tempfile
import threading
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QPixmap

import auth

APPDATA_DIR        = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
AI_MODE_FILE       = os.path.join(APPDATA_DIR, 'ai_mode.txt')
CLAUDE_KEY_FILE    = os.path.join(APPDATA_DIR, 'claude_key.txt')
CLAUDE_MODEL_FILE  = os.path.join(APPDATA_DIR, 'claude_model.txt')
OPENAI_KEY_FILE    = os.path.join(APPDATA_DIR, 'openai_key.txt')
OPENAI_MODEL_FILE  = os.path.join(APPDATA_DIR, 'openai_model.txt')
OPENCODE_KEY_FILE   = os.path.join(APPDATA_DIR, 'opencode_key.txt')
OPENCODE_MODEL_FILE = os.path.join(APPDATA_DIR, 'opencode_model.txt')
OPENCODE_PLAN_FILE  = os.path.join(APPDATA_DIR, 'opencode_plan.txt')
OPENCODE_VISION_FILE = os.path.join(APPDATA_DIR, 'opencode_vision_models.json')
OPENCODE_MODELS_FILE = os.path.join(APPDATA_DIR, 'opencode_models.json')   # {id: {image, protocol}}
CLAUDE_CODE_PATH_FILE  = os.path.join(APPDATA_DIR, 'claude_code_path.txt')  # 사용자가 지정한 claude 경로 (비우면 자동)
CLAUDE_CODE_MODEL_FILE = os.path.join(APPDATA_DIR, 'claude_code_model.txt')

# 접속 주소 (None = 각 SDK 기본 주소). 시험에서만 바꾼다.
CLAUDE_BASE_URL = None
OPENAI_BASE_URL = None

# ── 자격증명 암호화 헬퍼 (Windows DPAPI, 미지원 시 평문 폴백) ──
def _read_secret(path: str) -> str:
    try:
        with open(path, encoding='utf-8') as f:
            raw = f.read().strip()
        if not raw:
            return ''
        return auth._unprotect(raw)
    except Exception:
        return ''


def _write_secret(path: str, val: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(auth._protect(val))
    except Exception:
        pass

# 세션 키 (API 키 저장 체크 해제 시 메모리에만 보관)
_session_keys: dict[str, str] = {}

# AI 모드
_AI_MODES = ('none', 'claude', 'openai', 'opencode', 'claude_code')
PROVIDER_LABELS = {'opencode': 'OpenCode', 'claude': 'Claude', 'openai': 'ChatGPT',
                   'claude_code': 'Claude 구독'}
# 예전 버전에서 쓰던(지금은 지원하지 않는) 모드 → 안내용 이름
_LEGACY_MODES = {'external': 'Gemini', 'gemini': 'Gemini', 'nvidia': 'NVIDIA', 'internal': 'Ollama(내부 AI)',
                 'exaone': 'EXAONE'}

# Claude — 서버에서 목록을 못 받을 때 쓰는 예비 목록 (2026-11)
CLAUDE_DEFAULT_MODEL = 'claude-haiku-5-5'
CLAUDE_DEFAULT_MODELS = [
    "claude-haiku-5-5",
    "claude-sonnet-5-5",
    "claude-opus-5-5",
]

# ChatGPT — 서버에서 목록을 못 받을 때 쓰는 예비 목록 (2026-11)
OPENAI_DEFAULT_MODEL = 'gpt-5.4-mini'
OPENAI_DEFAULT_MODELS = [
    "gpt-5.4-mini",
    "gpt-5.4-nano",
    "gpt-5.5",
    "gpt-6-luna",
    "gpt-6-sol",
]
# 대화(글쓰기)용이 아닌 OpenAI 모델 이름에 들어가는 말
_OPENAI_SKIP = ('audio', 'realtime', 'tts', 'transcribe', 'image', 'search', 'codex', 'embedding',
                'moderation', 'dall-e', 'whisper', 'instruct', 'computer-use')

# OpenCode — 요금제별 접속 주소 (https://opencode.ai/docs/go/ , /docs/zen/)
#   요청에는 x-opencode-session 헤더가 꼭 있어야 한다 (없으면 400 MissingSessionID).
OPENCODE_PLANS = {
    'go':  {'label': 'OpenCode Go (월 구독)',
            'base_url': 'https://opencode.ai/inference/go/openai/v1',          # OpenAI(chat-completions) 방식
            'anthropic_base_url': 'https://opencode.ai/inference/go/anthropic',  # Claude(messages) 방식
            'models_url': 'https://opencode.ai/inference/go/v1/models'},
    'zen': {'label': 'OpenCode Zen (충전식)',
            'base_url': 'https://opencode.ai/inference/openai/v1',
            'anthropic_base_url': 'https://opencode.ai/inference/anthropic',
            'models_url': 'https://opencode.ai/inference/v1/models'},
}
# 저렴하고(100만 토큰당 $0.10/$0.50) 학습에 쓰이지 않으며 이미지도 읽는 모델 (OpenCode Go 문서, 2026-11)
OPENCODE_DEFAULT_MODEL = 'claude-haiku-5-5'
# 서버에서 모델 목록을 못 받을 때 쓰는 예비 목록 (Go, 2026-11 기준)
OPENCODE_DEFAULT_MODELS = [
    "claude-haiku-5-5",
    "kimi-k3",
    "deepseek-v4.1-flash",
    "kimi-k3.1-pro",
    "glm-5.3-flash",
    "glm-5.3",
    "deepseek-v4-pro",
    "mimo-v2.6-pro",
    "mimo-v2.6-flash",
    "minimax-m3.1-flash",
]

# 외부 AI(대형 모델): 형식 제약 최소화, 모델 자율성 최대
EXTERNAL_DOC_SYSTEM_PROMPT = (
    "당신은 한국 학교 행정 공문 작성 전문가입니다. "
    "공식 공문 문체(개조식, '~함', '~바람', '~고자 함')로 작성하세요. "
    "불필요한 설명 없이 공문 본문만 작성하세요. "
    "마크다운(**·##) 사용 금지."
)


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




# ── AI 모드 ──────────────────────────────────────────────────────

def _read_mode_file() -> str:
    try:
        with open(AI_MODE_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except OSError:
        return ''


def load_ai_mode() -> str:
    """'opencode' | 'claude' | 'openai' | 'claude_code' | 'none'. 지원하지 않는 예전 모드(Gemini·NVIDIA·Ollama·EXAONE)는 'none'."""
    mode = _read_mode_file()
    return mode if mode in _AI_MODES else 'none'


def legacy_ai_name() -> str:
    """예전 버전에서 고른 AI가 이제 지원되지 않으면 그 이름 (안내용), 아니면 ''."""
    return _LEGACY_MODES.get(_read_mode_file(), '')


def save_ai_mode(mode: str):
    if mode not in _AI_MODES:
        raise ValueError(f'유효하지 않은 모드: {mode}')
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(AI_MODE_FILE, 'w', encoding='utf-8') as f:
        f.write(mode)


def _load_text(path: str, default: str) -> str:
    try:
        with open(path, 'r', encoding='utf-8') as f:
            v = f.read().strip()
        return v or default
    except OSError:
        return default


def _save_text(path: str, value: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(value.strip())


def _load_key(name: str, path: str) -> str:
    return _session_keys.get(name, '') or _read_secret(path)


def _save_key(name: str, path: str, key: str, persist: bool):
    _session_keys.pop(name, None)
    if persist:
        _write_secret(path, key.strip())
    else:
        _session_keys[name] = key.strip()
        try:                      # 저장하지 않기로 했으면 예전에 저장한 키도 지운다
            os.remove(path)
        except OSError:
            pass


# ── Claude 키·모델 ───────────────────────────────────────────────

def load_claude_key() -> str:
    return _load_key('claude', CLAUDE_KEY_FILE)


def save_claude_key(key: str, persist: bool = True):
    _save_key('claude', CLAUDE_KEY_FILE, key, persist)


def load_claude_model() -> str:
    return _load_text(CLAUDE_MODEL_FILE, CLAUDE_DEFAULT_MODEL)


def save_claude_model(model: str):
    _save_text(CLAUDE_MODEL_FILE, model)


# ── Claude 구독(Claude Code) 설정 ─────────────────────────────────
# 키가 없다: 로그인은 Claude Code가 자기 자리에 보관한다. 여기서는 실행 파일 경로와 모델 이름만 저장한다.
CLAUDE_CODE_MODELS = (('', '기본 모델'), ('sonnet', 'Sonnet'), ('opus', 'Opus'), ('haiku', 'Haiku'))


def load_claude_code_path() -> str:
    return _load_text(CLAUDE_CODE_PATH_FILE, '')


def save_claude_code_path(path: str):
    _save_text(CLAUDE_CODE_PATH_FILE, path)


def load_claude_code_model() -> str:
    """CLI 별칭(sonnet·opus·haiku) 또는 '' (CLI 기본 모델)."""
    v = _load_text(CLAUDE_CODE_MODEL_FILE, '')
    return v if v in ('sonnet', 'opus', 'haiku') else ''


def save_claude_code_model(model: str):
    _save_text(CLAUDE_CODE_MODEL_FILE, model if model in ('sonnet', 'opus', 'haiku') else '')


# ── ChatGPT 키·모델 ──────────────────────────────────────────────

def load_openai_key() -> str:
    return _load_key('openai', OPENAI_KEY_FILE)


def save_openai_key(key: str, persist: bool = True):
    _save_key('openai', OPENAI_KEY_FILE, key, persist)


def load_openai_model() -> str:
    return _load_text(OPENAI_MODEL_FILE, OPENAI_DEFAULT_MODEL)


def save_openai_model(model: str):
    _save_text(OPENAI_MODEL_FILE, model)


# ── OpenCode 키·모델·요금제 ──────────────────────────────────────

def load_opencode_key() -> str:
    return _load_key('opencode', OPENCODE_KEY_FILE)


def save_opencode_key(key: str, persist: bool = True):
    _save_key('opencode', OPENCODE_KEY_FILE, key, persist)


def load_opencode_model() -> str:
    return _load_text(OPENCODE_MODEL_FILE, OPENCODE_DEFAULT_MODEL)


def save_opencode_model(model: str):
    _save_text(OPENCODE_MODEL_FILE, model)


def load_opencode_plan() -> str:
    v = _load_text(OPENCODE_PLAN_FILE, 'go')
    return v if v in OPENCODE_PLANS else 'go'


def save_opencode_plan(plan: str):
    _save_text(OPENCODE_PLAN_FILE, plan if plan in OPENCODE_PLANS else 'go')


def save_opencode_vision_models(model_ids):
    """이미지를 읽을 수 있는 모델 목록 (마지막으로 서버에서 받은 것)."""
    import json
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(OPENCODE_VISION_FILE, 'w', encoding='utf-8') as f:
        json.dump(sorted(model_ids), f)


def save_opencode_model_info(models):
    """모델별 이미지 지원·요청 방식 저장 (마지막으로 서버에서 받은 목록)."""
    import json
    os.makedirs(APPDATA_DIR, exist_ok=True)
    info = {m['id']: {'image': bool(m.get('image')), 'protocol': m.get('protocol', 'chat')} for m in models}
    with open(OPENCODE_MODELS_FILE, 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False)
    save_opencode_vision_models([m['id'] for m in models if m.get('image')])


def _opencode_model_info(model: str) -> dict:
    import json
    try:
        with open(OPENCODE_MODELS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f).get(model) or {}
    except (OSError, ValueError):
        return {}


def opencode_model_protocol(model: str) -> str:
    """'messages'(Claude 방식) 또는 'chat'(OpenAI 방식). 모르면 이름으로 짐작 (claude-* → messages)."""
    p = _opencode_model_info(model).get('protocol')
    if p in ('chat', 'messages'):
        return p
    return 'messages' if model.startswith('claude-') else 'chat'


def opencode_model_supports_image(model: str):
    """True / False / None(모름 — 목록을 받은 적 없음)."""
    import json
    info = _opencode_model_info(model)
    if 'image' in info:
        return info['image']
    try:
        with open(OPENCODE_VISION_FILE, 'r', encoding='utf-8') as f:
            return model in json.load(f)
    except (OSError, ValueError):
        return None


_OPENCODE_SESSION = None


def _app_version() -> str:
    try:
        from updater import _read_app_version
        return _read_app_version().lstrip('v') or '0'
    except Exception:
        return '0'


def opencode_headers() -> dict:
    """OpenCode가 요구하는 헤더: 앱 이름(User-Agent)과 대화 세션 ID (앱 실행마다 하나)."""
    global _OPENCODE_SESSION
    if _OPENCODE_SESSION is None:
        import uuid
        _OPENCODE_SESSION = 'ssnnote-' + uuid.uuid4().hex
    return {'x-opencode-session': _OPENCODE_SESSION, 'User-Agent': f'ssnnote/{_app_version()}'}


def fetch_opencode_models(plan: str) -> list:
    """요금제에서 쓸 수 있는 대화 모델 [{'id', 'name', 'image', 'protocol'}].

    OpenAI(chat-completions) 방식 또는 Claude(messages) 방식을 기본 지원하는 모델만 고른다.
    두 방식 모두 되면 OpenAI 방식. 실패하면 예비 목록 (image=None: 모름).
    """
    import requests
    try:
        r = requests.get(OPENCODE_PLANS[plan]['models_url'], headers=opencode_headers(), timeout=15)
        r.raise_for_status()
        out = []
        for m in r.json().get('data', []):
            api = m.get('api') or {}
            if m.get('type', 'chat') != 'chat' or not ('chat-completions' in api or 'messages' in api):
                continue
            price = (((m.get('cost') or {}).get('text') or {}).get('input') or {}).get('price')
            if price == 0:      # 무료 모델은 OpenCode 앱 안에서만 쓸 수 있어 다른 앱에서는 항상 거절됨
                continue
            inputs = (m.get('modalities') or {}).get('input', [])
            out.append({'id': m['id'], 'name': m.get('name', m['id']),
                        'image': any(str(x).startswith('image/') for x in inputs),
                        'protocol': 'chat' if 'chat-completions' in api else 'messages'})
        if out:
            return out
    except Exception:
        pass
    return [{'id': mid, 'name': mid, 'image': None, 'protocol': 'messages' if mid.startswith('claude-') else 'chat'}
            for mid in OPENCODE_DEFAULT_MODELS]


# ── 모델 목록 받기 (Claude·ChatGPT) ──────────────────────────────

def fetch_claude_models(key: str) -> list:
    """키로 쓸 수 있는 Claude 모델 id 목록 (최신순). 실패하면 예비 목록."""
    try:
        client = ClaudeAdapter._client(key, max_retries=0, timeout=15)
        ids = [m.id for m in client.models.list(limit=100) if str(m.id).startswith('claude-')]
        if ids:
            return ids
    except Exception:
        pass
    return list(CLAUDE_DEFAULT_MODELS)


def _openai_chat_model(mid: str) -> bool:
    m = mid.lower()
    if not (m.startswith('gpt-') or m.startswith('chatgpt-') or (m[:1] == 'o' and m[1:2].isdigit())):
        return False
    return not any(w in m for w in _OPENAI_SKIP)


def fetch_openai_models(key: str) -> list:
    """키로 쓸 수 있는 ChatGPT 글쓰기 모델 id 목록 (최신순). 실패하면 예비 목록."""
    try:
        client = OpenAIAdapter._client(key, max_retries=0, timeout=15)
        models = [m for m in client.models.list() if _openai_chat_model(str(m.id))]
        models.sort(key=lambda m: getattr(m, 'created', 0) or 0, reverse=True)
        ids = [m.id for m in models]
        if ids:
            return ids
    except Exception:
        pass
    return list(OPENAI_DEFAULT_MODELS)


# ── 연결 확인 ────────────────────────────────────────────────────

def friendly_error(e: Exception, model: str = '', provider: str = '') -> str:
    """API 오류를 선생님이 알아볼 수 있는 한국어 안내로 바꾼다."""
    code = getattr(e, 'status_code', None)
    text = str(e)
    if code == 401:
        if provider == 'opencode':
            return 'API 키가 맞지 않습니다.\n키를 다시 확인하고, 요금제(Go/Zen)를 바르게 골랐는지도 확인해 주세요.'
        return 'API 키가 맞지 않습니다. 키를 다시 확인해 주세요.'
    if code in (402, 403):
        if provider == 'opencode':
            return '이 요금제에서 쓸 수 없거나 잔액·사용 한도가 부족합니다.\nOpenCode 콘솔에서 요금제와 잔액을 확인해 주세요.'
        return '이 키로는 이 모델을 쓸 수 없거나 결제(크레딧)가 필요합니다.\n제공사 콘솔에서 결제 정보와 권한을 확인해 주세요.'
    if code == 429 and 'FreeUsageLimit' in text:
        return '무료 모델은 OpenCode 앱 안에서만 쓸 수 있습니다. 유료 모델을 골라 주세요.'
    if code == 429 and ('insufficient_quota' in text or 'credit balance' in text):
        return '크레딧(잔액)이 부족합니다. 제공사 콘솔에서 결제 정보를 확인해 주세요.'
    if code == 429:
        return '사용 한도에 걸렸습니다. 잠시 후 다시 시도하거나 다른 모델을 고르세요.'
    if code == 404:
        return f'모델 "{model}"을(를) 찾을 수 없습니다. 다른 모델을 골라 주세요.' if model else '요청한 주소를 찾을 수 없습니다.'
    if code is not None and code >= 500:
        return 'AI 서버에 일시적인 문제가 있습니다. 잠시 후 다시 시도해 주세요.'
    if code is None and ('connect' in type(e).__name__.lower() or 'timeout' in type(e).__name__.lower()):
        return '인터넷 연결을 확인해 주세요.'
    return f'연결 실패: {e}'


_PING = "OK라고만 답하세요."


def test_opencode(plan: str, key: str, model: str):
    """아주 짧은 요청으로 연결을 확인한다. (ok, 한국어 메시지)."""
    try:
        if opencode_model_protocol(model) == 'messages':
            client = OpenCodeAdapter._anthropic_client(key, plan, max_retries=0, timeout=30)
            client.messages.create(model=model, max_tokens=256, messages=[{"role": "user", "content": _PING}])
        else:
            client = OpenCodeAdapter._client(key, plan, max_retries=0, timeout=30)
            client.chat.completions.create(model=model, max_tokens=16, messages=[{"role": "user", "content": _PING}])
        return True, '연결 성공'
    except Exception as e:
        return False, friendly_error(e, model, 'opencode')


def test_claude(key: str, model: str):
    try:
        ClaudeAdapter._client(key, max_retries=0, timeout=30).messages.create(
            model=model, max_tokens=256, messages=[{"role": "user", "content": _PING}])
        return True, '연결 성공'
    except Exception as e:
        return False, friendly_error(e, model, 'claude')


def test_openai(key: str, model: str):
    try:
        OpenAIAdapter._client(key, max_retries=0, timeout=30).responses.create(
            model=model, input=_PING, max_output_tokens=256)
        return True, '연결 성공'
    except Exception as e:
        return False, friendly_error(e, model, 'openai')


def test_claude_code(model: str = None):
    """짧은 질문 한 번으로 Claude 구독 연결을 확인한다. (ok, 한국어 메시지 — 성공이면 걸린 시간 포함)."""
    import time
    t0 = time.monotonic()
    try:
        _claude_ask('간단히 답하세요.', '안녕이라고만 답해', model=model, timeout=CLAUDE_CODE_TIMEOUT)
        return True, f'연결 성공 ({time.monotonic() - t0:.1f}초)'
    except ClaudeCodeError as e:
        return False, str(e)
    except Exception as e:
        return False, f'연결 실패: {e}'


def load_external_model_name() -> str:
    """현재 AI 제공자의 모델 이름 ('' = AI 꺼짐)."""
    return {'claude': load_claude_model, 'openai': load_openai_model,
            'opencode': load_opencode_model,
            'claude_code': lambda: load_claude_code_model() or '기본 모델'}.get(load_ai_mode(), lambda: '')()


def ai_status_text() -> str:
    """'OpenCode · claude-haiku-5-5' 같은 표시용 글. AI 꺼짐이면 ''."""
    mode = load_ai_mode()
    if mode == 'none':
        return ''
    return f'{PROVIDER_LABELS[mode]} · {load_external_model_name()}'


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

    m = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass
    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        pass
    try:
        from schema import OfficialDocument
        doc = OfficialDocument(body=raw.strip())
        return doc.model_dump()
    except Exception:
        pass
    return {"body": raw.strip()}


_REFUSAL = 'AI가 이 요청에는 답하지 않았습니다. 내용을 조금 바꿔 다시 시도하거나 다른 모델을 골라 주세요.'
_OPENCODE_REFUSAL = _REFUSAL     # 예전 이름


# ── 요청 방식별 공통 코드 ────────────────────────────────────────

def _messages_content(prompt: str, image_b64: str = None):
    content = []
    if image_b64:
        content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": image_b64}})
    content.append({"type": "text", "text": prompt})
    return content


def _messages_stream(client, model, system, prompt, on_chunk, on_done, on_error, max_tokens=4096):
    """Claude(messages) 방식 스트리밍. 글자(text)만 보내고 생각(thinking) 블록은 뺀다."""
    full = ""
    with client.messages.stream(model=model, max_tokens=max_tokens, system=system,
                                messages=[{"role": "user", "content": prompt}]) as stream:
        for text in stream.text_stream:
            full += text
            on_chunk(text)
        final = stream.get_final_message()
    if getattr(final, 'stop_reason', '') == 'refusal':
        on_error(_REFUSAL)
        return
    on_done(full)


def _messages_complete(client, model, system, prompt, image_b64=None, max_tokens=4000) -> str:
    r = client.messages.create(model=model, max_tokens=max_tokens, system=system,
                               messages=[{"role": "user", "content": _messages_content(prompt, image_b64)}])
    if getattr(r, 'stop_reason', '') == 'refusal':
        raise RuntimeError(_REFUSAL)
    return ''.join(b.text for b in r.content if getattr(b, 'type', '') == 'text').strip()


def _chat_stream(client, model, system, prompt, on_chunk, on_done, max_tokens=4096):
    """OpenAI chat-completions 방식 스트리밍 (OpenCode의 대부분 모델)."""
    full = ""
    with client.chat.completions.create(
            model=model, stream=True,
            max_tokens=max_tokens,          # 먼저 '생각'하는 모델도 답을 끝까지 쓰도록 (쓴 만큼만 과금)
            messages=[{"role": "system", "content": system}, {"role": "user", "content": prompt}]) as stream:
        for chunk in stream:
            if not chunk.choices:      # 일부 호환 서버(OpenCode 등)는 마지막에 내용 없는 조각을 보냄
                continue
            text = chunk.choices[0].delta.content or ""
            full += text
            if text:
                on_chunk(text)
    on_done(full)


def _chat_complete(client, model, system, prompt, image_b64=None, max_tokens=4000) -> str:
    if image_b64:
        user = [{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
                {"type": "text", "text": prompt}]
    else:
        user = prompt
    r = client.chat.completions.create(
        model=model, max_tokens=max_tokens,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}])
    return (r.choices[0].message.content or '').strip()


# ── ClaudeAdapter (Anthropic Messages API) ───────────────────────

class ClaudeAdapter:
    requires_api_key = True

    @staticmethod
    def get_api_key() -> str:
        return load_claude_key()

    @staticmethod
    def _client(key: str = None, **kw):
        import anthropic
        if CLAUDE_BASE_URL:
            kw['base_url'] = CLAUDE_BASE_URL
        return anthropic.Anthropic(api_key=key if key is not None else load_claude_key(), **kw)

    @staticmethod
    def _system(system: str):
        return [{"type": "text", "text": system, "cache_control": {"type": "ephemeral"}}]

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        try:
            _messages_stream(cls._client(), load_claude_model(), cls._system(system), prompt,
                             on_chunk, on_done, on_error)
        except Exception as e:
            on_error(friendly_error(e, load_claude_model(), 'claude'))

    @classmethod
    def complete(cls, system: str, prompt: str, image_b64: str = None, max_tokens: int = 4000) -> str:
        """한 번에 응답 전체를 받는다 (이미지 선택). 실패 시 예외."""
        return _messages_complete(cls._client(), load_claude_model(), cls._system(system), prompt,
                                  image_b64=image_b64, max_tokens=max_tokens)


# ── Claude 구독 (이 PC의 Claude Code CLI를 자식 프로세스로 실행) ──────
# 앱은 로그인 파일·토큰을 열거나 저장하지 않는다. --bare는 쓰지 않는다(OAuth 로그인을 읽지 않음).
CLAUDE_CODE_TIMEOUT = 120            # 한 번 묻는 제한 시간(초)
CLAUDE_CODE_LOGIN_TIMEOUT = 600      # 브라우저에서 승인할 시간(초)
CLAUDE_CODE_INSTALL_PAGE = 'https://code.claude.com/docs'
_SYSTEM_ARG_MAX = 16000              # 명령줄 길이 제한(Windows 32K) 안에 두려고, 이보다 길면 질문 앞에 붙여 보냄
_CLAUDE_DROP_ENV = ('ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_BASE_URL', 'CLAUDE_CODE_OAUTH_TOKEN')
_CLAUDE_MSG_NOT_INSTALLED = 'Claude Code가 설치돼 있지 않아요. AI 설정 → Claude 구독 → [설치]를 눌러 주세요.'
_CLAUDE_MSG_LOGIN = 'Claude에 로그인돼 있지 않아요. [Claude로 로그인]을 눌러 주세요.'
_CLAUDE_MSG_LIMIT = '구독 사용 한도에 도달했어요. 잠시 후 다시 시도하세요.'
_CLAUDE_MSG_TIMEOUT = 'Claude 응답이 너무 오래 걸려요. 다시 시도해 주세요.'
_CLAUDE_MSG_CANCEL = '취소했어요.'
_claude_procs: dict = {}             # 실행 중인 claude 프로세스 → 취소 여부
_claude_lock = threading.Lock()


class ClaudeCodeError(RuntimeError):
    """사용자에게 그대로 보여 줄 한국어 오류."""


def find_claude_cli():
    """claude 실행 파일 경로 (없으면 None). 지정 경로 → PATH → 흔한 설치 위치 순."""
    custom = load_claude_code_path()
    if custom and os.path.isfile(custom):
        return custom
    found = shutil.which('claude')
    if found:
        return found
    home = os.path.expanduser('~')
    local = os.environ.get('LOCALAPPDATA', os.path.join(home, 'AppData', 'Local'))
    appdata = os.environ.get('APPDATA', os.path.join(home, 'AppData', 'Roaming'))
    for c in (os.path.join(home, '.local', 'bin', 'claude.exe'),
              os.path.join(appdata, 'npm', 'claude.cmd'),
              os.path.join(local, 'Programs', 'claude', 'claude.exe')):
        if os.path.isfile(c):
            return c
    return None


def cancel_claude_code():
    """실행 중인 Claude Code 프로세스를 끊는다 (취소 버튼·앱 종료)."""
    with _claude_lock:
        for p in list(_claude_procs):
            _claude_procs[p] = True
            try:
                p.terminate()
            except OSError:
                pass


atexit.register(cancel_claude_code)


def _run_claude(args, stdin_text=None, timeout=CLAUDE_CODE_TIMEOUT) -> subprocess.CompletedProcess:
    """claude를 앱 전용 빈 폴더에서 실행한다. 구독 로그인을 쓰도록 API 키 환경변수는 뺀다.
    실행 파일이 없거나 시간이 넘으면 ClaudeCodeError, 취소하면 취소 오류."""
    exe = find_claude_cli()
    if not exe:
        raise ClaudeCodeError(_CLAUDE_MSG_NOT_INSTALLED)
    env = {k: v for k, v in os.environ.items() if k not in _CLAUDE_DROP_ENV}
    kw = {'creationflags': subprocess.CREATE_NO_WINDOW} if sys.platform == 'win32' else {}
    workdir = tempfile.mkdtemp(prefix='ssnnote_claude_')     # 메모 DB 폴더·사용자 프로젝트 설정을 읽지 않도록
    try:
        try:
            proc = subprocess.Popen([exe, *args], cwd=workdir, env=env, stdin=subprocess.PIPE,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    encoding='utf-8', errors='replace', **kw)
        except FileNotFoundError:
            raise ClaudeCodeError(_CLAUDE_MSG_NOT_INSTALLED)
        with _claude_lock:
            _claude_procs[proc] = False
        try:
            out, err = proc.communicate(stdin_text or '', timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.communicate()
            raise ClaudeCodeError(_CLAUDE_MSG_TIMEOUT)
        finally:
            with _claude_lock:
                cancelled = _claude_procs.pop(proc, False)
        if cancelled:
            raise ClaudeCodeError(_CLAUDE_MSG_CANCEL)
        return subprocess.CompletedProcess(args, proc.returncode, out or '', err or '')
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def _claude_friendly(text: str) -> str:
    """Claude Code 출력에서 알아볼 수 있는 오류(로그인·한도)를 찾는다. 모르면 ''."""
    low = (text or '').lower()
    if any(w in low for w in ('usage limit', 'rate limit', 'rate_limit', '429', 'limit reached')):
        return _CLAUDE_MSG_LIMIT
    if any(w in low for w in ('/login', 'log in', 'login', 'not logged', 'authenticate', '401', 'unauthorized')):
        return _CLAUDE_MSG_LOGIN
    return ''


def _claude_result_text(p: subprocess.CompletedProcess) -> str:
    """--output-format json 결과에서 답 글(result)만 꺼낸다. 실패면 ClaudeCodeError."""
    raw = (p.stdout or '').strip()
    try:
        data = json.loads(raw) if raw else None
    except ValueError:
        data = None
    if not isinstance(data, dict) or p.returncode != 0 or data.get('is_error') is True:
        detail = str((data or {}).get('result') or '') or (p.stderr or '').strip() or raw
        raise ClaudeCodeError(_claude_friendly(detail) or f'Claude 연결 실패: {detail[-300:]}')
    return str(data.get('result') or '').strip()


def _claude_ask(system: str, prompt: str, model: str = None, timeout: int = CLAUDE_CODE_TIMEOUT) -> str:
    """질문 한 번 → 답 글. 도구는 모두 끄고, 질문은 표준입력으로 보낸다."""
    args = ['-p', '--output-format', 'json', '--no-session-persistence', '--tools', '']
    model = load_claude_code_model() if model is None else model
    if model:
        args += ['--model', model]
    stdin_text = prompt
    if len(system) > _SYSTEM_ARG_MAX:
        stdin_text = system + '\n\n' + prompt
    else:
        args += ['--system-prompt', system]
    return _claude_result_text(_run_claude(args, stdin_text, timeout=timeout))


class ClaudeCodeAdapter:
    """이 PC에 설치된 Claude Code로 Claude 구독(Pro/Max)을 쓴다. 본인 전용, 사용량은 구독 한도에서 차감."""
    requires_api_key = False     # 키가 없다: 로그인 여부는 호출할 때 Claude Code가 확인한다
    supports_image = False       # 이미지는 보내지 않고 PC에서 읽은 글자로 대신 보낸다

    @staticmethod
    def get_api_key() -> str:
        return 'subscription'    # 화면 표시용 표식일 뿐, 키가 아니다

    @classmethod
    def complete(cls, system: str, prompt: str, image_b64: str = None, max_tokens: int = 4000) -> str:
        """한 번에 답 전체를 받는다. 글 한도(max_tokens)는 CLI에 대응 옵션이 없어 무시한다."""
        return _claude_ask(system, prompt)

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        """1차 구현: 답을 다 받은 뒤 한 번에 보낸다 (동작 우선)."""
        try:
            full = cls.complete(system, prompt)
        except ClaudeCodeError as e:
            on_error(str(e))
            return
        except Exception as e:
            on_error(f'Claude 구독 연결에 실패했어요: {e}')
            return
        if full:
            on_chunk(full)
        on_done(full)


def claude_code_status() -> dict:
    """{'installed', 'logged_in', 'email', 'plan', 'path'}. claude auth status --json 결과로 판단한다 (토큰은 읽지 않음)."""
    path = find_claude_cli()
    st = {'installed': bool(path), 'logged_in': False, 'email': '', 'plan': '', 'path': path or ''}
    if not path:
        return st
    try:
        p = _run_claude(['auth', 'status', '--json'], timeout=30)
        data = json.loads(p.stdout or '{}')
    except (ClaudeCodeError, ValueError):
        return st
    if p.returncode == 0 and isinstance(data, dict) and data.get('loggedIn') is True \
            and data.get('authMethod') == 'claude.ai':
        st.update(logged_in=True, email=str(data.get('email') or ''),
                  plan=str(data.get('subscriptionType') or data.get('planTier') or ''))
    return st


def install_claude_code():
    """Claude Code 공식 Windows 설치 명령(PowerShell)을 숨김 창에서 실행하고 끝날 때까지 기다린다.
    출처: code.claude.com/docs 설치 안내 — irm https://claude.ai/install.ps1 | iex"""
    kw = {'creationflags': subprocess.CREATE_NO_WINDOW} if sys.platform == 'win32' else {}
    try:
        p = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
                            'irm https://claude.ai/install.ps1 | iex'],
                           stdin=subprocess.DEVNULL, capture_output=True, text=True,
                           encoding='utf-8', errors='replace', timeout=900, **kw)
    except FileNotFoundError:
        raise ClaudeCodeError('PowerShell을 찾지 못했어요. 공식 설치 안내를 보고 직접 설치해 주세요.')
    except subprocess.TimeoutExpired:
        raise ClaudeCodeError('설치가 너무 오래 걸려요. 인터넷 연결을 확인하고 다시 시도해 주세요.')
    if p.returncode != 0 or not find_claude_cli():
        tail = ((p.stderr or p.stdout or '').strip())[-200:]
        raise ClaudeCodeError('설치에 실패했어요. 인터넷 연결을 확인하고 다시 시도해 주세요.'
                              + (f'\n{tail}' if tail else ''))


def claude_code_login():
    """claude auth login --claudeai: 브라우저에서 Claude 공식 페이지로 로그인한다. 끝날 때까지 기다린다 (뒤 스레드에서 호출)."""
    p = _run_claude(['auth', 'login', '--claudeai'], timeout=CLAUDE_CODE_LOGIN_TIMEOUT)
    if p.returncode != 0:
        raise ClaudeCodeError('Claude 로그인이 끝나지 않았어요. 브라우저에서 승인했는지 확인하고 다시 시도해 주세요.')


def claude_code_logout():
    p = _run_claude(['auth', 'logout'], timeout=30)
    if p.returncode != 0:
        raise ClaudeCodeError('Claude 로그아웃에 실패했어요. ' + ((p.stderr or p.stdout or '').strip()[-200:]))


# ── OpenAIAdapter (ChatGPT, Responses API) ───────────────────────

class OpenAIAdapter:
    requires_api_key = True

    @staticmethod
    def get_api_key() -> str:
        return load_openai_key()

    @staticmethod
    def _client(key: str = None, **kw):
        from openai import OpenAI
        if OPENAI_BASE_URL:
            kw['base_url'] = OPENAI_BASE_URL
        return OpenAI(api_key=key if key is not None else load_openai_key(), **kw)

    @staticmethod
    def _input(prompt: str, image_b64: str = None):
        if not image_b64:
            return prompt
        return [{"role": "user", "content": [
            {"type": "input_image", "image_url": f"data:image/png;base64,{image_b64}"},
            {"type": "input_text", "text": prompt}]}]

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        model = load_openai_model()
        try:
            full, refused = "", False
            stream = cls._client().responses.create(
                model=model, instructions=system, input=prompt,
                max_output_tokens=8000,      # 먼저 '생각'하는 모델도 답을 끝까지 쓰도록 (쓴 만큼만 과금)
                stream=True)
            for ev in stream:
                t = getattr(ev, 'type', '')
                if t == 'response.output_text.delta':
                    full += ev.delta
                    on_chunk(ev.delta)
                elif t == 'response.refusal.delta':
                    refused = True
                elif t in ('response.failed', 'error'):
                    err = getattr(getattr(ev, 'response', None), 'error', None) or ev
                    on_error(f'AI 오류: {getattr(err, "message", "") or err}')
                    return
            if refused and not full:
                on_error(_REFUSAL)
                return
            on_done(full)
        except Exception as e:
            on_error(friendly_error(e, model, 'openai'))

    @classmethod
    def complete(cls, system: str, prompt: str, image_b64: str = None, max_tokens: int = 4000) -> str:
        r = cls._client().responses.create(model=load_openai_model(), instructions=system,
                                           input=cls._input(prompt, image_b64), max_output_tokens=max_tokens)
        text = (getattr(r, 'output_text', '') or '').strip()
        if not text and any(getattr(c, 'type', '') == 'refusal'
                            for o in (r.output or []) for c in (getattr(o, 'content', None) or [])):
            raise RuntimeError(_REFUSAL)
        return text


# ── OpenCodeAdapter (OpenCode Go/Zen) ────────────────────────────

class OpenCodeAdapter:
    """요금제에 맞는 주소와 필수 헤더를 붙이고, 모델에 따라
    OpenAI(chat-completions) 방식 또는 Claude(messages) 방식(예: claude-haiku-5-5)으로 요청한다."""
    requires_api_key = True

    @staticmethod
    def get_api_key() -> str:
        return load_opencode_key()

    @staticmethod
    def _model() -> str:
        return load_opencode_model()

    @staticmethod
    def _client(key: str = None, plan: str = None, **kw):
        from openai import OpenAI
        return OpenAI(api_key=key if key is not None else load_opencode_key(),
                      base_url=OPENCODE_PLANS[plan or load_opencode_plan()]['base_url'],
                      default_headers=opencode_headers(), **kw)

    @staticmethod
    def _anthropic_client(key: str, plan: str, **kw):
        import anthropic
        return anthropic.Anthropic(api_key=key, base_url=OPENCODE_PLANS[plan]['anthropic_base_url'],
                                   default_headers=opencode_headers(), **kw)

    @classmethod
    def _uses_messages(cls) -> bool:
        return opencode_model_protocol(cls._model()) == 'messages'

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        model = cls._model()
        try:
            if cls._uses_messages():
                client = cls._anthropic_client(cls.get_api_key(), load_opencode_plan())
                _messages_stream(client, model, system, prompt, on_chunk, on_done, on_error)
            else:
                _chat_stream(cls._client(), model, system, prompt, on_chunk, on_done)
        except Exception as e:
            on_error(friendly_error(e, model, 'opencode'))

    @classmethod
    def complete(cls, system: str, prompt: str, image_b64: str = None, max_tokens: int = 4000) -> str:
        model = cls._model()
        if image_b64 and opencode_model_supports_image(model) is False:
            raise RuntimeError(
                f'선택한 OpenCode 모델({model})은 이미지를 읽지 못합니다.\n'
                '메뉴 → 설정 → AI 연결(API 키) → OpenCode에서 🖼 표시가 있는 모델을 고르거나, '
                '개인정보 보호 모드를 켜서 글자로 분석하세요.')
        if cls._uses_messages():
            client = cls._anthropic_client(cls.get_api_key(), load_opencode_plan())
            return _messages_complete(client, model, system, prompt, image_b64=image_b64, max_tokens=max_tokens)
        return _chat_complete(cls._client(), model, system, prompt, image_b64=image_b64, max_tokens=max_tokens)


# ── 어댑터 선택 ──────────────────────────────────────────────────

_ADAPTERS = {'claude': ClaudeAdapter, 'openai': OpenAIAdapter, 'opencode': OpenCodeAdapter,
             'claude_code': ClaudeCodeAdapter}


def get_adapter():
    """현재 AI 모드의 어댑터. AI 꺼짐('none')이면 None."""
    return _ADAPTERS.get(load_ai_mode())


def masking_required() -> bool:
    """AI로 보내는 글의 개인정보를 가려야 하는지: 개인정보 보호 모드가 켜져 있으면 True."""
    try:
        from school_profile import privacy_enabled
        return privacy_enabled()
    except Exception:
        return False


def text_only_required() -> bool:
    """이미지를 AI에 보내지 않고 PC에서 읽은 글자만 보내야 하는지: 이 AI가 이미지를 읽지 못하거나(Claude 구독),
    개인정보 보호 모드가 켜져 있으면 True."""
    if get_adapter() is not None and not getattr(get_adapter(), 'supports_image', True):
        return True
    try:
        from school_profile import privacy_enabled
        return privacy_enabled()
    except Exception:
        return False


_get_adapter = get_adapter     # 예전 이름

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
        self.masked_counts = {}

    def stop(self):
        self._is_running = False
        cancel_claude_code()      # Claude 구독을 쓰는 중이면 실행 중인 claude도 끊는다

    def _retry_prompt(self, error_msg: str) -> str:
        return (
            self.prompt
            + "\n\n[시스템 지시] 이전 답변이 형식에 맞지 않았습니다.\n"
            + f"오류: {error_msg}\n"
            + '반드시 {"body": "여기에 공문 본문 전체를 작성"} JSON 형식으로만 출력하세요.'
        )

    def run(self):
        adapter = get_adapter()
        if adapter is None:
            self.error_signal.emit("AI가 꺼져 있습니다. [AI 설정]에서 OpenCode·Claude·ChatGPT·Claude 구독 중 하나를 연결하세요.")
            return
        if not adapter.get_api_key():
            self.error_signal.emit("API 키가 설정되지 않았습니다. [AI 설정]을 눌러 키를 입력하세요.")
            return

        # 개인정보 보호 모드가 켜져 있으면 AI로 나가는 글에서 개인정보를 가린다
        try:
            from school_profile import mask_for_ai
            if masking_required():
                self.prompt, self.masked_counts = mask_for_ai(self.prompt)
        except Exception:
            pass  # 설정을 못 읽어도 AI 기능은 계속 동작

        kwargs = {}
        if self.system is not None:
            kwargs['system'] = self.system

        if not self.structured:
            # 일반 스트리밍 모드 (자유 채팅 등)
            def on_chunk(text):
                if self._is_running:
                    self.new_text_signal.emit(text)

            def on_done(full):
                if self._is_running:
                    self.finished_signal.emit(full)

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

            def on_error(msg):
                err[0] = msg

            adapter.stream_text(self.prompt, on_chunk, lambda full: None, on_error, **kwargs)

            if err[0]:
                if attempt == 0:
                    self.prompt = self._retry_prompt(err[0])
                    continue
                self.error_signal.emit(err[0])
                return

            body = _parse_json_response(''.join(buffer)).get('body', '').strip()
            if len(body) > 20:
                # 성공: 본문을 스트리밍처럼 한 번에 전송
                self.new_text_signal.emit(body)
                self.finished_signal.emit(body)
                return

            if attempt == 0:
                self.prompt = self._retry_prompt('본문이 너무 짧음')

        self.error_signal.emit("AI 응답을 올바르게 받지 못했습니다. 다시 시도해 주세요.")
