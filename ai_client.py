"""
ai_client.py — AI 제공자 추상화 레이어

나중에 AI 제공자를 바꿀 때 이 파일만 수정하면 됩니다.
현재 구현: Gemini (google-generativeai)

다른 제공자로 교체 시:
  - ClaudeAdapter  : anthropic SDK, ask()/ask_with_image() 동일 인터페이스
  - OpenAIAdapter  : openai SDK, 동일 인터페이스
  - OllamaAdapter  : 로컬 REST (Gemma4 등), 동일 인터페이스
"""

import os
import base64
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QPixmap

APPDATA_DIR        = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
KEY_FILE           = os.path.join(APPDATA_DIR, 'gemini_key.txt')
AI_MODE_FILE       = os.path.join(APPDATA_DIR, 'ai_mode.txt')
OLLAMA_MODEL_FILE  = os.path.join(APPDATA_DIR, 'ollama_model.txt')
OLLAMA_HOST        = 'http://localhost:11434'

# AI 모드: 'external' | 'internal' | 'none'
_AI_MODES = ('external', 'internal', 'none')

# 외부 AI(Gemini 등 대형 모델): 형식 제약 최소화, 모델 자율성 최대
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

OCR_SYSTEM_PROMPT = (
    "당신은 한국 행정 공문서 이미지 분석 전문가입니다. "
    "이미지에서 텍스트를 정확하게 추출하세요."
)


# ── API 키 / AI 모드 관리 ────────────────────────────────────────

def load_api_key() -> str:
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def load_ai_mode() -> str:
    """저장된 AI 모드 반환. 기본값: 'external'."""
    if os.path.exists(AI_MODE_FILE):
        with open(AI_MODE_FILE, 'r', encoding='utf-8') as f:
            mode = f.read().strip()
            if mode in _AI_MODES:
                return mode
    return 'external'


def save_ai_mode(mode: str):
    if mode not in _AI_MODES:
        raise ValueError(f'유효하지 않은 모드: {mode}')
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(AI_MODE_FILE, 'w', encoding='utf-8') as f:
        f.write(mode)


def save_api_key(key: str):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(KEY_FILE, 'w', encoding='utf-8') as f:
        f.write(key.strip())


def load_ollama_model() -> str:
    """저장된 Ollama 모델명 반환. 기본값: 'gemma3:2b'."""
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


# ── 내부 유틸 ────────────────────────────────────────────────────

def _pixmap_to_base64(pixmap: QPixmap) -> str:
    """QPixmap → PNG base64 문자열."""
    from PyQt5.QtCore import QBuffer, QIODevice
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    pixmap.save(buf, 'PNG')
    buf.close()
    return base64.b64encode(bytes(buf.data())).decode('utf-8')


# ── GeminiAdapter ────────────────────────────────────────────────
# 다른 제공자로 바꿀 때: 아래 클래스만 교체하거나 새 Adapter 추가 후
# AiStreamThread / ask_with_image_sync 의 import 대상을 변경하면 됩니다.

class GeminiAdapter:
    requires_api_key = True
    MODEL_TEXT  = "models/gemini-2.0-flash-lite"
    MODEL_VISION = "models/gemini-2.0-flash-lite"

    @staticmethod
    def _client():
        import google.generativeai as genai
        genai.configure(api_key=load_api_key())
        return genai

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=EXTERNAL_DOC_SYSTEM_PROMPT):
        """텍스트 스트리밍. on_chunk(str), on_done(full_str), on_error(msg)."""
        try:
            genai = cls._client()
            model = genai.GenerativeModel(cls.MODEL_TEXT, system_instruction=system)
            response = model.generate_content(prompt, stream=True)
            full = ""
            for chunk in response:
                text = chunk.text or ""
                full += text
                if text:
                    on_chunk(text)
            on_done(full)
        except Exception as e:
            on_error(str(e))

    @classmethod
    def ask_with_image(cls, prompt: str, image_b64: str) -> dict:
        """
        이미지 + 텍스트 프롬프트를 보내고 JSON dict 반환.
        반환 형식: {"title": str, "doc_number": str}
        인식 실패 시: {"title": "", "doc_number": "", "raw": str}
        """
        import json, re
        try:
            import google.generativeai as genai
            from google.generativeai.types import HarmCategory, HarmBlockThreshold
            genai.configure(api_key=load_api_key())
            model = genai.GenerativeModel(
                cls.MODEL_VISION,
                system_instruction=OCR_SYSTEM_PROMPT
            )
            image_part = {
                "inline_data": {
                    "mime_type": "image/png",
                    "data": image_b64
                }
            }
            response = model.generate_content([prompt, image_part])
            raw = response.text.strip()

            # JSON 파싱 (마크다운 코드블록 제거 후 시도)
            cleaned = re.sub(r'^```[a-z]*\n?', '', raw, flags=re.MULTILINE)
            cleaned = re.sub(r'```$', '', cleaned, flags=re.MULTILINE).strip()
            try:
                return json.loads(cleaned)
            except json.JSONDecodeError:
                return {"title": "", "doc_number": "", "raw": raw}
        except Exception as e:
            return {"title": "", "doc_number": "", "error": str(e)}


# ── OllamaAdapter ────────────────────────────────────────────────

class OllamaAdapter:
    requires_api_key = False

    @staticmethod
    def _model():
        return load_ollama_model()

    @classmethod
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=INTERNAL_DOC_SYSTEM_PROMPT):
        """Ollama /api/generate 스트리밍."""
        import json, requests
        url = f'{OLLAMA_HOST}/api/generate'
        model = cls._model()
        payload = {
            'model': model,
            'prompt': f'{system}\n\n{prompt}',
            'stream': True,
            'think': False,   # qwen3 등 thinking 모드 비활성화 (빠른 응답)
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
                    # Ollama가 스트림 중 오류를 내는 경우
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
        """Ollama /api/generate — 이미지 지원(비전 모델인 경우)."""
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
            raw = resp.json().get('response', '').strip()
            cleaned = __import__('re').sub(r'^```[a-z]*\n?', '', raw, flags=__import__('re').MULTILINE)
            cleaned = __import__('re').sub(r'```$', '', cleaned, flags=__import__('re').MULTILINE).strip()
            try:
                return json.loads(cleaned)
            except json.JSONDecodeError:
                return {'title': '', 'doc_number': '', 'raw': raw}
        except requests.exceptions.ConnectionError:
            return {'title': '', 'doc_number': '', 'error': 'Ollama 서버에 연결할 수 없습니다.'}
        except Exception as e:
            return {'title': '', 'doc_number': '', 'error': str(e)}


# ── 현재 사용할 어댑터 (여기만 바꾸면 전체 교체) ──────────────────
_ADAPTER = GeminiAdapter


def _get_adapter():
    """현재 AI 모드에 맞는 어댑터 반환."""
    mode = load_ai_mode()
    if mode == 'internal':
        return OllamaAdapter
    return GeminiAdapter


# ── 공개 인터페이스 ──────────────────────────────────────────────

class AiStreamThread(QThread):
    """
    텍스트 스트리밍 스레드.
    document_editor.py 의 OllamaThread 를 대체합니다.
    """
    new_text_signal  = pyqtSignal(str)
    finished_signal  = pyqtSignal(str)
    error_signal     = pyqtSignal(str)

    def __init__(self, prompt: str):
        super().__init__()
        self.prompt = prompt
        self._is_running = True

    def stop(self):
        self._is_running = False

    def run(self):
        adapter = _get_adapter()
        if getattr(adapter, 'requires_api_key', True) and not load_api_key():
            self.error_signal.emit("API 키가 설정되지 않았습니다. 설정 버튼을 눌러 키를 입력하세요.")
            return

        def on_chunk(text):
            if self._is_running:
                self.new_text_signal.emit(text)

        def on_done(full):
            if self._is_running:
                self.finished_signal.emit(full)

        adapter.stream_text(self.prompt, on_chunk, on_done, self.error_signal.emit)


class AiImageWorker(QThread):
    """
    이미지 → 공문 제목/번호 추출 스레드.
    finished 시그널: {"title": str, "doc_number": str} 딕셔너리 전달.
    """
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
        if getattr(adapter, 'requires_api_key', True) and not load_api_key():
            self.error.emit("API 키가 설정되지 않았습니다.")
            return
        result = adapter.ask_with_image(self._PROMPT, self._image_b64)
        if "error" in result:
            self.error.emit(result["error"])
        else:
            self.finished.emit(result)
