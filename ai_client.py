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

APPDATA_DIR  = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
KEY_FILE     = os.path.join(APPDATA_DIR, 'gemini_key.txt')
AI_MODE_FILE = os.path.join(APPDATA_DIR, 'ai_mode.txt')

# AI 모드: 'external' | 'internal' | 'none'
_AI_MODES = ('external', 'internal', 'none')

DOC_SYSTEM_PROMPT = (
    "당신은 한국 학교 행정 공문 작성 전문가입니다. "
    "공식 공문 문체(개조식, '~함', '~바람', '~고자 함')로만 작성하세요. "
    "불필요한 설명 없이 공문 본문만 작성하세요."
)

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
    def stream_text(cls, prompt: str, on_chunk, on_done, on_error, system=DOC_SYSTEM_PROMPT):
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


# ── 현재 사용할 어댑터 (여기만 바꾸면 전체 교체) ──────────────────
_ADAPTER = GeminiAdapter


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
        if getattr(_ADAPTER, 'requires_api_key', True) and not load_api_key():
            self.error_signal.emit("API 키가 설정되지 않았습니다. 설정 버튼을 눌러 키를 입력하세요.")
            return

        def on_chunk(text):
            if self._is_running:
                self.new_text_signal.emit(text)

        def on_done(full):
            if self._is_running:
                self.finished_signal.emit(full)

        _ADAPTER.stream_text(self.prompt, on_chunk, on_done, self.error_signal.emit)


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
        if getattr(_ADAPTER, 'requires_api_key', True) and not load_api_key():
            self.error.emit("API 키가 설정되지 않았습니다.")
            return
        result = _ADAPTER.ask_with_image(self._PROMPT, self._image_b64)
        if "error" in result:
            self.error.emit(result["error"])
        else:
            self.finished.emit(result)
