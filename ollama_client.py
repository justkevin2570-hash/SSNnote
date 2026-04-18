import os
import google.generativeai as genai
from PyQt5.QtCore import QThread, pyqtSignal

APPDATA_DIR = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote')
KEY_FILE = os.path.join(APPDATA_DIR, 'gemini_key.txt')

SYSTEM_PROMPT = (
    "당신은 한국 학교 행정 공문 작성 전문가입니다. "
    "공식 공문 문체(개조식, '~함', '~바람', '~고자 함')로만 작성하세요. "
    "불필요한 설명 없이 공문 본문만 작성하세요."
)


def load_api_key():
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return ''


def save_api_key(key):
    os.makedirs(APPDATA_DIR, exist_ok=True)
    with open(KEY_FILE, 'w', encoding='utf-8') as f:
        f.write(key.strip())


class OllamaThread(QThread):
    new_text_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, prompt, model="models/gemini-3.1-flash-lite-preview"):
        super().__init__()
        self.prompt = prompt
        self.model = model
        self._is_running = True

    def stop(self):
        self._is_running = False

    def run(self):
        api_key = load_api_key()
        if not api_key:
            self.error_signal.emit("API 키가 설정되지 않았습니다. 설정 버튼을 눌러 키를 입력하세요.")
            return

        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                self.model,
                system_instruction=SYSTEM_PROMPT
            )
            response = model.generate_content(self.prompt, stream=True)

            full_response = ""
            for chunk in response:
                if not self._is_running:
                    break
                text = chunk.text if chunk.text else ""
                full_response += text
                if text:
                    self.new_text_signal.emit(text)

            if self._is_running:
                self.finished_signal.emit(full_response)

        except Exception as e:
            self.error_signal.emit(f"오류 발생: {str(e)}")
