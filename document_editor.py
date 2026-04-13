import re
import sys
import qtawesome as qta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QWidget, QLabel, QLineEdit,
    QMessageBox, QInputDialog, QFrame, QGraphicsDropShadowEffect, QSpinBox
)
from PyQt5.QtCore import Qt, QTimer, QSize, QEvent
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor
from ai_client import (AiStreamThread, load_api_key, save_api_key,
                        load_ai_mode, save_ai_mode,
                        load_ollama_model, save_ollama_model, OLLAMA_HOST,
                        load_external_model_name,
                        load_gemini_model, save_gemini_model,
                        GEMINI_DEFAULT_MODELS, GeminiAdapter)
from db import save_official_document, get_official_documents


# ── 전체 QSS 테마 ─────────────────────────────────────────────────
STYLE = """
    QMainWindow {
        background-color: #F9FAFB;
    }
    QWidget {
        font-family: 'Pretendard GOV', 'Malgun Gothic';
        font-size: 12pt;
        color: #111827;
    }
    QLineEdit {
        background-color: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        padding: 6px 12px;
        color: #111827;
    }
    QLineEdit:focus {
        border: 2px solid #6366F1;
        background-color: #FAFAFA;
    }
    QTextEdit#editor {
        background-color: #FFFFFF;
        border: 5px solid #6366F1;
        border-radius: 10px;
        padding: 28px 32px;
        color: #374151;
    }
    QFrame#AIInputFrame {
        background-color: #FFFFFF;
        border: 1.5px solid #E5E7EB;
        border-radius: 14px;
    }
    QFrame#AIInputFrame QTextEdit {
        background: transparent;
        border: none;
        border-radius: 0;
    }
    QPushButton {
        background-color: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        padding: 6px 16px;
        font-weight: 600;
        color: #374151;
    }
    QPushButton:hover {
        background-color: #F3F4F6;
        border-color: #9CA3AF;
    }
    QPushButton:pressed {
        background-color: #E5E7EB;
    }
    QPushButton:disabled {
        color: #9CA3AF;
        border-color: #E5E7EB;
        background-color: #F9FAFB;
    }
    QPushButton#AIPrimaryButton {
        background-color: #6366F1;
        color: white;
        border: none;
        border-radius: 10px;
    }
    QPushButton#AIPrimaryButton:hover {
        background-color: #4F46E5;
    }
    QPushButton#AIPrimaryButton:disabled {
        background-color: #A5B4FC;
    }
    QPushButton#DraftButton {
        background-color: #EEF2FF;
        border: 1px solid #C7D2FE;
        color: #4338CA;
        font-weight: 700;
    }
    QPushButton#DraftButton:hover {
        background-color: #E0E7FF;
    }
    QPushButton#DraftButton:disabled {
        background-color: #F5F3FF;
        border-color: #E0E7FF;
        color: #A5B4FC;
    }
    QPushButton#HelpButton {
        background-color: #FEF3C7;
        border: 1px solid #FDE68A;
        color: #92400E;
        font-weight: 700;
    }
    QPushButton#HelpButton:hover {
        background-color: #FDE68A;
    }
    QPushButton#SaveButton {
        background-color: #6366F1;
        color: white;
        border: none;
        font-weight: 700;
    }
    QPushButton#SaveButton:hover {
        background-color: #4F46E5;
    }
    QPushButton#SaveButton:disabled {
        background-color: #A5B4FC;
    }
    QStatusBar {
        background-color: #F3F4F6;
        color: #6B7280;
        font-size: 10pt;
        border-top: 1px solid #E5E7EB;
    }
    QLabel {
        color: #6B7280;
        font-size: 11pt;
        background: transparent;
    }
"""

# ── 문서 유형 ─────────────────────────────────────────────────────

_DOC_TYPE_LABELS = {
    'result_report':    '결과 보고',
    'home_letter':      '가정통신문 발송',
    'submit':           '제출',
    'recruit_announce': '채용 공고',
    'recruit_result':   '채용 결정',
    'purchase':         '품의/지출',
    'general':          '일반 계획/운영',
}


def _detect_doc_type(title: str) -> str:
    """제목 키워드로 문서 유형 자동 판별. 전각 괄호(［］)도 지원."""
    # 전각 괄호를 반각으로 정규화 후 비교
    t = title.strip().replace('［', '[').replace('］', ']')
    if '[공고]' in t or '채용 계획 공고' in t or '채용 계획 재공고' in t:
        return 'recruit_announce'
    if '[채용]' in t:
        return 'recruit_result'
    if '가정통신문 발송' in t:
        return 'home_letter'
    if '결과 보고' in t or re.search(r'결과\s*$', t):
        return 'result_report'
    if '제출' in t:
        return 'submit'
    if any(k in t for k in ['지출', '지급']):
        return 'purchase'
    return 'general'


def _eul_reul(text: str) -> str:
    """마지막 글자 받침 여부로 '을'/'를' 자동 선택."""
    if not text:
        return '을(를)'
    code = ord(text[-1]) - 0xAC00
    if 0 <= code < 11172 and code % 28 != 0:
        return '을'
    return '를'


def _extract_subject(title: str) -> str:
    """제목에서 연도·접두어를 제거한 핵심 사업명 추출."""
    s = re.sub(r'^\d{4}(학년도|년도|년)?\.?\s*', '', title)  # 앞 연도+점 제거
    s = re.sub(r'^\[.+?\]\s*', '', s)                        # [공고] 등 반각 제거
    s = re.sub(r'^［.+?］\s*', '', s)                         # ［채용］ 전각 제거
    return s.strip()


def _make_attach_block(first_text: str, count: int) -> str:
    """붙임 블록 생성. first_text는 첫 번째 항목 내용(1부 포함). count=0이면 빈 문자열."""
    if count == 0:
        return ''
    if count == 1:
        return f"붙임  {first_text}.  끝."
    lines = [f"붙임  1. {first_text}."]
    for i in range(2, count + 1):
        suffix = "  끝." if i == count else ""
        lines.append(f"      {i}. 1부.{suffix}")
    return '\n'.join(lines)


def _get_template(title: str, doc_type: str, ref: str = '', attach_count: int = 1) -> str:
    """문서 유형별 뼈대 템플릿 반환. ref가 있으면 관련번호 자동 삽입."""
    subject  = _extract_subject(title)
    josa     = _eul_reul(title)
    ref_line = f"1. 관련: {ref}\n" if ref else "1. 관련: \n"

    if doc_type == 'result_report':
        attach = re.sub(r'\s*(결과\s*보고|결과)\s*$', '', subject).strip()
        attach_block = _make_attach_block(f"{attach} 결과 1부", attach_count)
        body = (
            f"{ref_line}"
            f"2. {title}{josa} 다음과 같이 보고합니다.\n"
            "  가. 일시: \n"
            "  나. 장소: \n"
            "  다. 인원: \n"
            "  라. 내용: \n"
        )
        return body + ("\n" + attach_block if attach_block else "")
    elif doc_type == 'home_letter':
        attach = re.sub(r'\s*가정통신문\s*발송\s*$', '', subject).strip()
        attach_block = _make_attach_block(f"{attach} 가정통신문 1부", attach_count)
        body = (
            f"{ref_line}"
            f"2. {title}{josa} 붙임과 같이 발송하고자 합니다.\n"
        )
        return body + ("\n" + attach_block if attach_block else "")
    elif doc_type == 'submit':
        attach_word = re.sub(r'\s*제출\s*$', '', subject).strip()
        josa_submit = _eul_reul(attach_word)
        attach_block = _make_attach_block(f"{attach_word} 1부", attach_count)
        body = (
            f"{ref_line}"
            f"2. {attach_word}{josa_submit} 붙임과 같이 제출하고자 합니다.\n"
        )
        return body + ("\n" + attach_block if attach_block else "")
    elif doc_type == 'recruit_announce':
        attach = re.sub(r'\s*(채용\s*계획\s*(재)?공고(\(.+?\))?|공고(\(.+?\))?)$', '', subject).strip()
        return (
            f"{ref_line}"
            f"2. {title}{josa} 다음과 같이 공고하고자 합니다.\n"
            "  가. 채용예정인원 및 과목:\n"
            "      구분 | 채용 과목 | 선발예정 인원 | 채용기간 | 비고\n\n"
            "  나. 접수기간: \n"
            "  다. 접수장소: \n"
            "  라. 공고방법: 인터넷 홈페이지(전라남도교육청)\n"
            "  마. 공고일자: \n\n"
            f"붙임  1. {attach} 공고문 1부.\n"
            "      2. 서류평가 및 면접평가표 1부.(별첨)\n"
            "      3. 서류 및 면접 전형 평가표 1부.(별첨)  끝."
        )
    elif doc_type == 'recruit_result':
        return (
            f"{ref_line}"
            f"2. {title}{josa} 다음과 같이 채용하고자 합니다.\n"
            "  가. 채용 대상자:\n"
            "      연번 | 학교명 | 성명 | 생년월일 | 자격증 | 운용요일 | 비고\n\n"
            "  나. 강사 채용 기간: \n"
            "  다. 강사 채용 과목: \n"
            "  라. 강사 근무 시간: \n"
            "  마. 강사 수당: \n\n"
            "붙임  1. 서류(적부) 전형 평가표 1부.(별첨)\n"
            "      2. 응시원서 및 자기소개서 1부.(별첨)\n"
            "      3. 채용계약서 1부.(별첨)  끝."
        )
    elif doc_type == 'purchase':
        return (
            f"{ref_line}"
            f"2. {title}{josa} 다음과 같이 지출하고자 합니다.\n"
            "  가. 일시: \n"
            "  나. 장소: \n"
            "  다. 대상: \n"
            "  라. 내용: \n"
            "  마. 내역: \n"
            "  바. 예상금액: 원.  끝."
        )
    else:  # general
        attach = re.sub(r'\s*(운영\s*)?계획\s*$', '', subject).strip()
        attach_block = _make_attach_block(f"{attach} 계획 1부", attach_count)
        body = (
            f"{ref_line}"
            f"2. {title}{josa} 다음과 같이 운영하고자 합니다.\n"
            "  가. 일시: \n"
            "  나. 장소: \n"
            "  다. 대상: \n"
            "  라. 내용: \n"
        )
        return body + ("\n" + attach_block if attach_block else "")


class DocumentEditorWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ai_thread = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("AI 공문 도우미")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self._restore_window_state()  # 저장된 위치/크기 복원, 없으면 기본값

        _font = QFont('Pretendard GOV', 12)
        _font.setStyleHint(QFont.SansSerif)
        _input_h = 36  # 기본 ~28px의 130%

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 16, 20, 16)

        # ── 헤더: 제목 ──
        header = QHBoxLayout()
        header.setSpacing(8)
        self.title_input = QLineEdit()
        self.title_input.setFixedHeight(_input_h)
        self.title_input.setPlaceholderText("공문 제목을 입력하세요")
        self.title_input.textChanged.connect(self._update_draft_button)
        self.title_input.returnPressed.connect(self._try_generate_draft)
        header.addWidget(QLabel("공문 제목"))
        header.addWidget(self.title_input)

        # ── 관련 공문번호 ──
        ref_row = QHBoxLayout()
        ref_row.setSpacing(8)
        self.ref_input = QLineEdit()
        self.ref_input.setFixedHeight(_input_h)
        self.ref_input.setPlaceholderText("관련 공문번호 (선택)")
        self.ref_input.textChanged.connect(self._update_draft_button)
        self.ref_input.returnPressed.connect(self._try_generate_draft)
        btn_move_to_title = QPushButton()
        btn_move_to_title.setMinimumHeight(_input_h)
        btn_move_to_title.setFixedWidth(_input_h)
        btn_move_to_title.setToolTip("공문 제목으로 이동")
        _arrow_path = __import__('os').path.join(
            __import__('os').path.dirname(__import__('os').path.abspath(__file__)), '위화살표.png'
        )
        btn_move_to_title.setIcon(QIcon(QPixmap(_arrow_path).scaled(
            _input_h - 8, _input_h - 8, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )))
        btn_move_to_title.setIconSize(QSize(_input_h - 8, _input_h - 8))
        btn_move_to_title.clicked.connect(self._move_ref_to_title)
        btn_help = QPushButton("도움말")
        btn_help.setObjectName("HelpButton")
        btn_help.setMinimumHeight(_input_h)
        btn_help.clicked.connect(self._show_help)
        lbl_attach = QLabel("붙임 개수")
        lbl_attach.setFixedHeight(_input_h)
        self.spin_attach = QSpinBox()
        self.spin_attach.setRange(0, 9)
        self.spin_attach.setValue(1)
        self.spin_attach.setFixedHeight(_input_h)
        self.spin_attach.setFixedWidth(52)
        self.spin_attach.setToolTip("붙임 항목 수 (0 = 붙임 없음)")
        self.spin_attach.lineEdit().returnPressed.connect(self._try_generate_draft)
        ref_row.addWidget(QLabel("관련 번호"))
        ref_row.addWidget(self.ref_input, stretch=1)
        ref_row.addWidget(btn_move_to_title)
        ref_row.addSpacing(10)
        ref_row.addWidget(lbl_attach)
        ref_row.addWidget(self.spin_attach)
        ref_row.addStretch(1)
        ref_row.addWidget(btn_help)

        # ── 본문 편집 영역 ──
        self.editor = QTextEdit()
        self.editor.setObjectName("editor")
        _editor_font = QFont('굴림체', 12)
        _editor_font.setStretch(100)
        _editor_font.setLetterSpacing(QFont.PercentageSpacing, 100.0)
        self.editor.setFont(_editor_font)
        self.editor.document().blockCountChanged.connect(self._apply_editor_line_height)
        self.editor.setPlaceholderText(
            "자동 초안 생성 버튼을 누르거나, 아래 채팅창에 요청을 입력하세요."
        )
        _shadow = QGraphicsDropShadowEffect()
        _shadow.setBlurRadius(20)
        _shadow.setXOffset(0)
        _shadow.setYOffset(2)
        _shadow.setColor(QColor(0, 0, 0, 30))
        self.editor.setGraphicsEffect(_shadow)

        # ── AI 채팅바 (캡슐 프레임) ──
        chat_row = QHBoxLayout()
        chat_row.setSpacing(0)
        ai_frame = QFrame()
        ai_frame.setObjectName("AIInputFrame")
        ai_frame_layout = QHBoxLayout(ai_frame)
        ai_frame_layout.setContentsMargins(12, 4, 4, 4)
        ai_frame_layout.setSpacing(6)
        self.ai_input = QTextEdit()
        self.ai_input.setPlaceholderText("AI에게 요청")
        self.ai_input.setFixedHeight(_input_h * 2)
        self.ai_input.installEventFilter(self)
        self.editor.installEventFilter(self)
        self.btn_send = QPushButton()
        self.btn_send.setObjectName("AIPrimaryButton")
        self.btn_send.setIcon(qta.icon('fa5s.paper-plane', color='white'))
        self.btn_send.setIconSize(QSize(16, 16))
        self.btn_send.setFixedSize(52, _input_h * 2 - 8)
        self.btn_send.setToolTip("전송 (Enter)")
        self.btn_send.clicked.connect(self.ask_ai)
        ai_frame_layout.addWidget(self.ai_input)
        ai_frame_layout.addWidget(self.btn_send)
        chat_row.addWidget(ai_frame)

        # ── 컨트롤바 ──
        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(8)
        self.btn_draft = QPushButton("자동 초안 생성")
        self.btn_draft.setObjectName("DraftButton")
        self.btn_draft.setMinimumHeight(_input_h)
        self.btn_draft.clicked.connect(self.generate_draft)
        self.btn_apikey = QPushButton("AI 설정")
        self.btn_apikey.setMinimumHeight(_input_h)
        self.btn_apikey.clicked.connect(self._show_ai_settings_menu)
        self.btn_copy_title = QPushButton("제목 복사")
        self.btn_copy_title.setMinimumHeight(_input_h)
        self.btn_copy_title.clicked.connect(self._copy_title)
        self.btn_copy = QPushButton("본문 복사")
        self.btn_copy.setMinimumHeight(_input_h)
        self.btn_copy.clicked.connect(self._copy_body)
        self.btn_save_exit = QPushButton("저장 후 나가기")
        self.btn_save_exit.setObjectName("SaveButton")
        self.btn_save_exit.setMinimumHeight(_input_h)
        self.btn_save_exit.clicked.connect(self.save_and_close)
        ctrl_row.addWidget(self.btn_draft)
        ctrl_row.addWidget(self.btn_apikey)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self.btn_copy_title)
        ctrl_row.addWidget(self.btn_copy)
        ctrl_row.addWidget(self.btn_save_exit)

        # ── 전체 배치 ──
        layout.addLayout(header)
        layout.addLayout(ref_row)
        layout.addWidget(self.editor, stretch=1)
        layout.addLayout(chat_row)
        layout.addLayout(ctrl_row)

        self.setStyleSheet(STYLE)
        self.statusBar().showMessage("준비 완료")
        self._apply_ai_mode()
        self._update_draft_button()  # 초기 상태: 빈칸이면 비활성
        self._apply_editor_line_height()

    # ── AI 초안 생성 ──
    def generate_draft(self):
        title = self.title_input.text().strip()
        if not title:
            self.statusBar().showMessage("제목을 먼저 입력하세요.")
            return

        doc_type   = _detect_doc_type(title)
        type_label = _DOC_TYPE_LABELS.get(doc_type, '일반 계획/운영')
        ref        = self.ref_input.text().strip()
        template   = _get_template(title, doc_type, ref, self.spin_attach.value())
        mode       = load_ai_mode()

        # 1단계: 뼈대 템플릿 즉시 삽입 (AI 없이도 바로 사용 가능)
        self.editor.setPlainText(template)
        self._apply_editor_line_height()

        if mode == 'none':
            self.statusBar().showMessage(f'템플릿 삽입 완료 — {type_label}')
            return

        # 2단계: AI가 빈칸 채우기
        from ai_client import (FEWSHOT_PLAN_SYSTEM, FEWSHOT_PURCHASE_SYSTEM,
                                FEWSHOT_DEFAULT_SYSTEM, FEWSHOT_SUBMIT_SYSTEM)
        if doc_type == 'purchase':
            system = FEWSHOT_PURCHASE_SYSTEM
        elif doc_type == 'submit':
            system = FEWSHOT_SUBMIT_SYSTEM
        elif doc_type == 'general':
            system = FEWSHOT_PLAN_SYSTEM
        else:
            system = FEWSHOT_DEFAULT_SYSTEM
        ref_line = f"관련: {ref}\n" if ref else ""
        prompt = f"제목: {title}\n{ref_line}"
        self.editor.clear()
        self._start_ai(prompt, system=system)

    # ── AI 입력창 Enter 키 → 전송 ──
    def eventFilter(self, obj, event):
        if obj is self.ai_input and event.type() == QEvent.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter) and not (event.modifiers() & Qt.ShiftModifier):
                self.ask_ai()
                return True
        if obj is self.editor and event.type() == QEvent.KeyPress:
            if event.modifiers() == Qt.ControlModifier:
                if event.key() == Qt.Key_C:
                    selected = self.editor.textCursor().selectedText().replace('\u2029', '\n')
                    if selected:
                        QApplication.clipboard().setText(selected)
                    return True
                if event.key() == Qt.Key_X:
                    cursor = self.editor.textCursor()
                    selected = cursor.selectedText().replace('\u2029', '\n')
                    if selected:
                        QApplication.clipboard().setText(selected)
                        cursor.removeSelectedText()
                    return True
        return super().eventFilter(obj, event)

    # ── 자유 AI 채팅 ──
    def ask_ai(self):
        user_text = self.ai_input.toPlainText().strip()
        if not user_text:
            return

        current = self.editor.toPlainText()
        if current:
            prompt = f"현재 작성된 공문:\n{current}\n\n사용자 요청: {user_text}"
        else:
            prompt = user_text

        self.editor.clear()
        self._start_ai(prompt)

    # ── 공통 AI 실행 ──
    def _start_ai(self, prompt, system=None):
        if self.ai_thread and self.ai_thread.isRunning():
            return

        self.statusBar().showMessage("AI가 생각 중...")
        self._set_controls_enabled(False)

        self.ai_thread = AiStreamThread(prompt, system=system)
        self.ai_thread.new_text_signal.connect(self._append_text)
        self.ai_thread.finished_signal.connect(self._on_finished)
        self.ai_thread.error_signal.connect(self._on_error)
        self.ai_thread.start()

    def _append_text(self, text):
        cursor = self.editor.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(text)
        self.editor.setTextCursor(cursor)

    def _on_finished(self, _):
        self.statusBar().showMessage("완료")
        self._set_controls_enabled(True)
        self.ai_input.clear()

    def _on_error(self, msg):
        self.statusBar().showMessage(f"오류: {msg}")
        self._set_controls_enabled(True)

    def _set_controls_enabled(self, enabled):
        mode = load_ai_mode()
        ai_available = mode != 'none'
        self.btn_send.setEnabled(enabled and ai_available)
        self.ai_input.setEnabled(enabled and ai_available)
        if enabled:
            self._update_draft_button()   # 입력값 기반으로 복원
        else:
            self.btn_draft.setEnabled(False)  # AI 실행 중엔 무조건 비활성

    def _update_draft_button(self):
        """제목·관련번호 입력 여부로 버튼 활성화. AI 모드에 따라 텍스트 변경."""
        has_input = bool(self.title_input.text().strip() or self.ref_input.text().strip())
        self.btn_draft.setEnabled(has_input)
        self.btn_draft.setText('자동 초안 생성')

    def _try_generate_draft(self):
        """Enter 키 → 버튼 활성 상태일 때만 초안 생성."""
        if self.btn_draft.isEnabled():
            self.generate_draft()

    def _copy_title(self):
        """공문 제목을 클립보드에 복사."""
        title = self.title_input.text().strip()
        if not title:
            self.statusBar().showMessage('복사할 제목이 없습니다.')
            return
        QApplication.clipboard().setText(title)
        self.statusBar().showMessage('제목이 클립보드에 복사되었습니다.')
        QTimer.singleShot(1500, lambda: self.statusBar().showMessage(''))

    def _copy_body(self):
        """본문 텍스트를 클립보드에 복사."""
        content = self.editor.toPlainText().strip()
        if not content:
            self.statusBar().showMessage('복사할 내용이 없습니다.')
            return
        QApplication.clipboard().setText(content)
        self.statusBar().showMessage('본문이 클립보드에 복사되었습니다.')
        QTimer.singleShot(1500, lambda: self.statusBar().showMessage(''))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self._cancel_ai()
        elif event.key() == Qt.Key_X and event.modifiers() == (Qt.ControlModifier | Qt.ShiftModifier):
            self._start_capture()
        else:
            super().keyPressEvent(event)

    def _show_help(self):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
        dlg = QDialog(self)
        dlg.setWindowTitle("사용 방법")
        dlg.setWindowFlags(dlg.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        dlg.setFixedSize(580, 535)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(6)

        _s = "font-family:'Malgun Gothic'; font-size:10pt;"
        _b = "font-family:'Malgun Gothic'; font-size:10pt; font-weight:bold;"
        _li = "margin:0; padding:0;"
        lbl = QLabel(
            f"<p style='{_b} margin-bottom:4px;'>📋&nbsp; 공문 제목 / 관련번호 캡처</p>"
            f"<p style='{_s} {_li}'>• Ctrl + Shift + X 를 누르면 캡처화면이 뜨면서 AI 공문 도우미가 실행됩니다.</p>"
            f"<p style='{_s} {_li}'>• 캡쳐하면 텍스트를 판별해서 공문제목 또는 관련번호에 적힙니다.</p>"
            f"<p style='{_s} {_li}'>• 두 칸 중 하나만 차있으면 반대쪽 빈 칸으로 자동 배분합니다.</p>"
            f"<br>"
            f"<p style='{_b} margin-bottom:4px;'>⚡&nbsp; 자동 초안 생성</p>"
            f"<p style='{_s} {_li}'>• 공문 제목 또는 관련번호 입력 후 Enter</p>"
            f"<br>"
            f"<p style='{_b} margin-bottom:4px;'>⛔&nbsp; AI 생성 취소</p>"
            f"<p style='{_s} {_li}'>• 생성 중 ESC 키를 누르면 즉시 중단됩니다.</p>"
        )
        lbl.setWordWrap(True)
        lbl.setFont(QFont('Malgun Gothic', 10))
        layout.addWidget(lbl)

        btn_ok = QPushButton("확인")
        btn_ok.setFixedWidth(80)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        dlg.exec_()

    def _cancel_ai(self):
        if self.ai_thread and self.ai_thread.isRunning():
            self.ai_thread.stop()
            self.ai_thread.wait(500)
            self.statusBar().showMessage('AI 생성이 취소되었습니다.')
            self._set_controls_enabled(True)

    def _apply_editor_line_height(self):
        """에디터 본문 줄간격 123% 적용."""
        from PyQt5.QtGui import QTextCursor, QTextBlockFormat
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.Document)
        fmt = QTextBlockFormat()
        fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
        cursor.mergeBlockFormat(fmt)

    def _apply_ai_mode(self):
        """현재 AI 모드에 따라 AI 관련 버튼 활성/비활성."""
        mode = load_ai_mode()
        ai_available = mode != 'none'
        self._update_draft_button()       # 입력값 + 모드 조합으로 판단
        self.btn_send.setEnabled(ai_available)
        self.ai_input.setEnabled(ai_available)
        if not ai_available:
            self.ai_input.setPlaceholderText('AI 이용 안함 — AI 설정에서 변경하세요.')
        elif mode == 'internal':
            self.ai_input.setPlaceholderText(
                f"AI에게 요청 ({load_ollama_model()} 사용중)"
            )
        else:
            self.ai_input.setPlaceholderText(
                f"AI에게 요청 ({load_external_model_name()} 사용중)"
            )

    # ── AI 설정 메뉴 ──
    def _show_ai_settings_menu(self):
        from PyQt5.QtWidgets import QMenu
        mode = load_ai_mode()
        menu = QMenu(self)

        act_external = menu.addAction('외부 AI 이용')
        act_internal = menu.addAction('내부 AI 이용')
        menu.addSeparator()
        act_none = menu.addAction('AI 이용 안함')

        # 현재 선택된 항목에 체크 표시
        act_external.setCheckable(True)
        act_internal.setCheckable(True)
        act_none.setCheckable(True)
        act_external.setChecked(mode == 'external')
        act_internal.setChecked(mode == 'internal')
        act_none.setChecked(mode == 'none')

        action = menu.exec_(self.btn_apikey.mapToGlobal(
            self.btn_apikey.rect().bottomLeft()
        ))

        if action == act_external:
            self._setup_external_ai()
        elif action == act_internal:
            self._setup_internal_ai()
        elif action == act_none:
            save_ai_mode('none')
            self._apply_ai_mode()
            self.statusBar().showMessage('AI 이용 안함 — OCR(winrt)만 사용합니다.')

    def _setup_external_ai(self):
        # 1단계: API 키 입력
        current_key = load_api_key()
        key, ok = QInputDialog.getText(
            self, '외부 AI 설정', 'Gemini API 키를 입력하세요:',
            text=current_key
        )
        if not ok:
            return
        if key.strip():
            save_api_key(key.strip())

        # 2단계: 모델 목록 조회
        self.statusBar().showMessage('모델 목록 조회 중...')
        QApplication.processEvents()
        models = GeminiAdapter.fetch_models() if load_api_key() else GEMINI_DEFAULT_MODELS

        # 3단계: 모델 선택 다이얼로그
        dlg = _GeminiModelDialog(models, load_gemini_model(), self)
        if dlg.exec_() != dlg.Accepted:
            return
        model = dlg.selected_model()
        if not model:
            return

        save_gemini_model(model)
        save_ai_mode('external')
        self._apply_ai_mode()
        self.statusBar().showMessage(f'외부 AI(Gemini · {model}) 설정 완료.')

    def _setup_internal_ai(self):
        # Ollama 연결 및 모델 목록 가져오기
        try:
            import requests
            resp = requests.get(f'{OLLAMA_HOST}/api/tags', timeout=5)
            resp.raise_for_status()
            available = [m['name'] for m in resp.json().get('models', [])]
        except Exception:
            QMessageBox.warning(
                self, 'Ollama 연결 실패',
                f'Ollama 서버({OLLAMA_HOST})에 연결할 수 없습니다.\n'
                'Ollama가 실행 중인지 확인하세요.'
            )
            return

        dlg = _OllamaModelDialog(available, load_ollama_model(), self)
        if dlg.exec_() != dlg.Accepted:
            return

        model = dlg.selected_model()
        if not model:
            return

        save_ollama_model(model)
        save_ai_mode('internal')
        self._apply_ai_mode()
        self.statusBar().showMessage(f'내부 AI(Ollama · {model}) 설정 완료.')

    # ── 저장 ──
    def save_document(self):
        title = self.title_input.text().strip()
        content = self.editor.toPlainText().strip()

        if not title:
            QMessageBox.warning(self, '저장 실패', '제목을 입력하세요.')
            return
        if not content:
            QMessageBox.warning(self, '저장 실패', '본문 내용이 없습니다.')
            return

        doc_number = self.ref_input.text().strip()

        existing = get_official_documents()
        for doc in existing:
            if (doc.get('title', '') == title and
                    doc.get('doc_number', '') == doc_number and
                    doc.get('content', '') == content):
                self.statusBar().showMessage('같은 내용이 저장되어 있습니다.')
                QTimer.singleShot(1000, lambda: self.statusBar().showMessage(''))
                return

        save_official_document(title, doc_number, content, '')
        self.statusBar().showMessage('저장되었습니다.')
        QTimer.singleShot(1000, lambda: self.statusBar().showMessage(''))
        return True

    def save_and_close(self):
        self.save_document()
        self.close()

    # ── 창 위치/크기 저장·복원 ────────────────────────────────────
    _STATE_FILE = __import__('os').path.join(
        __import__('os').environ.get('APPDATA', '.'), 'SSNnote', 'doc_editor_state.json'
    )

    def _restore_window_state(self):
        import json, os
        try:
            with open(self._STATE_FILE, 'r', encoding='utf-8') as f:
                s = json.load(f)
            self.setGeometry(s['x'], s['y'], s['width'], s['height'])
        except Exception:
            self.setGeometry(150, 150, 860, 640)

    def closeEvent(self, event):
        import json, os
        os.makedirs(os.path.dirname(self._STATE_FILE), exist_ok=True)
        g = self.geometry()
        with open(self._STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'x': g.x(), 'y': g.y(),
                       'width': g.width(), 'height': g.height()}, f)
        super().closeEvent(event)

    # ── 캡처 ──────────────────────────────────────────────────────
    def _start_capture(self):
        self.hide()
        QTimer.singleShot(150, self._show_capture_overlay)

    def _show_capture_overlay(self):
        from capture import grab_fullscreen, ScreenCaptureOverlay
        screenshot = grab_fullscreen()
        self._overlay = ScreenCaptureOverlay(screenshot)
        self._overlay.region_captured.connect(self._on_capture_complete)
        self._overlay.cancelled.connect(self._restore_after_capture)
        self._overlay.show()
        self._overlay.activateWindow()

    def _restore_after_capture(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def _move_ref_to_title(self):
        text = self.ref_input.text().strip()
        if text:
            self.title_input.setText(text)
            self.ref_input.clear()

    def _capture_target(self):
        """
        캡처 결과를 어느 칸에 넣을지 결정.
        반환값: 'title' | 'ref' | 'auto'
        - 둘 다 비어있으면 'auto' (AI/OCR 패턴으로 자동 판별)
        - 제목만 차있으면 'ref'
        - 관련번호만 차있으면 'title'
        - 둘 다 차있으면 'auto' (덮어쓰지 않도록 자동 판별)
        """
        has_title = bool(self.title_input.text().strip())
        has_ref   = bool(self.ref_input.text().strip())
        if has_title and not has_ref:
            return 'ref'
        if has_ref and not has_title:
            return 'title'
        return 'auto'

    def _on_capture_complete(self, pixmap):
        self._overlay.region_captured.disconnect()
        self._restore_after_capture()

        target = self._capture_target()
        self._ocr_fallback(pixmap, target)

    def _on_ai_result(self, result, target):
        title      = result.get('title', '').strip()
        doc_number = result.get('doc_number', '').strip()

        if target == 'title':
            self.title_input.setText(title or doc_number)
        elif target == 'ref':
            self.ref_input.setText(doc_number or title)
        else:  # auto
            if title:
                self.title_input.setText(title)
            if doc_number:
                self.ref_input.setText(doc_number)
        self.statusBar().showMessage('캡처 완료 (AI 인식)')

    def _ocr_fallback(self, pixmap, target='auto'):
        """API 키 없거나 AI 실패 시 Windows OCR로 폴백. winrt 없으면 안내."""
        try:
            import winrt  # noqa: F401
        except ImportError:
            self.statusBar().showMessage('OCR 불가 — API 키를 설정하거나 winsdk 패키지를 설치하세요.')
            QMessageBox.warning(
                self, 'OCR 불가',
                'Windows OCR(winrt) 모듈이 없습니다.\n\n'
                '① AI 설정 버튼으로 Gemini 키를 입력하거나\n'
                '② 터미널에서 pip install winsdk 를 실행하세요.'
            )
            return

        from capture import run_ocr, _normalize_doc_number
        self.statusBar().showMessage('OCR 인식 중...')

        def _after_ocr(text):
            text = text.strip()
            if target == 'title':
                self.title_input.setText(text)
                self.statusBar().showMessage('캡처 완료 → 제목')
            elif target == 'ref':
                self.ref_input.setText(text)
                self.statusBar().showMessage('캡처 완료 → 관련번호')
            else:  # auto
                normalized = _normalize_doc_number(text)
                if normalized != text:
                    self.ref_input.setText(normalized)
                    self.statusBar().showMessage('캡처 완료 (공문번호 인식)')
                else:
                    self.title_input.setText(text)
                    self.statusBar().showMessage('캡처 완료 (제목 인식)')

        def _on_error(msg):
            self.statusBar().showMessage(f'OCR 오류: {msg}')

        self._ocr_worker = run_ocr(pixmap, _after_ocr, _on_error)


class _OllamaModelDialog(
    __import__('PyQt5.QtWidgets', fromlist=['QDialog']).QDialog
):
    """Ollama 설치 모델 목록에서 클릭으로 선택하는 다이얼로그."""

    def __init__(self, models: list, current: str, parent=None):
        from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout,
                                     QListWidget, QListWidgetItem,
                                     QLineEdit, QPushButton, QLabel)
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QFont
        super().__init__(parent)
        self.setWindowTitle('내부 AI 모델 선택 (Ollama)')
        self.setMinimumWidth(320)
        self.setMinimumHeight(360)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        _font = QFont('Malgun Gothic', 11)
        self.setFont(_font)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(14, 14, 14, 14)

        layout.addWidget(QLabel('설치된 Ollama 모델 중 하나를 선택하세요.'))

        # 검색창
        self._search = QLineEdit()
        self._search.setPlaceholderText('모델 검색...')
        self._search.setMinimumHeight(32)
        self._search.textChanged.connect(self._filter)
        layout.addWidget(self._search)

        # 모델 목록
        self._list = QListWidget()
        self._list.setFont(_font)
        self._list.setSpacing(2)
        self._list.itemDoubleClicked.connect(self.accept)
        self._models = models
        self._populate(models, current)
        layout.addWidget(self._list, stretch=1)

        # 버튼
        btn_row = QHBoxLayout()
        btn_ok = QPushButton('선택')
        btn_ok.setMinimumHeight(34)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def _populate(self, models, select=''):
        from PyQt5.QtWidgets import QListWidgetItem
        from PyQt5.QtCore import Qt
        self._list.clear()
        for name in models:
            item = QListWidgetItem(name)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self._list.addItem(item)
            if name == select:
                self._list.setCurrentItem(item)
        if not self._list.currentItem() and self._list.count():
            self._list.setCurrentRow(0)

    def _filter(self, text):
        filtered = [m for m in self._models if text.lower() in m.lower()]
        current = self._list.currentItem()
        self._populate(filtered, current.text() if current else '')

    def selected_model(self) -> str:
        item = self._list.currentItem()
        return item.text() if item else ''


class _GeminiModelDialog(
    __import__('PyQt5.QtWidgets', fromlist=['QDialog']).QDialog
):
    """Gemini 모델 목록에서 클릭으로 선택하는 다이얼로그."""

    def __init__(self, models: list, current: str, parent=None):
        from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout,
                                     QListWidget, QListWidgetItem,
                                     QLineEdit, QPushButton, QLabel)
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QFont
        super().__init__(parent)
        self.setWindowTitle('외부 AI 모델 선택 (Gemini)')
        self.setMinimumWidth(360)
        self.setMinimumHeight(360)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        _font = QFont('Pretendard GOV', 11)
        self.setFont(_font)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(14, 14, 14, 14)

        layout.addWidget(QLabel('사용할 Gemini 모델을 선택하세요.'))

        # 검색창
        self._search = QLineEdit()
        self._search.setPlaceholderText('모델 검색...')
        self._search.setMinimumHeight(32)
        self._search.textChanged.connect(self._filter)
        layout.addWidget(self._search)

        # 모델 목록
        self._list = QListWidget()
        self._list.setFont(_font)
        self._list.setSpacing(2)
        self._list.itemDoubleClicked.connect(self.accept)
        self._models = models
        self._populate(models, current)
        layout.addWidget(self._list, stretch=1)

        # 버튼
        btn_row = QHBoxLayout()
        btn_ok = QPushButton('선택')
        btn_ok.setMinimumHeight(34)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def _populate(self, models, select=''):
        from PyQt5.QtWidgets import QListWidgetItem
        from PyQt5.QtCore import Qt
        self._list.clear()
        for name in models:
            item = QListWidgetItem(name)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self._list.addItem(item)
            if name == select:
                self._list.setCurrentItem(item)
        if not self._list.currentItem() and self._list.count():
            self._list.setCurrentRow(0)

    def _filter(self, text):
        filtered = [m for m in self._models if text.lower() in m.lower()]
        current = self._list.currentItem()
        self._populate(filtered, current.text() if current else '')

    def selected_model(self) -> str:
        item = self._list.currentItem()
        return item.text() if item else ''


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentEditorWindow()
    window.show()
    sys.exit(app.exec_())
