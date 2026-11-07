import re
import sys
import os

def _base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))
import qtawesome as qta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QWidget, QLabel, QLineEdit,
    QMessageBox, QInputDialog, QFrame, QSpinBox, QCheckBox, QDialog,
)
from PyQt5.QtCore import Qt, QTimer, QSize, QEvent
from PyQt5.QtGui import QFont, QTextCursor, QTextCharFormat
from ai_client import (AiStreamThread, PROVIDER_LABELS,
                        load_ai_mode, save_ai_mode, legacy_ai_name, load_external_model_name,
                        load_claude_key, save_claude_key, load_claude_model, save_claude_model,
                        fetch_claude_models, test_claude,
                        load_openai_key, save_openai_key, load_openai_model, save_openai_model,
                        fetch_openai_models, test_openai,
                        load_opencode_key, save_opencode_key,
                        load_opencode_model, save_opencode_model,
                        OPENCODE_PLANS, load_opencode_plan, save_opencode_plan,
                        fetch_opencode_models, save_opencode_model_info, test_opencode)
from db import save_official_document, get_official_documents, get_all_embeddings, save_embedding
# rag(유사 공문 찾기)는 numpy·sentence-transformers가 필요해 쓸 때만 불러온다 — 없으면 그 기능만 꺼지고 창은 열림
from ui.theme import keep_light  # 예산·공문 창은 라이트 전용
from ui.common import fit_to_screen, default_geometry, px


# ── 색 (차분한 회청색 + 포인트 색 하나) ──────────────────────────
class C:
    BG        = '#F5F6F8'    # 창 바탕
    SURFACE   = '#FFFFFF'    # 카드·종이
    BORDER    = '#E3E6EB'
    BORDER_HI = '#CBD2DC'    # 마우스 올림
    TEXT      = '#1E293B'
    BODY      = '#2B3443'    # 공문 본문 글자
    MUTED     = '#64748B'
    FAINT     = '#94A3B8'
    HOVER     = '#F1F3F6'
    ACCENT    = '#4F46E5'    # 포인트 (저장·전송·초안)
    ACCENT_HI = '#4338CA'
    ACCENT_SOFT = '#EEF0FD'
    ACCENT_LINE = '#C7CCF7'
    FOCUS     = '#A5B0F5'
    OK        = '#16A34A'
    OFF       = '#94A3B8'
    WARN      = '#D97706'


_FONT = "'Pretendard GOV', 'Pretendard', 'Malgun Gothic'"

def _arrow_png(name: str, icon: str, color: str = C.MUTED) -> str:
    """스핀 박스 화살표 그림 (qtawesome → 임시 PNG). 실패하면 '' (Qt 기본 화살표)."""
    try:
        import tempfile
        path = os.path.join(tempfile.gettempdir(), f'ssnnote_{name}.png')
        if not os.path.exists(path):
            qta.icon(icon, color=color).pixmap(16, 16).save(path)
        return path.replace('\\', '/')
    except Exception:
        return ''


# ── 전체 QSS 테마 ─────────────────────────────────────────────────
_STYLE_TEMPLATE = f"""
    QMainWindow, QDialog {{ background-color: {C.BG}; }}
    QWidget {{ font-family: {_FONT}; font-size: 11pt; color: {C.TEXT}; }}
    QLabel {{ background: transparent; color: {C.MUTED}; font-size: 10pt; }}
    QLabel#WindowTitle {{ color: {C.TEXT}; font-size: 13pt; font-weight: 700; }}
    QLabel#FieldLabel {{ color: {C.MUTED}; font-size: 10pt; font-weight: 600; }}
    QLabel#DialogTitle {{ color: {C.TEXT}; font-size: 12pt; font-weight: 700; }}
    QLabel#Hint {{ color: {C.MUTED}; font-size: 9pt; }}

    QFrame#Header {{ background: {C.SURFACE}; border: none; border-bottom: 1px solid {C.BORDER}; }}
    QFrame#Card {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 10px; }}

    QLineEdit, QSpinBox {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 7px;
        padding: 5px 10px; color: {C.TEXT}; selection-background-color: {C.ACCENT_LINE};
        selection-color: {C.TEXT};
    }}
    QLineEdit:hover, QSpinBox:hover {{ border-color: {C.BORDER_HI}; }}
    QLineEdit:focus, QSpinBox:focus {{ border: 1px solid {C.ACCENT}; }}
    QSpinBox::up-button, QSpinBox::down-button {{
        subcontrol-origin: border; width: 18px; border: none; border-left: 1px solid {C.BORDER}; background: {C.BG};
    }}
    QSpinBox::up-button {{ subcontrol-position: top right; border-top-right-radius: 7px; }}
    QSpinBox::down-button {{ subcontrol-position: bottom right; border-bottom-right-radius: 7px;
        border-top: 1px solid {C.BORDER}; }}
    QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {C.HOVER}; }}
    QSpinBox::up-arrow {{ image: url(__UP__); width: 8px; height: 8px; }}
    QSpinBox::down-arrow {{ image: url(__DOWN__); width: 8px; height: 8px; }}

    QTextEdit#editor {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 10px;
        padding: 26px 34px; color: {C.BODY};
        selection-background-color: {C.ACCENT_LINE}; selection-color: {C.TEXT};
    }}
    QTextEdit#editor:focus {{ border: 1px solid {C.FOCUS}; }}

    QFrame#ChatBar {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 12px; }}
    QFrame#ChatBar[focused="true"] {{ border: 1px solid {C.ACCENT}; }}
    QFrame#ChatBar QTextEdit {{ background: transparent; border: none; padding: 2px 0; color: {C.TEXT}; }}

    QPushButton {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 7px;
        padding: 6px 14px; color: {C.TEXT}; font-weight: 600;
    }}
    QPushButton:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton:pressed {{ background: {C.BORDER}; }}
    QPushButton:disabled {{ color: {C.FAINT}; background: {C.SURFACE}; border-color: {C.BORDER}; }}
    QPushButton#Ghost {{ background: transparent; border: 1px solid transparent; color: {C.MUTED}; padding: 6px 10px; }}
    QPushButton#Ghost:hover {{ background: {C.HOVER}; color: {C.TEXT}; }}
    QPushButton#Chip {{
        background: {C.BG}; border: 1px solid {C.BORDER}; border-radius: 15px;
        padding: 4px 12px; color: {C.TEXT}; font-size: 10pt; font-weight: 600; text-align: left;
    }}
    QPushButton#Chip:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton#IconButton {{ padding: 0; }}
    QPushButton#DraftButton {{
        background: {C.ACCENT_SOFT}; border: 1px solid {C.ACCENT_LINE}; color: {C.ACCENT}; font-weight: 700;
    }}
    QPushButton#DraftButton:hover {{ background: #E3E6FC; border-color: {C.FOCUS}; }}
    QPushButton#DraftButton:disabled {{ background: {C.SURFACE}; border-color: {C.BORDER}; color: {C.FAINT}; }}
    QPushButton#SaveButton, QPushButton#Primary {{
        background: {C.ACCENT}; border: 1px solid {C.ACCENT}; color: white; font-weight: 700;
    }}
    QPushButton#SaveButton:hover, QPushButton#Primary:hover {{ background: {C.ACCENT_HI}; border-color: {C.ACCENT_HI}; }}
    QPushButton#SaveButton:disabled, QPushButton#Primary:disabled {{
        background: {C.ACCENT_LINE}; border-color: {C.ACCENT_LINE}; color: white; }}
    QPushButton#AIPrimaryButton {{ background: {C.ACCENT}; border: none; border-radius: 9px; padding: 0; }}
    QPushButton#AIPrimaryButton:hover {{ background: {C.ACCENT_HI}; }}
    QPushButton#AIPrimaryButton:disabled {{ background: {C.BORDER}; }}

    QCheckBox {{ color: {C.TEXT}; font-size: 10pt; spacing: 6px; background: transparent; }}

    QFrame#Panel {{ background: {C.SURFACE}; border: none; border-right: 1px solid {C.BORDER}; }}
    QWidget#PanelInner {{ background: {C.SURFACE}; }}
    QLabel#SectionTitle {{ color: {C.TEXT}; font-size: 11pt; font-weight: 700; }}
    QLabel#Thumb {{ background: {C.BG}; border: 1px solid {C.BORDER}; border-radius: 6px; }}
    QLabel#Status {{ color: {C.MUTED}; font-size: 9pt; }}
    QLabel#Status[state="error"] {{ color: #DC2626; }}
    QLabel#Status[state="done"] {{ color: {C.TEXT}; }}
    QPlainTextEdit, QDateEdit, QTimeEdit {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 7px; padding: 4px 8px;
        color: {C.TEXT}; selection-background-color: {C.ACCENT_LINE}; selection-color: {C.TEXT};
    }}
    QPlainTextEdit:focus, QDateEdit:focus, QTimeEdit:focus {{ border: 1px solid {C.ACCENT}; }}
    QLineEdit:disabled, QPlainTextEdit:disabled, QDateEdit:disabled, QTimeEdit:disabled {{
        background: {C.BG}; color: {C.FAINT}; }}
    QWidget#PanelFooter {{ background: {C.SURFACE}; border-top: 1px solid {C.BORDER}; }}
    QCheckBox#FieldCheck {{ color: {C.MUTED}; font-weight: 600; }}
    QCheckBox::indicator {{ width: 15px; height: 15px; border: 1px solid {C.BORDER_HI}; border-radius: 4px;
        background: {C.SURFACE}; }}
    QCheckBox::indicator:checked {{ background: {C.ACCENT}; border-color: {C.ACCENT}; image: url(__CHECK__); }}
    QCheckBox::indicator:disabled {{ background: {C.BG}; border-color: {C.BORDER}; }}
    QDateEdit::drop-down {{ subcontrol-origin: padding; subcontrol-position: center right; width: 20px; border: none; }}
    QDateEdit::down-arrow {{ image: url(__DOWN__); width: 8px; height: 8px; }}
    QScrollBar:vertical {{ background: transparent; width: 8px; margin: 2px; }}
    QScrollBar::handle:vertical {{ background: {C.BORDER_HI}; border-radius: 3px; min-height: 30px; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    QSplitter::handle {{ background: {C.BORDER}; }}
    QScrollArea {{ background: transparent; }}
    QListWidget {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px; padding: 4px; outline: 0;
    }}
    QListWidget::item {{ padding: 7px 10px; border-radius: 6px; color: {C.TEXT}; }}
    QListWidget::item:hover {{ background: {C.HOVER}; }}
    QListWidget::item:selected {{ background: {C.ACCENT_SOFT}; color: {C.ACCENT}; }}

    QMenu {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px; padding: 6px; }}
    QMenu::item {{ padding: 7px 16px 7px 12px; border-radius: 5px; color: {C.TEXT}; }}
    QMenu::item:selected {{ background: {C.HOVER}; }}
    QMenu::item:disabled {{ color: {C.FAINT}; }}
    QMenu::indicator {{ width: 0; height: 0; }}
    QMenu::separator {{ height: 1px; background: {C.BORDER}; margin: 5px 6px; }}

    QStatusBar {{ background: {C.BG}; color: {C.MUTED}; font-size: 9pt; border-top: 1px solid {C.BORDER}; }}
    QStatusBar::item {{ border: none; }}
    QToolTip {{ background: {C.TEXT}; color: white; border: none; padding: 4px 8px; }}
"""


def _style() -> str:
    """QApplication이 생긴 뒤에 만든다 (화살표 그림을 그리려면 필요)."""
    return (_STYLE_TEMPLATE.replace('__UP__', _arrow_png('spin_up', 'fa5s.chevron-up'))
            .replace('__DOWN__', _arrow_png('spin_down', 'fa5s.chevron-down'))
            .replace('__CHECK__', _arrow_png('check_white', 'fa5s.check', 'white')))

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


# ── 공문 규정 자동 보정 ──────────────────────────────────────────

def _correct_document_format(body: str) -> str:
    """공문 규정에 따라 본문 자동 보정 (날짜·들여쓰기·끝.·빈필드)."""
    if not body.strip():
        return body

    # 1. 날짜 마침표 보정: "2026. 4. 28" → "2026. 4. 28."
    body = re.sub(
        r'(?<!\d)(\d{4}\.\s*\d{1,2}\.\s*\d{1,2})(?![.\d])',
        r'\1.',
        body
    )

    # 2. 들여쓰기 보정: 가/나/다/라/마/바 앞 공백 2칸 강제
    lines = body.split('\n')
    corrected = []
    for line in lines:
        # 가-바 항목: 앞에 공백 2칸
        if re.match(r'^ {0,1}[가나라다마바]\.\s', line):
            line = '  ' + line.lstrip()
        # 1)2)3) 항목: 앞에 공백 4칸
        elif re.match(r'^ {0,3}\d\)\s', line):
            line = '    ' + line.lstrip()
        corrected.append(line)
    body = '\n'.join(corrected)

    # 3. 빈 필드 마커: ": \n" 또는 ": $" → ": [미입력]"
    body = re.sub(r':\s*\n', ': [미입력]\n', body)
    body = re.sub(r':\s*$', ': [미입력]', body)

    # 4. 끝. 보장: 마지막이 "끝."으로 끝나지 않으면 추가
    stripped = body.rstrip()
    if not stripped.endswith('끝.'):
        body = stripped + '\n\n끝.'

    return body


def _profile_block() -> str:
    try:
        from school_profile import prompt_block
        block = prompt_block()
        return block + "\n" if block else ''
    except Exception:
        return ''


class PlainPasteDocEdit(QTextEdit):
    def insertFromMimeData(self, source):
        if not source.hasText():
            super().insertFromMimeData(source)
            return
        cursor = self.textCursor()
        start = cursor.position()
        cursor.insertText(source.text())
        cursor.setPosition(start, QTextCursor.KeepAnchor)
        char_fmt = QTextCharFormat()
        char_fmt.setFont(self.font())
        cursor.mergeCharFormat(char_fmt)
        cursor.clearSelection()
        self.setTextCursor(cursor)


class ApiKeyDialog(QDialog):
    """API 키 입력 + 저장 여부 체크박스 다이얼로그."""
    def __init__(self, title: str, label: str, current_key: str = '', parent=None, hint: str = ''):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setStyleSheet(_style())
        self.setFixedWidth(px(460))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(8)

        head = QLabel(title)
        head.setObjectName('DialogTitle')
        layout.addWidget(head)
        lbl = QLabel(label)
        lbl.setWordWrap(True)
        layout.addWidget(lbl)
        layout.addSpacing(2)

        row = QHBoxLayout()
        row.setSpacing(6)
        self._key_edit = QLineEdit()
        self._key_edit.setPlaceholderText('API 키를 붙여 넣으세요')
        self._key_edit.setText(current_key)
        self._key_edit.setEchoMode(QLineEdit.Password)
        self._key_edit.setMinimumHeight(36)
        row.addWidget(self._key_edit, 1)
        self._btn_eye = QPushButton()
        self._btn_eye.setObjectName('IconButton')
        self._btn_eye.setFixedSize(36, 36)
        self._btn_eye.setIcon(qta.icon('fa5s.eye', color=C.MUTED))
        self._btn_eye.setToolTip('키 보기')
        self._btn_eye.setCheckable(True)
        self._btn_eye.toggled.connect(lambda on: self._key_edit.setEchoMode(
            QLineEdit.Normal if on else QLineEdit.Password))
        row.addWidget(self._btn_eye)
        layout.addLayout(row)

        if hint:
            h = QLabel(hint)
            h.setObjectName('Hint')
            h.setWordWrap(True)
            h.setTextInteractionFlags(Qt.TextSelectableByMouse)
            layout.addWidget(h)

        self._save_cb = QCheckBox('이 PC에 키 저장 (Windows 암호화)')
        self._save_cb.setChecked(True)
        layout.addWidget(self._save_cb)
        layout.addSpacing(6)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton('다음')
        btn_ok.setObjectName('Primary')
        btn_ok.setMinimumHeight(34)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)
        keep_light(self)

    def key(self) -> str:
        return self._key_edit.text().strip()

    def should_save(self) -> bool:
        return self._save_cb.isChecked()


# AI 제공자별 설정 정보
_PROVIDERS = {
    'opencode': {'desc': 'Go·Zen 요금제'},
    'claude':   {'desc': 'Anthropic API 키',
                 'key_label': 'Anthropic(Claude) API 키를 입력하세요.',
                 'hint': '키 발급: console.anthropic.com → API Keys (사용한 만큼 결제)',
                 'load_key': load_claude_key, 'save_key': save_claude_key,
                 'load_model': load_claude_model, 'save_model': save_claude_model,
                 'fetch': fetch_claude_models, 'test': test_claude},
    'openai':   {'desc': 'OpenAI API 키',
                 'key_label': 'OpenAI(ChatGPT) API 키를 입력하세요.',
                 'hint': '키 발급: platform.openai.com → API keys (ChatGPT 구독과 별도, 사용한 만큼 결제)',
                 'load_key': load_openai_key, 'save_key': save_openai_key,
                 'load_model': load_openai_model, 'save_model': save_openai_model,
                 'fetch': fetch_openai_models, 'test': test_openai},
}


class DocumentEditorWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ai_thread = None
        self._capturing = False
        DocumentEditorWindow._instance = self
        self.init_ui()
        keep_light(self)  # 공문 작성 창은 아직 라이트 전용 (다크 테마에서도 원래 모습 유지)

    def init_ui(self):
        self.setWindowTitle("공문 작성")
        self._restore_window_state()  # 저장된 위치/크기 복원, 없으면 기본값

        _input_h = 36

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(0)
        root.setContentsMargins(0, 0, 0, 0)

        # ── 머리말: 창 이름 · AI 상태(누르면 AI 설정) · 도움말 ──
        header = QFrame()
        header.setObjectName('Header')
        hl = QHBoxLayout(header)
        hl.setContentsMargins(20, 10, 14, 10)
        hl.setSpacing(6)
        icon = QLabel()
        icon.setPixmap(qta.icon('fa5s.file-signature', color=C.ACCENT).pixmap(18, 18))
        hl.addWidget(icon)
        title = QLabel('공문 작성')
        title.setObjectName('WindowTitle')
        hl.addWidget(title)
        hl.addStretch()
        self.btn_panel = QPushButton('받은 공문')
        self.btn_panel.setObjectName('Ghost')
        self.btn_panel.setIcon(qta.icon('fa5s.angle-double-right', color=C.MUTED))
        self.btn_panel.setToolTip('받은 공문 칸 펼치기')
        self.btn_panel.setCursor(Qt.PointingHandCursor)
        self.btn_panel.clicked.connect(lambda: self._show_panel(True))
        self.btn_panel.hide()
        hl.addWidget(self.btn_panel)
        self.btn_capture = QPushButton('공문 캡처')
        self.btn_capture.setIcon(qta.icon('fa5s.crop-alt', color=C.TEXT))
        self.btn_capture.setIconSize(QSize(13, 13))
        self.btn_capture.setMinimumHeight(30)
        self.btn_capture.setCursor(Qt.PointingHandCursor)
        self.btn_capture.setToolTip('받은 공문을 화면에서 잘라 AI로 읽기 (Ctrl+Shift+X)')
        self.btn_capture.clicked.connect(lambda: self._start_capture())
        hl.addWidget(self.btn_capture)
        hl.addSpacing(4)
        self.btn_apikey = QPushButton("AI 설정")
        self.btn_apikey.setObjectName('Chip')
        self.btn_apikey.setCursor(Qt.PointingHandCursor)
        self.btn_apikey.setMinimumHeight(30)
        self.btn_apikey.setToolTip('AI 연결 바꾸기 (OpenCode · Claude · ChatGPT)')
        self.btn_apikey.clicked.connect(self._show_ai_settings_menu)
        hl.addWidget(self.btn_apikey)
        self.btn_help = QPushButton()
        self.btn_help.setObjectName('Ghost')
        self.btn_help.setIcon(qta.icon('fa5s.question-circle', color=C.MUTED))
        self.btn_help.setIconSize(QSize(16, 16))
        self.btn_help.setFixedSize(32, 30)
        self.btn_help.setToolTip('사용 방법')
        self.btn_help.setCursor(Qt.PointingHandCursor)
        self.btn_help.clicked.connect(self._show_help)
        hl.addWidget(self.btn_help)
        root.addWidget(header)

        # ── 왼쪽: 받은 공문 칸 (캡처하면 나타남) | 오른쪽: 공문 작성 ──
        from PyQt5.QtWidgets import QSplitter
        from ui.capture_panel import CapturePanel
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(1)
        self.panel = CapturePanel()
        self.panel.setMinimumWidth(260)
        self.panel.hide()
        self.panel.register_requested.connect(self._register_from_panel)
        self.panel.reply_requested.connect(self.start_reply)
        self.panel.use_title.connect(lambda t: t and self.title_input.setText(t))
        self.panel.use_ref.connect(lambda t: t and self.ref_input.setText(t))
        self.panel.collapse_requested.connect(lambda: self._show_panel(False))
        self.panel.message.connect(self.statusBar().showMessage)
        self.panel.analyzed.connect(self._fill_from_capture)
        self.splitter.addWidget(self.panel)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 16, 20, 14)
        self.splitter.addWidget(body)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        root.addWidget(self.splitter, 1)
        self.register_task_callback = None   # 메모 창이 넣어 줌: (name, deadline, notes, related_no)

        # ── 입력 카드: 제목 · 관련 번호 · 붙임 · 자동 초안 ──
        card = QFrame()
        card.setObjectName('Card')
        grid = QGridLayout(card)
        grid.setContentsMargins(16, 14, 16, 14)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        def field_label(text):
            lb = QLabel(text)
            lb.setObjectName('FieldLabel')
            lb.ensurePolished()     # 글자 크기(QSS)를 먼저 적용해야 폭 계산이 맞음
            lb.setFixedWidth(lb.fontMetrics().horizontalAdvance('관련 번호') + 6)   # 글자 길이 기준
            return lb

        self.title_input = QLineEdit()
        self.title_input.setFixedHeight(_input_h)
        self.title_input.setPlaceholderText("공문 제목  (예: 2026학년도 교무회의 운영 계획)")
        self.title_input.textChanged.connect(self._update_draft_button)
        self.title_input.returnPressed.connect(self._try_generate_draft)
        grid.addWidget(field_label('제목'), 0, 0)
        grid.addWidget(self.title_input, 0, 1, 1, 5)

        self.ref_input = QLineEdit()
        self.ref_input.setFixedHeight(_input_h)
        self.ref_input.setPlaceholderText("관련 공문번호 (선택)")
        self.ref_input.textChanged.connect(self._update_draft_button)
        self.ref_input.returnPressed.connect(self._try_generate_draft)
        self.btn_move_to_title = QPushButton()
        self.btn_move_to_title.setObjectName('IconButton')
        self.btn_move_to_title.setFixedSize(_input_h, _input_h)
        self.btn_move_to_title.setIcon(qta.icon('fa5s.arrow-up', color=C.MUTED))
        self.btn_move_to_title.setIconSize(QSize(13, 13))
        self.btn_move_to_title.setToolTip("관련 번호 칸의 글을 제목으로 올리기")
        self.btn_move_to_title.clicked.connect(self._move_ref_to_title)
        lbl_attach = QLabel("붙임")
        lbl_attach.setObjectName('FieldLabel')
        self.spin_attach = QSpinBox()
        self.spin_attach.setRange(0, 9)
        self.spin_attach.setValue(1)
        self.spin_attach.setSuffix(' 개')
        self.spin_attach.setFixedHeight(_input_h)
        self.spin_attach.ensurePolished()
        self.spin_attach.setFixedWidth(self.spin_attach.fontMetrics().horizontalAdvance('0 개') + 52)
        self.spin_attach.setToolTip("붙임 항목 수 (0 = 붙임 없음)")
        self.spin_attach.lineEdit().returnPressed.connect(self._try_generate_draft)
        self.btn_draft = QPushButton("자동 초안")
        self.btn_draft.setObjectName("DraftButton")
        self.btn_draft.setIcon(qta.icon('fa5s.magic', color=C.ACCENT, color_disabled=C.FAINT))
        self.btn_draft.setIconSize(QSize(13, 13))
        self.btn_draft.setFixedHeight(_input_h)
        self.btn_draft.setToolTip('제목에 맞는 공문 뼈대를 넣고, AI가 켜져 있으면 내용까지 채웁니다 (Enter)')
        self.btn_draft.clicked.connect(self.generate_draft)
        grid.addWidget(field_label('관련 번호'), 1, 0)
        grid.addWidget(self.ref_input, 1, 1)
        grid.addWidget(self.btn_move_to_title, 1, 2)
        grid.addWidget(lbl_attach, 1, 3)
        grid.addWidget(self.spin_attach, 1, 4)
        grid.addWidget(self.btn_draft, 1, 5)
        grid.setColumnStretch(1, 1)
        layout.addWidget(card)

        # ── 본문 (종이) ──
        self.editor = PlainPasteDocEdit()
        self.editor.setObjectName("editor")
        _editor_font = QFont('굴림체', 12)
        _editor_font.setStretch(100)
        _editor_font.setLetterSpacing(QFont.PercentageSpacing, 100.0)
        self.editor.setFont(_editor_font)
        self.editor.document().blockCountChanged.connect(self._apply_editor_line_height)
        self._apply_editor_line_height()
        self.editor.setPlaceholderText(
            "제목을 넣고 [자동 초안]을 누르거나, 아래에 AI에게 부탁할 내용을 적어 보세요."
        )
        layout.addWidget(self.editor, stretch=1)

        # ── AI 요청 칸 ──
        self.ai_frame = QFrame()
        self.ai_frame.setObjectName("ChatBar")
        ai_frame_layout = QHBoxLayout(self.ai_frame)
        ai_frame_layout.setContentsMargins(14, 6, 6, 6)
        ai_frame_layout.setSpacing(8)
        spark = QLabel()
        spark.setPixmap(qta.icon('fa5s.comment-dots', color=C.FAINT).pixmap(16, 16))
        spark.setAlignment(Qt.AlignTop)
        spark.setContentsMargins(0, 8, 0, 0)
        ai_frame_layout.addWidget(spark)
        self.ai_input = QTextEdit()
        self.ai_input.setPlaceholderText("AI에게 요청")
        self.ai_input.setFixedHeight(52)
        self.ai_input.setAcceptRichText(False)
        self.ai_input.installEventFilter(self)
        self.editor.installEventFilter(self)
        self.btn_send = QPushButton()
        self.btn_send.setObjectName("AIPrimaryButton")
        self.btn_send.setIcon(qta.icon('fa5s.arrow-up', color='white'))
        self.btn_send.setIconSize(QSize(15, 15))
        self.btn_send.setFixedSize(38, 38)
        self.btn_send.setCursor(Qt.PointingHandCursor)
        self.btn_send.setToolTip("보내기 (Enter · 줄바꿈은 Shift+Enter)")
        self.btn_send.clicked.connect(self.ask_ai)
        ai_frame_layout.addWidget(self.ai_input, 1)
        ai_frame_layout.addWidget(self.btn_send, 0, Qt.AlignBottom)
        layout.addWidget(self.ai_frame)

        # ── 아래 버튼 ──
        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(8)
        self.btn_copy_title = QPushButton("제목 복사")
        self.btn_copy_title.setIcon(qta.icon('fa5.copy', color=C.MUTED))
        self.btn_copy_title.setMinimumHeight(_input_h)
        self.btn_copy_title.clicked.connect(self._copy_title)
        self.btn_copy = QPushButton("본문 복사")
        self.btn_copy.setIcon(qta.icon('fa5.copy', color=C.MUTED))
        self.btn_copy.setMinimumHeight(_input_h)
        self.btn_copy.clicked.connect(self._copy_body)
        self.btn_save_exit = QPushButton("저장 후 나가기")
        self.btn_save_exit.setObjectName("SaveButton")
        self.btn_save_exit.setMinimumHeight(_input_h)
        self.btn_save_exit.setMinimumWidth(130)
        self.btn_save_exit.clicked.connect(self.save_and_close)
        ctrl_row.addWidget(self.btn_copy_title)
        ctrl_row.addWidget(self.btn_copy)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self.btn_save_exit)
        layout.addLayout(ctrl_row)

        self.setStyleSheet(_style())
        self.statusBar().setSizeGripEnabled(False)
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

        # RAG: 유사 공문 검색하여 컨텍스트 추가
        rag_context = ''
        try:
            from rag import search_similar, make_search_prompt
            embeddings = get_all_embeddings()
            if embeddings:
                similar = search_similar(title, embeddings, top_k=2)
                if similar:
                    rag_context = make_search_prompt(similar)
        except Exception:
            pass  # RAG 실패해도 진행에는 문제 없음

        prompt = (
            f"제목: {title}\n{ref_line}\n"
            f"{rag_context}\n"
            f"{self._received_context()}"
            f"{_profile_block()}"
            '반드시 JSON 형식({"body": "공문 본문"})으로만 출력하세요.'
        )
        self._start_ai(prompt, system=system, structured=True)

    # ── 회신 공문 (받은 공문 칸 → 회신 공문 쓰기) ──
    def start_reply(self, analysis: dict, instruction: str = ''):
        """캡처·분석한 공문에 대한 회신 공문: 제목·관련 칸을 채우고 AI 초안을 만든다."""
        import doc_ai
        title = doc_ai.reply_title(analysis)
        ref = doc_ai.related_line(analysis)
        self.title_input.setText(title)
        self.ref_input.setText(ref)
        if load_ai_mode() == 'none':
            self.editor.setPlainText(_get_template(title, 'submit', ref, self.spin_attach.value()))
            self._apply_editor_line_height()
            self.statusBar().showMessage('회신 공문 뼈대 삽입 완료 — AI를 켜면 본문까지 작성합니다.')
            return
        if self.ai_thread and self.ai_thread.isRunning():
            self.statusBar().showMessage('AI가 작업 중입니다. 끝난 뒤 다시 시도하세요.')
            return
        prompt = doc_ai.build_reply_prompt(analysis, _profile_block(), instruction)
        self._start_ai(prompt, system=doc_ai.REPLY_SYSTEM, structured=True)
        self.statusBar().showMessage('AI가 회신 공문을 작성 중...')

    # ── AI 입력창 Enter 키 → 전송 ──
    def eventFilter(self, obj, event):
        if obj is self.ai_input and event.type() in (QEvent.FocusIn, QEvent.FocusOut):
            self.ai_frame.setProperty('focused', event.type() == QEvent.FocusIn)   # 요청 칸 테두리 강조
            self.ai_frame.style().unpolish(self.ai_frame)
            self.ai_frame.style().polish(self.ai_frame)
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
        ctx = self._received_context()
        if ctx:
            prompt = f"{ctx}\n{prompt}"

        self._start_ai(prompt)

    # ── 공통 AI 실행 ──
    # AI가 쓰는 동안 원래 내용을 보관한다. 오류·빈 답·취소면 원래 내용으로 되돌리고,
    # 성공하면 결과를 '되돌리기(Ctrl+Z) 한 번'으로 원래 내용에 돌아갈 수 있게 넣는다.
    def _start_ai(self, prompt, system=None, structured=False):
        if self.ai_thread and self.ai_thread.isRunning():
            return

        self.statusBar().showMessage("AI가 생각 중..." if not structured else "AI가 공문을 작성 중...")
        self._set_controls_enabled(False)

        self._ai_gen = getattr(self, '_ai_gen', 0) + 1          # 이번 요청 번호 (늦게 온 이전 응답 무시용)
        gen = self._ai_gen
        self._before_ai = self.editor.toPlainText()
        self.editor.setUndoRedoEnabled(False)                     # 흘러들어오는 글자는 되돌리기 기록에 안 남김
        self.editor.clear()

        self.ai_thread = AiStreamThread(prompt, system=system, structured=structured)
        self.ai_thread.new_text_signal.connect(lambda t, g=gen: self._append_text(t, g))
        self.ai_thread.finished_signal.connect(lambda t, g=gen: self._on_finished(t, g))
        self.ai_thread.error_signal.connect(lambda m, g=gen: self._on_error(m, g))
        self.ai_thread.start()

    def _is_current_ai(self, gen):
        return gen is None or gen == getattr(self, '_ai_gen', 0)

    def _append_text(self, text, gen=None):
        if not self._is_current_ai(gen):
            return
        cursor = self.editor.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(text)
        self.editor.setTextCursor(cursor)

    def _restore_before_ai(self):
        """AI가 쓰기 전 내용으로 되돌린다."""
        self.editor.blockSignals(True)
        self.editor.setPlainText(getattr(self, '_before_ai', ''))
        self.editor.blockSignals(False)
        self._apply_editor_line_height()
        self.editor.setUndoRedoEnabled(True)

    def _on_finished(self, _, gen=None):
        if not self._is_current_ai(gen):
            return
        self._ai_gen = getattr(self, '_ai_gen', 0) + 1          # 이 요청은 끝 — 이후 신호 무시
        raw = self.editor.toPlainText()
        if not raw.strip():
            self._restore_before_ai()
            self.statusBar().showMessage("AI가 빈 답을 보냈습니다 — 원래 내용을 그대로 두었습니다. 다시 시도해 주세요.")
            self._set_controls_enabled(True)
            return
        # AI 응답 후 공문 규정 자동 보정
        corrected = _correct_document_format(raw)
        # 원래 내용 → (한 번의 편집으로) 새 내용: Ctrl+Z 한 번이면 원래 내용으로 돌아감
        self._restore_before_ai()
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.Document)
        cursor.insertText(corrected)
        cursor.select(QTextCursor.Document)                       # 줄간격도 같은 편집 안에서 (되돌리기 1번)
        from PyQt5.QtGui import QTextBlockFormat
        fmt = QTextBlockFormat()
        fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
        cursor.mergeBlockFormat(fmt)
        cursor.endEditBlock()
        cursor.movePosition(QTextCursor.End)
        self.editor.setTextCursor(cursor)
        masked = getattr(self.ai_thread, 'masked_counts', None) or {}
        if masked:
            from privacy import describe
            self.statusBar().showMessage(f"완료 — 🛡 개인정보 {describe(masked)}을 가리고 보냈습니다. (Ctrl+Z: 이전 내용)")
        else:
            self.statusBar().showMessage("완료 — Ctrl+Z를 누르면 이전 내용으로 돌아갑니다.")
        self._set_controls_enabled(True)
        self.ai_input.clear()

    def _on_error(self, msg, gen=None):
        if not self._is_current_ai(gen):
            return
        self._ai_gen = getattr(self, '_ai_gen', 0) + 1
        self._restore_before_ai()
        self.statusBar().showMessage(f"오류: {msg} — 원래 내용으로 되돌렸습니다.")
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
            if self.ai_thread and self.ai_thread.isRunning():
                self._cancel_ai()          # AI가 쓰는 중이면 먼저 멈추기만
            elif self._confirm_discard():
                self.close()
        elif event.key() == Qt.Key_X and event.modifiers() == (Qt.ControlModifier | Qt.ShiftModifier):
            self._start_capture()
        else:
            super().keyPressEvent(event)

    def _show_help(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("사용 방법")
        dlg.setWindowFlags(dlg.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        dlg.setStyleSheet(_style())
        dlg.setFixedWidth(px(500))

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(24, 22, 24, 18)
        layout.setSpacing(12)
        head = QLabel('공문 작성 사용 방법')
        head.setObjectName('DialogTitle')
        layout.addWidget(head)

        items = [
            ('fa5s.magic', '자동 초안',
             '제목(또는 관련 번호)을 넣고 Enter나 [자동 초안]을 누르면 공문 뼈대가 들어가고, '
             'AI가 연결되어 있으면 내용까지 채웁니다.'),
            ('fa5s.comment-dots', 'AI에게 요청',
             '아래 칸에 "장소를 시청각실로 바꿔줘"처럼 적고 Enter. 줄바꿈은 Shift+Enter. '
             '결과가 마음에 안 들면 Ctrl+Z 한 번으로 이전 내용으로 돌아갑니다.'),
            ('fa5s.crop-alt', '공문 캡처',
             'Ctrl+Shift+X나 [공문 캡처]로 받은 공문을 잘라 오면 왼쪽 칸에 공문번호·기한·요약·할 일이 나옵니다. '
             '번호가 안 잡혔으면 그 부분만 한 번 더 캡처하세요 — 같은 공문에 합쳐서 다시 읽습니다. '
             '다른 공문은 왼쪽 [새 공문]. 업무로 등록하거나 회신 공문을 쓰고, AI 대화도 그 공문을 참고합니다.'),
            ('fa5s.stop-circle', '취소', 'AI가 쓰는 중에 Esc를 누르면 멈추고 원래 내용으로 되돌립니다.'),
            ('fa5s.shield-alt', '개인정보 보호',
             '보호 모드가 켜져 있으면 AI로 보내기 전에 주민번호·전화번호·이름 등을 가립니다.'),
        ]
        for ic, name, desc in items:
            row = QHBoxLayout()
            row.setSpacing(12)
            i = QLabel()
            i.setPixmap(qta.icon(ic, color=C.ACCENT).pixmap(16, 16))
            i.setFixedWidth(20)
            i.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
            i.setContentsMargins(0, 3, 0, 0)
            row.addWidget(i)
            t = QLabel(f"<span style='color:{C.TEXT}; font-weight:700;'>{name}</span><br>"
                       f"<span style='color:{C.MUTED};'>{desc}</span>")
            t.setWordWrap(True)
            t.setStyleSheet('font-size: 10pt;')
            t.ensurePolished()                              # 글자 크기(10pt)를 먼저 적용해야 높이 계산이 맞음
            t.setFixedWidth(px(420))
            t.setMinimumHeight(t.heightForWidth(px(420)))      # 줄바꿈한 높이만큼 (잘림 방지)
            row.addWidget(t, 1)
            layout.addLayout(row)

        layout.addSpacing(4)
        btn_ok = QPushButton("확인")
        btn_ok.setObjectName('Primary')
        btn_ok.setMinimumSize(84, 34)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)
        keep_light(dlg)
        dlg.exec_()

    def _cancel_ai(self):
        if self.ai_thread and self.ai_thread.isRunning():
            self.ai_thread.stop()
            self._ai_gen = getattr(self, '_ai_gen', 0) + 1      # 취소 뒤에 도착하는 글자는 무시
            self.ai_thread.wait(500)
            self._restore_before_ai()
            self.statusBar().showMessage('AI 생성을 취소하고 원래 내용으로 되돌렸습니다.')
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
        """현재 AI 모드에 따라 AI 상태 표시와 AI 관련 버튼을 맞춘다."""
        mode = load_ai_mode()
        ai_available = mode != 'none'
        self._update_draft_button()
        self.btn_send.setEnabled(ai_available)
        self.ai_input.setEnabled(ai_available)
        if ai_available:
            model = load_external_model_name()
            self.btn_apikey.setText(f'{PROVIDER_LABELS[mode]} · {model}')
            self.btn_apikey.setIcon(qta.icon('fa5s.circle', color=C.OK))
            self.ai_input.setPlaceholderText('AI에게 요청  (예: 장소를 시청각실로 바꾸고 일시는 11월 10일로 해줘)')
            return
        self.btn_apikey.setIcon(qta.icon('fa5s.circle', color=C.OFF))
        legacy = legacy_ai_name()
        if legacy:
            self.btn_apikey.setText('AI 다시 연결 필요')
            self.btn_apikey.setIcon(qta.icon('fa5s.exclamation-circle', color=C.WARN))
            self.ai_input.setPlaceholderText(
                f'이전에 쓰던 {legacy}는 더 이상 지원하지 않아요 — 위의 [AI 다시 연결 필요]를 눌러 '
                'OpenCode·Claude·ChatGPT 중 하나를 연결하세요.')
        else:
            self.btn_apikey.setText('AI 꺼짐 · 연결하기')
            self.ai_input.setPlaceholderText('AI가 꺼져 있어요 — 위의 [AI 꺼짐 · 연결하기]에서 연결하면 쓸 수 있어요.')
        self.btn_apikey.setIconSize(QSize(9, 9) if not legacy else QSize(12, 12))

    # ── AI 설정 메뉴 ──
    def _show_ai_settings_menu(self):
        from PyQt5.QtWidgets import QMenu
        mode = load_ai_mode()
        menu = QMenu(self)
        menu.setStyleSheet(_style())
        title = menu.addAction('AI 연결')
        title.setEnabled(False)
        acts = {}
        for key in ('opencode', 'claude', 'openai'):
            mark = '✓  ' if mode == key else '     '
            a = menu.addAction(f'{mark}{PROVIDER_LABELS[key]}\t{_PROVIDERS[key]["desc"]}')
            acts[a] = key
        menu.addSeparator()
        act_none = menu.addAction(('✓  ' if mode == 'none' else '     ') + 'AI 끄기')

        action = menu.exec_(self.btn_apikey.mapToGlobal(self.btn_apikey.rect().bottomLeft()))
        if action is None:
            return
        if action == act_none:
            save_ai_mode('none')
            self._apply_ai_mode()
            self.statusBar().showMessage('AI를 껐습니다 — 자동 초안은 뼈대만 넣습니다.')
        elif acts.get(action) == 'opencode':
            self._setup_opencode_ai()
        elif acts.get(action) == 'claude':
            self._setup_claude_ai()
        elif acts.get(action) == 'openai':
            self._setup_openai_ai()

    def _busy(self, msg):
        self.statusBar().showMessage(msg)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

    def _idle(self):
        QApplication.restoreOverrideCursor()

    def _setup_keyed_ai(self, mode: str):
        """Claude·ChatGPT: API 키 → 모델(서버의 최신 목록) → 연결 시험 → 저장."""
        p = _PROVIDERS[mode]
        label = PROVIDER_LABELS[mode]
        dlg = ApiKeyDialog(f'{label} 연결', p['key_label'], p['load_key'](), self, hint=p['hint'])
        if dlg.exec_() != dlg.Accepted:
            return
        key = dlg.key()
        if not key:
            self.statusBar().showMessage(f'{label} API 키가 비어 있습니다.')
            return

        self._busy(f'{label} 모델 목록을 받는 중...')
        try:
            models = p['fetch'](key)
        finally:
            self._idle()
        dlg2 = _ModelSelectDialog(f'{label} 모델 선택', models, p['load_model'](), self,
                                  subtitle='목록 위쪽이 최신 모델입니다. 공문 작성에는 빠르고 저렴한 모델로 충분해요.')
        if dlg2.exec_() != dlg2.Accepted or not dlg2.selected_model():
            return
        model = dlg2.selected_model()

        self._busy(f'{label} · {model} 연결 확인 중...')
        try:
            ok, msg = p['test'](key, model)
        finally:
            self._idle()
        if not ok:
            self.statusBar().showMessage(f'{label} 연결 실패')
            QMessageBox.warning(self, f'{label} 연결 실패', msg)
            return
        p['save_key'](key, persist=dlg.should_save())
        p['save_model'](model)
        save_ai_mode(mode)
        self._apply_ai_mode()
        self.statusBar().showMessage(f'{label} · {model} 연결 완료.')

    def _setup_claude_ai(self):
        self._setup_keyed_ai('claude')

    def _setup_openai_ai(self):
        self._setup_keyed_ai('openai')

    def _setup_opencode_ai(self):
        """OpenCode: 요금제 → API 키 → 모델(서버의 최신 목록) → 연결 시험 → 저장."""
        plan_keys = list(OPENCODE_PLANS)
        labels = [OPENCODE_PLANS[k]['label'] for k in plan_keys]
        label, ok = QInputDialog.getItem(self, 'OpenCode 연결', '사용 중인 OpenCode 요금제를 고르세요:',
                                         labels, plan_keys.index(load_opencode_plan()), False)
        if not ok:
            return
        plan = plan_keys[labels.index(label)]

        dlg = ApiKeyDialog('OpenCode 연결', 'OpenCode API 키를 입력하세요.', load_opencode_key(), self,
                           hint='키 발급: opencode.ai 로그인 → 콘솔 → API Keys')
        if dlg.exec_() != dlg.Accepted:
            return
        key = dlg.key()
        if not key:
            self.statusBar().showMessage('OpenCode API 키가 비어 있습니다.')
            return

        self._busy('OpenCode 모델 목록을 받는 중...')
        try:
            models = fetch_opencode_models(plan)
        finally:
            self._idle()

        def mark(m):
            tags = []
            if m['id'] == 'claude-haiku-5-5':
                tags.append('추천')
            if m['image']:
                tags.append('🖼 이미지')
            if m['id'].startswith('deepseek'):
                tags.append('⚠ 지역 설정 필요')     # OpenCode 기본(유럽·미국)에서는 막혀 있음
            return m['id'] + ('   ' + ' · '.join(tags) if tags else '')
        shown = [mark(m) for m in models]
        current = next((mark(m) for m in models if m['id'] == load_opencode_model()), '')
        dlg2 = _ModelSelectDialog('OpenCode 모델 선택', shown, current, self,
                                  subtitle='🖼 표시가 있는 모델은 캡처한 공문 이미지도 읽을 수 있어요.')
        if dlg2.exec_() != dlg2.Accepted:
            return
        picked = dlg2.selected_model().split()[0] if dlg2.selected_model() else ''
        if not picked:
            return

        if any(m['image'] is not None for m in models):      # 서버 목록을 받았을 때만 (요청 방식·이미지 지원)
            save_opencode_model_info(models)
        self._busy(f'OpenCode · {picked} 연결 확인 중...')
        try:
            ok, msg = test_opencode(plan, key, picked)
        finally:
            self._idle()
        if not ok:
            self.statusBar().showMessage('OpenCode 연결 실패')
            QMessageBox.warning(self, 'OpenCode 연결 실패', msg)
            return

        save_opencode_key(key, persist=dlg.should_save())
        save_opencode_plan(plan)
        save_opencode_model(picked)
        save_ai_mode('opencode')
        self._apply_ai_mode()
        self.statusBar().showMessage(f'OpenCode · {picked} 연결 완료.')
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

        doc_id = save_official_document(title, doc_number, content, '')
        self.statusBar().showMessage('저장되었습니다.')
        QTimer.singleShot(1000, lambda: self.statusBar().showMessage(''))

        # RAG: 저장한 문서 자동 임베딩
        try:
            from rag import encode, embedding_to_bytes
            vec = encode(f"{title}\n{content}")
            save_embedding(doc_id, title, content, '', embedding_to_bytes(vec))
        except Exception:
            pass

        return True

    def _confirm_discard(self) -> bool:
        """Esc로 닫을 때: 본문에 쓴 내용이 있으면 한 번 묻는다 (Esc를 한 번 더 누르면 취소)."""
        if not self.editor.toPlainText().strip():
            return True
        box = QMessageBox(QMessageBox.Question, '공문 작성 닫기',
                          '작성 중인 본문이 있습니다. 저장하지 않고 닫을까요?', parent=self)
        b_close = box.addButton('닫기', QMessageBox.AcceptRole)
        box.addButton('계속 쓰기', QMessageBox.RejectRole)
        box.setDefaultButton(b_close)
        box.exec_()
        return box.clickedButton() is b_close

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
            self.setGeometry(*fit_to_screen(s['x'], s['y'], s['width'], s['height'], min_w=560, min_h=420))
        except Exception:
            self.setGeometry(*default_geometry(900, 680))   # 처음엔 화면 크기에 맞춘 크기로

    def closeEvent(self, event):
        import json, os
        os.makedirs(os.path.dirname(self._STATE_FILE), exist_ok=True)
        g = self.geometry()
        with open(self._STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'x': g.x(), 'y': g.y(),
                       'width': g.width() - (self._panel_extra if self.panel.isVisible() else 0),
                       'height': g.height()}, f)
        self.panel.shutdown()
        if DocumentEditorWindow._instance is self:
            DocumentEditorWindow._instance = None
        super().closeEvent(event)

    # ── 캡처 ──────────────────────────────────────────────────────
    def _start_capture(self, extra_hide=()):
        """공문 캡처 (Ctrl+Shift+X·[공문 캡처]). AI가 연결되어 있으면 왼쪽 「받은 공문」 칸에서 분석하고,
        AI가 꺼져 있으면 글자만 읽어 제목·관련 번호 칸에 넣는다. extra_hide: 캡처하는 동안 숨길 다른 창."""
        self._capturing = True
        self._capture_hidden = [w for w in extra_hide if w.isVisible()]
        for w in self._capture_hidden:
            w.hide()
        self.hide()
        QApplication.processEvents()
        QTimer.singleShot(250, self._show_capture_overlay)

    def _show_capture_overlay(self):
        from capture import grab_fullscreen, ScreenCaptureOverlay
        screenshot = grab_fullscreen()
        self._overlay = ScreenCaptureOverlay(screenshot)
        self._overlay.region_captured.connect(self._on_capture_complete)
        self._overlay.cancelled.connect(self._restore_after_capture)
        self._overlay.show()
        self._overlay.activateWindow()

    def _restore_after_capture(self):
        self._capturing = False
        for w in getattr(self, '_capture_hidden', []):
            w.show()
        self._capture_hidden = []
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
        if load_ai_mode() == 'none':
            self._ocr_fallback(pixmap, self._capture_target())
            return
        self._show_panel(True)
        self.panel.add_capture(pixmap)
        n = self.panel.capture_count()
        self.statusBar().showMessage(
            f'캡처 {n}장을 합쳐서 AI가 다시 읽는 중… (다른 공문은 왼쪽 [새 공문])' if n > 1
            else '받은 공문을 AI가 읽는 중… (번호가 안 잡혔으면 그 부분만 한 번 더 캡처하세요)')

    # ── 받은 공문 칸 ─────────────────────────────────────────────
    _instance = None          # 지금 열려 있는 공문 작성 창 (메모 창·단축키가 같은 창을 쓰도록)
    _panel_extra = 0          # 칸을 펼치면서 넓힌 창 폭

    def _show_panel(self, show: bool):
        if show == self.panel.isVisible():
            return
        if show:
            # 칸 폭은 화면 크기에 비례 (노트북처럼 좁은 화면에선 좁게)
            screen = QApplication.desktop().availableGeometry(self)
            need = self.panel.content_min_width()
            self.panel.setMinimumWidth(need)
            w = max(need, min(340, int(screen.width() * 0.22)))
            self.panel.show()
            self.splitter.setSizes([w, max(self.width(), 600)])
            # 칸만큼 창을 넓힌다. 오른쪽에 자리가 없으면 창을 왼쪽으로 옮긴다 (화면 밖으로 안 나가게)
            g = self.frameGeometry()
            grow = max(0, min(w, screen.width() - g.width()))
            shift = max(0, min(g.right() + grow - screen.right(), g.left() - screen.left()))
            self._panel_extra = grow
            if grow:
                if shift:
                    self.move(self.x() - shift, self.y())
                self.resize(self.width() + grow, self.height())
        else:
            self.panel.hide()
            if self._panel_extra:
                self.resize(self.width() - self._panel_extra, self.height())
            self._panel_extra = 0
        self.btn_panel.setVisible(not show and self.panel.has_content())

    def _fill_from_capture(self, title: str, doc_number: str):
        """캡처 분석 결과를 오른쪽 제목·관련 번호 칸에도 넣는다. 찾은 값만 넣고, 칸이 비어 있거나
        지난번에 캡처로 넣은 값 그대로일 때만 바꾼다 (사용자가 고친 제목·번호는 유지)."""
        for value, box, key in ((title, self.title_input, '_auto_title'), (doc_number, self.ref_input, '_auto_ref')):
            cur = box.text().strip()
            if value and (not cur or cur == getattr(self, key, None)):
                box.setText(value)
                box.setCursorPosition(0)          # 긴 번호도 앞부분부터 보이게
                setattr(self, key, value)

    def _received_context(self) -> str:
        """AI 대화·자동 초안에 함께 보낼 받은 공문 정리 (칸의 [참고] 체크가 켜져 있을 때)."""
        if not self.panel.use_as_context():
            return ''
        import doc_ai
        return doc_ai.context_block(self.panel.collect())

    def _register_from_panel(self, a: dict):
        import doc_ai
        args = (a['title'], doc_ai.task_deadline(a), doc_ai.task_notes(a), doc_ai.related_line(a))
        try:
            if self.register_task_callback:
                self.register_task_callback(*args)
            else:
                import db
                wid = db.get_first_window_id() or db.create_window()
                db.add_task(wid, args[0], args[1], notes=args[2], related_no=args[3])
                db.set_setting('external_change', __import__('datetime').datetime.now().isoformat())
        except Exception as e:
            QMessageBox.warning(self, '업무 등록 실패', str(e))
            return
        if not self.editor.toPlainText().strip():
            self.close()          # 공문 본문을 쓰지 않고 업무만 등록했으면 창을 닫는다
            return
        self.statusBar().showMessage(f'업무로 등록했습니다: {a["title"]}')

    def _ocr_fallback(self, pixmap, target='auto'):
        """Windows OCR로 캡처한 글자를 읽어 제목·관련 번호 칸에 넣는다. winrt 없으면 안내."""
        try:
            import winrt  # noqa: F401
        except ImportError:
            self.statusBar().showMessage('OCR을 쓸 수 없습니다 (Windows OCR 없음).')
            QMessageBox.warning(
                self, 'OCR 불가',
                'Windows OCR(winrt) 모듈이 없습니다.\n\n'
                'Windows 10 이상에서 쓸 수 있습니다. 제목·관련 번호를 직접 입력해 주세요.'
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


class _ModelSelectDialog(QDialog):
    """AI 모델 목록에서 클릭으로 선택하는 다이얼로그 (검색 가능)."""

    def __init__(self, title: str, models: list, current: str, parent=None, subtitle: str = ''):
        from PyQt5.QtWidgets import QListWidget
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setStyleSheet(_style())
        self.setMinimumWidth(px(420))
        self.setMinimumHeight(px(440))
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(20, 18, 20, 16)

        head = QLabel(title)
        head.setObjectName('DialogTitle')
        layout.addWidget(head)
        sub = QLabel(subtitle or '사용할 모델을 고르세요.')
        sub.setWordWrap(True)
        layout.addWidget(sub)

        self._search = QLineEdit()
        self._search.setPlaceholderText('모델 검색')
        self._search.addAction(qta.icon('fa5s.search', color=C.FAINT), QLineEdit.LeadingPosition)
        self._search.setMinimumHeight(34)
        self._search.textChanged.connect(self._filter)
        layout.addWidget(self._search)

        self._list = QListWidget()
        self._list.itemDoubleClicked.connect(self.accept)
        self._models = models
        self._populate(models, current)
        layout.addWidget(self._list, stretch=1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton('선택')
        btn_ok.setObjectName('Primary')
        btn_ok.setMinimumHeight(34)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)
        keep_light(self)

    def _populate(self, models, select=''):
        from PyQt5.QtWidgets import QListWidgetItem
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
