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
    QMessageBox, QInputDialog, QFrame, QSpinBox, QCheckBox, QDialog, QComboBox,
)
from PyQt5.QtCore import Qt, QTimer, QSize, QEvent, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QTextCursor, QTextCharFormat, QColor
from ai_client import (AiStreamThread, PROVIDER_LABELS,
                        load_ai_mode, save_ai_mode, legacy_ai_name, load_external_model_name,
                        claude_code_status, claude_code_login, claude_code_logout,
                        install_claude_code, test_claude_code, CLAUDE_CODE_MODELS, CLAUDE_CODE_INSTALL_PAGE,
                        load_claude_code_model, save_claude_code_model,
                        load_claude_code_path, save_claude_code_path, masking_required)
from school_profile import mask_for_ai
from doc_polish import (POLISH_SYSTEM, build_polish_prompt, diff_spans, restore_masked, leftover_masks)
from db import save_official_document, get_official_documents
import doc_style   # 내 공문 스타일: 비슷한 저장 공문 + 나만의 작성 규칙을 AI에 함께 보냄
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
    QPushButton#TypeChip {{ padding: 0 10px; text-align: left; font-weight: 600; color: {C.TEXT}; }}
    QPushButton#TypeChip[guess="true"] {{ color: {C.FAINT}; font-weight: 500; }}
    QPushButton#TypeChip[manual="true"] {{ border-color: {C.ACCENT_LINE}; color: {C.ACCENT}; }}
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


# 공문 종류 고르기 메뉴에 보일 순서
_DOC_TYPE_ORDER = ('general', 'result_report', 'submit', 'purchase',
                   'home_letter', 'recruit_announce', 'recruit_result')


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


class _BgCall(QThread):
    """오래 걸리는 일(claude 실행 등)을 뒤 스레드에서 한다. 끝나면 done(결과), 실패하면 failed(메시지)."""
    done = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, fn, *args):
        super().__init__()
        self._fn, self._args = fn, args

    def run(self):
        try:
            self.done.emit(self._fn(*self._args))
        except Exception as e:
            self.failed.emit(str(e))


class ClaudeCodeDialog(QDialog):
    """Claude 구독 연결: 설치 → 로그인 → 모델 → 연결 확인 → 저장. 모든 claude 호출은 뒤 스레드에서."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._job = None
        self._st = None
        self.setWindowTitle('Claude 구독 연결')
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setStyleSheet(_style())
        self.setFixedWidth(px(480))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(10)

        head = QLabel('Claude 구독 연결')
        head.setObjectName('DialogTitle')
        layout.addWidget(head)
        intro = QLabel('Claude Pro/Max 구독으로 동작합니다 · 사용량은 Claude 채팅·Claude Code와 같은 한도에서 차감 · '
                       '본인만 사용하세요 · 로그인 정보는 Claude Code가 보관하며 서서니노트는 저장하지 않습니다.')
        intro.setObjectName('Hint')
        intro.setWordWrap(True)
        intro.ensurePolished()                                 # 줄바꿈 높이를 미리 계산해 잘리지 않게
        intro.setMinimumHeight(intro.heightForWidth(px(480) - 44))
        layout.addWidget(intro)

        self.lbl_status = QLabel('확인 중…')
        self.lbl_status.setStyleSheet(f'color: {C.TEXT}; font-weight: 700;')
        self.lbl_status.setWordWrap(True)
        layout.addWidget(self.lbl_status)
        self.lbl_progress = QLabel()
        self.lbl_progress.setObjectName('Hint')
        self.lbl_progress.setWordWrap(True)
        layout.addWidget(self.lbl_progress)

        row = QHBoxLayout()
        row.setSpacing(6)
        self.btn_install = QPushButton('설치')
        self.btn_install.setIcon(qta.icon('fa5s.download', color=C.ACCENT))
        self.btn_install.setCursor(Qt.PointingHandCursor)
        self.btn_install.clicked.connect(self._install)
        self.btn_account = QPushButton('Claude로 로그인')
        self.btn_account.setCursor(Qt.PointingHandCursor)
        self.btn_account.clicked.connect(self._login_or_logout)
        row.addWidget(self.btn_install)
        row.addWidget(self.btn_account)
        row.addStretch()
        layout.addLayout(row)

        self.lbl_link = QLabel(f'설치가 안 되면 공식 설치 안내를 보세요: '
                               f'<a href="{CLAUDE_CODE_INSTALL_PAGE}">{CLAUDE_CODE_INSTALL_PAGE}</a>')
        self.lbl_link.setObjectName('Hint')
        self.lbl_link.setOpenExternalLinks(True)
        layout.addWidget(self.lbl_link)

        layout.addSpacing(4)
        lbl_model = QLabel('AI 모델')
        lbl_model.setObjectName('FieldLabel')
        layout.addWidget(lbl_model)
        self.cmb_model = QComboBox()
        for value, label in CLAUDE_CODE_MODELS:
            self.cmb_model.addItem(label, value)
        self.cmb_model.setCurrentIndex(max(0, self.cmb_model.findData(load_claude_code_model())))
        layout.addWidget(self.cmb_model)

        lbl_path = QLabel('claude 실행 파일 (비우면 자동으로 찾아요)')
        lbl_path.setObjectName('FieldLabel')
        layout.addWidget(lbl_path)
        self.edit_path = QLineEdit(load_claude_code_path())
        self.edit_path.setPlaceholderText('예: C:\\Users\\내이름\\.local\\bin\\claude.exe')
        self.edit_path.setMinimumHeight(34)
        layout.addWidget(self.edit_path)

        test_row = QHBoxLayout()
        test_row.setSpacing(8)
        self.btn_test = QPushButton('연결 확인')
        self.btn_test.clicked.connect(self._test)
        self.lbl_test = QLabel()
        self.lbl_test.setWordWrap(True)
        test_row.addWidget(self.btn_test)
        test_row.addWidget(self.lbl_test, 1)
        layout.addLayout(test_row)
        layout.addSpacing(6)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_save = QPushButton('저장')
        self.btn_save.setObjectName('Primary')
        self.btn_save.setMinimumHeight(34)
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self._save)
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_save)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)
        keep_light(self)
        self.btn_install.hide()
        self.btn_account.hide()
        self.lbl_link.hide()

        self._refresh_status()

    # ── 뒤 스레드 실행: 한 번에 하나만, 끝나면 on_done / 실패하면 on_fail ──
    def _run(self, fn, *args, progress='', on_done, on_fail=None):
        if self._job is not None and self._job.isRunning():
            return False
        self._set_busy(True, progress)
        job = _BgCall(fn, *args)
        job.done.connect(lambda r: self._finish(on_done, r))
        job.failed.connect(lambda m: self._finish_fail(on_fail, m))
        self._job = job
        job.start()
        return True

    def _set_busy(self, busy, progress=''):
        self.lbl_progress.setText(progress)
        for w in (self.btn_install, self.btn_account, self.btn_test, self.btn_save):
            w.setEnabled(not busy)

    def _finish(self, on_done, result):
        self._set_busy(False)
        on_done(result)

    def _finish_fail(self, on_fail, msg):
        self._set_busy(False)
        if on_fail:
            on_fail(msg)
        else:
            self._show_error('Claude 구독', msg)

    def _show_error(self, title, msg):
        self.lbl_progress.setText('')
        QMessageBox.warning(self, title, msg)

    # ── 상태 ──
    def _refresh_status(self, note=''):
        self.lbl_status.setText('확인 중…')
        self._run(claude_code_status, on_done=lambda st: self._after_status(st, note))

    def _after_status(self, st, note=''):
        self._show_status(st)
        self.lbl_progress.setText(note)

    def _show_status(self, st):
        self._st = st
        if not st['installed']:
            text, color = '설치 안 됨 — [설치]를 눌러 주세요.', C.WARN
        elif st['logged_in']:
            plan = st['plan'].capitalize() if st['plan'] else ''
            who = ' · '.join(x for x in (st['email'], f'{plan} 구독' if plan else '') if x)
            text, color = f'✓ 로그인됨 ({who})', C.OK
        else:
            text, color = '로그인 필요 — [Claude로 로그인]을 눌러 주세요.', C.WARN
        self.lbl_status.setText(text)
        self.lbl_status.setStyleSheet(f'color: {color}; font-weight: 700;')
        self.btn_install.setVisible(not st['installed'])
        self.lbl_link.setVisible(not st['installed'])
        self.btn_account.setVisible(st['installed'])
        self.btn_account.setText('로그아웃' if st['logged_in'] else 'Claude로 로그인')

    # ── 동작 ──
    def _install(self):
        self._run(install_claude_code, progress='Claude Code를 설치하는 중… (1~2분 걸릴 수 있어요)',
                  on_done=lambda _: self._refresh_status('설치했어요.'),
                  on_fail=lambda m: self._show_error('설치 실패', m))

    def _login_or_logout(self):
        if self._st and self._st.get('logged_in'):
            ans = QMessageBox.question(self, 'Claude 로그아웃', 'Claude 구독 로그인을 끊을까요?',
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans != QMessageBox.Yes:
                return
            self._run(claude_code_logout, progress='로그아웃하는 중…',
                      on_done=lambda _: self._refresh_status('로그아웃했어요.'),
                      on_fail=lambda m: self._show_error('로그아웃 실패', m))
            return
        self._run(claude_code_login, progress='브라우저에서 Claude 로그인 후 승인해 주세요…',
                  on_done=lambda _: self._refresh_status('로그인했어요.'),
                  on_fail=lambda m: self._show_error('로그인 실패', m))

    def _test(self):
        self._save_path_and_model()
        self.lbl_test.setText('확인하는 중… (몇 초 걸려요)')
        self._run(test_claude_code, self.cmb_model.currentData() or '',
                  on_done=self._show_test, on_fail=lambda m: self._show_test((False, m)))

    def _show_test(self, result):
        ok, msg = result
        self.lbl_test.setText(('✓ ' if ok else '⚠ ') + msg)
        self.lbl_test.setStyleSheet(f'color: {C.OK if ok else C.WARN};')

    def _save_path_and_model(self):
        save_claude_code_path(self.edit_path.text().strip())
        save_claude_code_model(self.cmb_model.currentData() or '')

    def _save(self):
        path = self.edit_path.text().strip()
        if path and not os.path.isfile(path):
            QMessageBox.warning(self, 'Claude 구독', '그 경로에 claude 파일이 없어요. 경로를 확인하거나 비워 두세요.')
            return
        if not (self._st and self._st.get('installed')):
            QMessageBox.warning(self, 'Claude 구독', '먼저 [설치]를 눌러 Claude Code를 설치해 주세요.')
            return
        if not self._st.get('logged_in'):
            QMessageBox.warning(self, 'Claude 구독', '먼저 [Claude로 로그인]을 눌러 로그인해 주세요.')
            return
        self._save_path_and_model()
        save_ai_mode('claude_code')
        self.accept()


class PolishDialog(QDialog):
    """선택 부분 다듬기 비교: 왼쪽 원문 · 오른쪽 다듬은 글(고칠 수 있음). 바뀐 단어는 색으로 표시."""
    _DEL = QColor('#FDE2E2')      # 원문에서 빠진 단어
    _INS = QColor('#DCFCE7')      # 다듬은 글에서 새로 들어간 단어

    def __init__(self, original: str, polished: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle('선택 부분 다듬기')
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setStyleSheet(_style())
        self.setMinimumSize(px(780), px(440))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(10)
        head = QLabel('선택 부분 다듬기')
        head.setObjectName('DialogTitle')
        layout.addWidget(head)
        hint = QLabel('왼쪽은 원래 글, 오른쪽은 다듬은 글입니다. 오른쪽 글은 직접 고칠 수 있어요. '
                      '바뀐 단어는 색으로 표시됩니다.')
        hint.setObjectName('Hint')
        hint.setWordWrap(True)
        layout.addWidget(hint)

        cols = QHBoxLayout()
        cols.setSpacing(12)
        left = QVBoxLayout()
        lbl_l = QLabel('원문')
        lbl_l.setObjectName('FieldLabel')
        left.addWidget(lbl_l)
        self._left = QTextEdit()
        self._left.setReadOnly(True)
        self._left.setPlainText(original)
        left.addWidget(self._left, 1)
        right = QVBoxLayout()
        lbl_r = QLabel('다듬은 글 (고칠 수 있어요)')
        lbl_r.setObjectName('FieldLabel')
        right.addWidget(lbl_r)
        self._right = QTextEdit()
        self._right.setPlainText(polished)
        self._right.textChanged.connect(self._highlight)
        right.addWidget(self._right, 1)
        cols.addLayout(left, 1)
        cols.addLayout(right, 1)
        layout.addLayout(cols, 1)

        self._warn = QLabel('⚠ 가린 표시(○○·***)가 남아 있어요. 원래 값으로 고친 뒤 적용하세요.')
        self._warn.setStyleSheet(f'color: {C.WARN}; font-weight: 600;')
        layout.addWidget(self._warn)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._btn_apply = QPushButton('적용')
        self._btn_apply.setObjectName('Primary')
        self._btn_apply.setMinimumHeight(34)
        self._btn_apply.setDefault(True)
        self._btn_apply.clicked.connect(self.accept)
        btn_cancel = QPushButton('취소')
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self._btn_apply)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)
        keep_light(self)
        self._highlight()

    def result_text(self) -> str:
        return self._right.toPlainText().strip()

    def _highlight(self):
        left_spans, right_spans = diff_spans(self._left.toPlainText(), self._right.toPlainText())
        self._left.setExtraSelections(self._selections(self._left, left_spans, self._DEL))
        self._right.setExtraSelections(self._selections(self._right, right_spans, self._INS))
        self._warn.setVisible(leftover_masks(self._right.toPlainText()))

    @staticmethod
    def _selections(edit, spans, color):
        out = []
        for s, e in spans:
            cur = QTextCursor(edit.document())
            cur.setPosition(s)
            cur.setPosition(e, QTextCursor.KeepAnchor)
            sel = QTextEdit.ExtraSelection()
            sel.cursor = cur
            sel.format.setBackground(color)
            out.append(sel)
        return out


# AI 제공자별 설정 정보 (예산 관리 창과 함께 씀)
from ui.ai_connect import PROVIDERS as _PROVIDERS  # noqa: E402,F401 (테스트·예전 코드가 여기서 찾음)
from ui import ai_connect  # noqa: E402


class DocumentEditorWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ai_thread = None
        self._polish = None            # 다듬기 중인 글 정보 (None = 다듬는 중 아님)
        self._polish_undo = None       # 마지막 다듬기 적용 (되돌리기용)
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
        self.btn_apikey.setToolTip('AI 연결 바꾸기 (OpenCode · Claude · ChatGPT · Claude 구독)')
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
        self.panel.header_source = lambda: (self.title_input.text().strip(), self.ref_input.text().strip())
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
        self.splitter.splitterMoved.connect(self._schedule_state_save)   # 칸 폭을 끌어 바꾸면 기억
        root.addWidget(self.splitter, 1)

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
        self.title_input.textChanged.connect(self._on_title_changed)
        self.title_input.returnPressed.connect(self._try_generate_draft)
        grid.addWidget(field_label('제목'), 0, 0)
        grid.addWidget(self.title_input, 0, 1, 1, 6)

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
        # 공문 종류: 제목을 칠 때마다 자동으로 바뀌고, 눌러서 직접 고를 수도 있다
        self._doc_type_manual = None      # 직접 고른 종류 (None = 제목으로 자동 판별)
        self.btn_doc_type = QPushButton()
        self.btn_doc_type.setObjectName('TypeChip')
        self.btn_doc_type.setFixedHeight(_input_h)
        self.btn_doc_type.setCursor(Qt.PointingHandCursor)
        self.btn_doc_type.ensurePolished()
        self.btn_doc_type.setFixedWidth(
            max(self.btn_doc_type.fontMetrics().horizontalAdvance(f'{v}  ▾') for v in _DOC_TYPE_LABELS.values()) + 26)
        self.btn_doc_type.clicked.connect(self._show_doc_type_menu)
        grid.addWidget(field_label('관련 번호'), 1, 0)
        grid.addWidget(self.ref_input, 1, 1)
        grid.addWidget(self.btn_move_to_title, 1, 2)
        grid.addWidget(lbl_attach, 1, 3)
        grid.addWidget(self.spin_attach, 1, 4)
        grid.addWidget(self.btn_doc_type, 1, 5)
        grid.addWidget(self.btn_draft, 1, 6)
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
        self.editor.setContextMenuPolicy(Qt.CustomContextMenu)
        self.editor.customContextMenuRequested.connect(self._show_editor_menu)
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
        self.btn_copy_ref = QPushButton("관련번호 복사")
        self.btn_copy_ref.setIcon(qta.icon('fa5.copy', color=C.MUTED))
        self.btn_copy_ref.setMinimumHeight(_input_h)
        self.btn_copy_ref.clicked.connect(self._copy_ref)
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
        ctrl_row.addWidget(self.btn_copy_ref)
        ctrl_row.addWidget(self.btn_copy)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self.btn_save_exit)
        layout.addLayout(ctrl_row)

        self.setStyleSheet(_style())
        self.statusBar().setSizeGripEnabled(False)
        self._polish_link = QLabel()                         # 다듬기 적용 뒤 [되돌리기] 링크
        self._polish_link.setTextFormat(Qt.RichText)
        self._polish_link.linkActivated.connect(lambda _: self._undo_polish())
        self._polish_link.hide()
        self.statusBar().addPermanentWidget(self._polish_link)
        self.statusBar().showMessage("준비 완료")
        self._apply_ai_mode()
        ai_connect.on_change(self._on_ai_changed)   # 예산 창 등 다른 곳에서 바꿔도 맞춤
        self._update_draft_button()  # 초기 상태: 빈칸이면 비활성
        self._update_doc_type_chip()
        self._apply_editor_line_height()

    # ── 선택 부분 다듬기 ──
    def _show_editor_menu(self, pos):
        from PyQt5.QtWidgets import QMenu, QAction
        menu = self.editor.createStandardContextMenu()
        first = menu.actions()[0] if menu.actions() else None
        act = QAction('선택 부분 다듬기  (선택이 없으면 본문 전체)', menu)
        act.setEnabled(bool(self.editor.toPlainText().strip()) and self._polish is None)
        act.triggered.connect(self.polish_selection)
        menu.insertAction(first, act)
        menu.insertSeparator(first)
        menu.exec_(self.editor.mapToGlobal(pos))

    def polish_selection(self):
        """선택한 글(없으면 본문 전체)을 AI가 공문 문체로 다듬는다. 결과는 비교 창에서 확인한 뒤 적용."""
        if self._polish is not None or (self.ai_thread and self.ai_thread.isRunning()):
            return
        if load_ai_mode() == 'none':
            self.statusBar().showMessage('AI가 꺼져 있어요. [AI 설정]에서 연결하면 다듬을 수 있어요.')
            return
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            start, end = cursor.selectionStart(), cursor.selectionEnd()
            text = cursor.selectedText().replace('\u2029', '\n')
        else:
            start, end = 0, self.editor.document().characterCount() - 1
            text = self.editor.toPlainText()
        if not text.strip():
            self.statusBar().showMessage('다듬을 글이 없습니다. 글을 선택하거나 본문을 써 주세요.')
            return

        counts = {}
        masked = text
        if masking_required():
            masked, counts = mask_for_ai(text)      # 가린 채로 보내고, 답이 오면 되살린다 (doc_polish)
        title = self.title_input.text().strip()
        examples = []
        try:
            docs = get_official_documents()
            for d in docs:
                d['doc_type'] = d.get('doc_type') or _detect_doc_type(d.get('title', ''))
            examples = doc_style.search_similar(title, docs, doc_type=self._current_doc_type(), top_k=2) if title else []
        except Exception:
            pass
        prompt = build_polish_prompt(masked, doc_style.style_notes_prompt(), doc_style.examples_prompt(examples))

        self._polish = {'start': start, 'end': end, 'original': text, 'masked': masked, 'counts': counts}
        self._set_polish_busy(True)
        self.statusBar().showMessage('선택한 글을 AI가 다듬는 중… (Esc: 멈추기)')
        self.ai_thread = AiStreamThread(prompt, system=POLISH_SYSTEM, structured=False)
        self.ai_thread.finished_signal.connect(self._on_polish_done)
        self.ai_thread.error_signal.connect(self._on_polish_error)
        self.ai_thread.start()

    def _set_polish_busy(self, busy: bool):
        self.editor.setReadOnly(busy)           # 다듬는 동안 본문은 고정 (선택 위치가 어긋나지 않게)
        self._set_controls_enabled(not busy)
        if busy:
            self._hide_polish_link()

    def _cancel_polish(self):
        if self.ai_thread and self.ai_thread.isRunning():
            self.ai_thread.stop()
        self._polish = None
        self._set_polish_busy(False)
        self.statusBar().showMessage('다듬기를 멈췄어요. 원래 글 그대로예요.')

    def _on_polish_error(self, msg):
        if self._polish is None:
            return
        self._polish = None
        self._set_polish_busy(False)
        self.statusBar().showMessage(f'다듬기 오류: {msg}')

    def _on_polish_done(self, result):
        ctx, self._polish = self._polish, None
        self._set_polish_busy(False)
        if ctx is None:
            return
        polished = (result or '').strip()
        if not polished:
            self.statusBar().showMessage('AI가 빈 답을 보냈어요. 원래 글 그대로예요. 다시 시도해 주세요.')
            return
        if ctx['counts']:
            polished, restored = restore_masked(polished, ctx['original'], ctx['masked'])
            if not restored:
                self.statusBar().showMessage('가린 개인정보를 되살릴 수 없어 적용하지 않았어요. '
                                             '개인정보가 든 부분을 빼고 다시 다듬어 주세요.')
                return
        dlg = PolishDialog(ctx['original'], polished, self)
        if dlg.exec_() != QDialog.Accepted:
            self.statusBar().showMessage('다듬기를 취소했어요. 원래 글 그대로예요.')
            return
        self._apply_polish(ctx['start'], ctx['end'], ctx['original'], dlg.result_text())

    def _apply_polish(self, start, end, original, new_text) -> bool:
        if self.editor.toPlainText()[start:end] != original:
            self.statusBar().showMessage('그 사이 본문이 바뀌어서 적용하지 못했어요. 다시 선택해 주세요.')
            return False
        self._replace_range(start, end, new_text)
        self._polish_undo = {'start': start, 'old': original, 'new': new_text}
        self.statusBar().showMessage('다듬은 글을 적용했어요. Ctrl+Z를 누르면 원래 글로 돌아갑니다.')
        self._polish_link.setText('<a href="undo" style="color:%s;">되돌리기</a>' % C.ACCENT)
        self._polish_link.show()
        return True

    def _undo_polish(self):
        info = self._polish_undo
        if not info:
            return
        self._polish_undo = None
        self._hide_polish_link()
        cur = self.editor.toPlainText()
        s, new = info['start'], info['new']
        if cur[s:s + len(new)] != new:
            self.statusBar().showMessage('적용한 뒤 글이 더 바뀌어서 되돌리지 못했어요. Ctrl+Z를 써 주세요.')
            return
        self._replace_range(s, s + len(new), info['old'])
        self.statusBar().showMessage('다듬기 전 글로 되돌렸어요.')

    def _hide_polish_link(self):
        self._polish_link.hide()

    def _replace_range(self, start, end, text):
        """본문의 [start, end) 글을 text로 바꾼다 — 편집 한 번이라 Ctrl+Z 한 번이면 원래대로."""
        from PyQt5.QtGui import QTextBlockFormat
        cursor = self.editor.textCursor()
        self._bulk_edit = True
        try:
            cursor.beginEditBlock()
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.KeepAnchor)
            cursor.insertText(text)
            cursor.setPosition(start)
            cursor.setPosition(start + len(text), QTextCursor.KeepAnchor)
            fmt = QTextBlockFormat()
            fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
            cursor.mergeBlockFormat(fmt)
            cursor.endEditBlock()
        finally:
            self._bulk_edit = False
        cursor.setPosition(start + len(text))
        self.editor.setTextCursor(cursor)

    # ── AI 초안 생성 ──
    def generate_draft(self):
        title = self.title_input.text().strip()
        if not title:
            self.statusBar().showMessage("제목을 먼저 입력하세요.")
            return

        doc_type   = self._current_doc_type()
        type_label = _DOC_TYPE_LABELS.get(doc_type, '일반 계획/운영')
        ref        = self.ref_input.text().strip()
        template   = _get_template(title, doc_type, ref, self.spin_attach.value())
        mode       = load_ai_mode()

        # 본문에 직접 쓴 글이 있으면 바꾸기 전에 물어본다 (방금 만든 초안 그대로면 묻지 않음)
        current = self.editor.toPlainText()
        if current.strip() and current != getattr(self, '_last_draft', None):
            ans = QMessageBox.question(
                self, '자동 초안',
                f'본문에 쓴 내용을 새 「{type_label}」 초안으로 바꿀까요?\n'
                '바꾼 뒤에도 Ctrl+Z를 누르면 지금 내용으로 돌아옵니다.',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if ans != QMessageBox.Yes:
                return

        # 1단계: 뼈대 템플릿 (AI 없이도 바로 사용 가능) — Ctrl+Z 한 번이면 원래 글로
        if mode == 'none':
            self._replace_body(template)
            self._last_draft = self.editor.toPlainText()
            self.statusBar().showMessage(f'템플릿 삽입 완료 — {type_label} (Ctrl+Z: 이전 내용)')
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

        # 내 스타일: 예전에 저장한 비슷한 공문 + 나만의 작성 규칙
        examples = []
        try:
            docs = get_official_documents()
            for d in docs:
                d['doc_type'] = d.get('doc_type') or _detect_doc_type(d.get('title', ''))
            examples = doc_style.search_similar(title, docs, doc_type=doc_type, top_k=2)
        except Exception:
            pass  # 못 찾아도 초안은 만든다
        notes = doc_style.load_style_notes()

        prompt = (
            f"제목: {title}\n{ref_line}"
            f"공문 종류: {type_label}\n"
            f"{doc_style.examples_prompt(examples)}"
            f"{doc_style.style_notes_prompt(notes)}\n"
            f"{self._received_context()}"
            '반드시 JSON 형식({"body": "공문 본문"})으로만 출력하세요.'
        )
        used = [f'내 공문 {len(examples)}개'] if examples else []
        if notes:
            used.append('작성 규칙')
        self._start_ai(prompt, system=system, structured=True, fallback=template,
                       note=(' · '.join(used) + ' 참고') if used else '')

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
        notes = doc_style.style_notes_prompt()
        if notes:
            prompt = f"{notes}\n{prompt}"

        self._start_ai(prompt, note='작성 규칙 참고' if notes else '')

    # ── 공통 AI 실행 ──
    # AI가 쓰는 동안 원래 내용을 보관한다. 오류·빈 답·취소면 원래 내용으로 되돌리고,
    # 성공하면 결과를 '되돌리기(Ctrl+Z) 한 번'으로 원래 내용에 돌아갈 수 있게 넣는다.
    def _start_ai(self, prompt, system=None, structured=False, fallback=None, note=''):
        """fallback: AI가 실패·취소되면 원래 내용 대신 넣을 글 (자동 초안의 뼈대). Ctrl+Z로 원래 내용.
        note: 상태줄에 함께 보일 참고 내용 (예: '내 공문 2개 · 작성 규칙 참고')."""
        if self.ai_thread and self.ai_thread.isRunning():
            return
        self._ai_fallback = fallback
        self._ai_note = note

        msg = "AI가 생각 중..." if not structured else "AI가 공문을 작성 중..."
        self.statusBar().showMessage(f"{msg} ({note})" if note else msg)
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
        """AI가 쓰기 전 내용으로 되돌린다. 자동 초안이었으면 뼈대를 넣는다 (Ctrl+Z: 원래 내용).
        뼈대를 넣었으면 True."""
        self.editor.blockSignals(True)
        self.editor.setPlainText(getattr(self, '_before_ai', ''))
        self.editor.blockSignals(False)
        self._apply_editor_line_height()
        self.editor.setUndoRedoEnabled(True)
        fallback = getattr(self, '_ai_fallback', None)
        if fallback:
            self._replace_body(fallback)
            self._last_draft = self.editor.toPlainText()
            return True
        return False

    def _replace_body(self, text):
        """본문을 text로 통째로 바꾼다 — 한 번의 편집이라 Ctrl+Z 한 번이면 이전 내용으로."""
        from PyQt5.QtGui import QTextBlockFormat
        cursor = self.editor.textCursor()
        self._bulk_edit = True
        try:
            cursor.beginEditBlock()
            cursor.select(QTextCursor.Document)
            cursor.insertText(text)
            cursor.select(QTextCursor.Document)                   # 줄간격도 같은 편집 안에서 (되돌리기 1번)
            fmt = QTextBlockFormat()
            fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
            cursor.mergeBlockFormat(fmt)
            cursor.endEditBlock()
        finally:
            self._bulk_edit = False
        cursor.movePosition(QTextCursor.End)
        self.editor.setTextCursor(cursor)

    def _on_finished(self, _, gen=None):
        if not self._is_current_ai(gen):
            return
        self._ai_gen = getattr(self, '_ai_gen', 0) + 1          # 이 요청은 끝 — 이후 신호 무시
        raw = self.editor.toPlainText()
        if not raw.strip():
            if self._restore_before_ai():
                self.statusBar().showMessage("AI가 빈 답을 보냈습니다 — 뼈대만 넣었습니다. 다시 시도해 주세요.")
            else:
                self.statusBar().showMessage("AI가 빈 답을 보냈습니다 — 원래 내용을 그대로 두었습니다. 다시 시도해 주세요.")
            self._set_controls_enabled(True)
            return
        # AI 응답 후 공문 규정 자동 보정
        corrected = _correct_document_format(raw)
        # 원래 내용 → (한 번의 편집으로) 새 내용: Ctrl+Z 한 번이면 원래 내용으로 돌아감
        self._ai_fallback = None
        self._restore_before_ai()
        self._replace_body(corrected)
        self._last_draft = self.editor.toPlainText()
        masked = getattr(self.ai_thread, 'masked_counts', None) or {}
        note = getattr(self, '_ai_note', '')
        note = f"{note} · " if note else ''
        if masked:
            from privacy import describe
            self.statusBar().showMessage(f"완료 — {note}🛡 개인정보 {describe(masked)}을 가리고 보냈습니다. (Ctrl+Z: 이전 내용)")
        else:
            self.statusBar().showMessage(f"완료 — {note}Ctrl+Z를 누르면 이전 내용으로 돌아갑니다.")
        self._set_controls_enabled(True)
        self.ai_input.clear()

    def _on_error(self, msg, gen=None):
        if not self._is_current_ai(gen):
            return
        self._ai_gen = getattr(self, '_ai_gen', 0) + 1
        if self._restore_before_ai():
            self.statusBar().showMessage(f"오류: {msg} — 뼈대만 넣었습니다. (Ctrl+Z: 이전 내용)")
        else:
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
        """제목이 있어야 초안을 만들 수 있다 (관련 번호만으로는 만들 수 없음)."""
        self.btn_draft.setEnabled(bool(self.title_input.text().strip()))
        self.btn_draft.setText('자동 초안 생성')
        self.btn_draft.setToolTip('제목에 맞는 공문 뼈대를 넣고, AI가 켜져 있으면 내용까지 채웁니다 (Enter)'
                                  if self.title_input.text().strip() else '제목을 먼저 입력하세요')

    # ── 공문 종류 ──
    def _current_doc_type(self) -> str:
        return self._doc_type_manual or _detect_doc_type(self.title_input.text())

    def _on_title_changed(self, text):
        if not text.strip():
            self._doc_type_manual = None   # 제목을 비우면(새 공문) 다시 자동 판별
        self._update_doc_type_chip()

    def _update_doc_type_chip(self):
        doc_type = self._current_doc_type()
        manual = self._doc_type_manual is not None
        # 아무 키워드도 안 맞아 '일반'으로 간 경우는 흐리게 — 추측한 것
        guess = not manual and doc_type == 'general'
        self.btn_doc_type.setText(f"{_DOC_TYPE_LABELS[doc_type]}  ▾")
        if manual:
            tip = '직접 고른 공문 종류 — 눌러서 바꾸거나 [자동]으로 되돌리기'
        elif guess:
            tip = '제목에서 종류를 못 찾아 일반 계획/운영으로 만듭니다 — 눌러서 바꾸기'
        else:
            tip = '제목으로 판별한 공문 종류 — 눌러서 바꾸기'
        self.btn_doc_type.setToolTip(tip)
        for name, val in (('manual', manual), ('guess', guess)):
            self.btn_doc_type.setProperty(name, 'true' if val else 'false')
        self.btn_doc_type.style().unpolish(self.btn_doc_type)
        self.btn_doc_type.style().polish(self.btn_doc_type)

    def _show_doc_type_menu(self):
        from PyQt5.QtWidgets import QMenu
        menu = QMenu(self)
        menu.setStyleSheet(_style())
        auto_type = _detect_doc_type(self.title_input.text())
        manual = self._doc_type_manual

        def mark(on):
            return '✓  ' if on else '     '
        act_auto = menu.addAction(f"{mark(manual is None)}자동 — 제목으로 판별 ({_DOC_TYPE_LABELS[auto_type]})")
        menu.addSeparator()
        actions = {}
        for key in _DOC_TYPE_ORDER:
            actions[menu.addAction(f"{mark(manual == key)}{_DOC_TYPE_LABELS[key]}")] = key
        chosen = menu.exec_(self.btn_doc_type.mapToGlobal(self.btn_doc_type.rect().bottomLeft()))
        if chosen is None:
            return
        self._doc_type_manual = None if chosen is act_auto else actions[chosen]
        self._update_doc_type_chip()

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

    def _copy_ref(self):
        """관련 번호(받은 공문 번호)를 클립보드에 복사."""
        ref = self.ref_input.text().strip()
        if not ref:
            self.statusBar().showMessage('복사할 관련 번호가 없습니다.')
            return
        QApplication.clipboard().setText(ref)
        self.statusBar().showMessage('관련 번호가 클립보드에 복사되었습니다.')
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
            if self._polish is not None:
                self._cancel_polish()          # 다듬기 중이면 본문은 건드리지 않고 멈추기만
            elif self.ai_thread and self.ai_thread.isRunning():
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
             '제목을 넣고 Enter나 [자동 초안]을 누르면 공문 뼈대가 들어가고, '
             'AI가 연결되어 있으면 내용까지 채웁니다. 공문 종류는 제목으로 정해져 버튼 옆에 보이고, '
             '눌러서 직접 고를 수 있습니다. 쓰던 본문은 Ctrl+Z로 되찾을 수 있습니다.'),
            ('fa5s.pen', '선택 부분 다듬기',
             '본문 글을 드래그한 뒤 우클릭 → 「선택 부분 다듬기」 (선택이 없으면 본문 전체). '
             '원문과 나란히 비교하고 고쳐서 [적용]하세요. 적용 후 Ctrl+Z나 상태 줄의 [되돌리기]로 원래 글에 돌아갑니다.'),
            ('fa5s.user-edit', '내 스타일로',
             '[저장 후 나가기]로 저장한 공문 중 제목이 비슷한 것을 AI가 예시로 보고 같은 말투로 씁니다. '
             '저장할수록 더 닮아 갑니다. 꼭 지킬 규칙은 [AI 설정] → 나만의 작성 규칙에 적어 두세요.'),
            ('fa5s.comment-dots', 'AI에게 요청',
             '아래 칸에 "장소를 시청각실로 바꿔줘"처럼 적고 Enter. 줄바꿈은 Shift+Enter. '
             '결과가 마음에 안 들면 Ctrl+Z 한 번으로 이전 내용으로 돌아갑니다.'),
            ('fa5s.crop-alt', '공문 캡처',
             'Ctrl+Shift+X나 [공문 캡처]로 받은 공문을 잘라 오면 제목·공문번호는 오른쪽 제목·관련 번호 칸에, '
             '요약(핵심 내용·기한)은 왼쪽 칸에 나옵니다. '
             '번호가 안 잡혔으면 그 부분만 한 번 더 캡처하세요 — 같은 공문에 합쳐서 다시 읽습니다. '
             '다른 공문은 왼쪽 [새 공문]. AI 대화·자동 초안도 그 공문을 참고합니다.'),
            ('fa5s.stop-circle', '취소·닫기', 'AI가 쓰는 중에 Esc를 누르면 멈추고 원래 내용으로 되돌립니다. '
             '그 밖에는 Esc로 창을 닫습니다.'),
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
            if self._restore_before_ai():
                self.statusBar().showMessage('AI 생성을 취소했습니다 — 뼈대만 넣었습니다. (Ctrl+Z: 이전 내용)')
            else:
                self.statusBar().showMessage('AI 생성을 취소하고 원래 내용으로 되돌렸습니다.')
            self._set_controls_enabled(True)

    def _apply_editor_line_height(self):
        """에디터 본문 줄간격 123% 적용."""
        if getattr(self, '_bulk_edit', False):
            return    # _replace_body가 같은 편집 안에서 직접 적용 (따로 하면 되돌리기가 여러 번으로 쪼개짐)
        from PyQt5.QtGui import QTextCursor, QTextBlockFormat
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.Document)
        fmt = QTextBlockFormat()
        fmt.setLineHeight(123, QTextBlockFormat.ProportionalHeight)
        cursor.mergeBlockFormat(fmt)

    def _on_ai_changed(self):
        try:
            self._apply_ai_mode()
        except RuntimeError:          # 이미 닫힌 창
            pass

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
            self.btn_apikey.setToolTip('AI 연결 바꾸기 (OpenCode · Claude · ChatGPT · Claude 구독)')
            self.ai_input.setPlaceholderText('AI에게 요청  (예: 장소를 시청각실로 바꾸고 일시는 11월 10일로 해줘)')
            return
        self.btn_apikey.setIcon(qta.icon('fa5s.circle', color=C.OFF))
        legacy = legacy_ai_name()
        if legacy:
            self.btn_apikey.setText('AI 다시 연결 필요')
            self.btn_apikey.setIcon(qta.icon('fa5s.exclamation-circle', color=C.WARN))
            self.ai_input.setPlaceholderText(
                f'이전에 쓰던 {legacy}는 더 이상 지원하지 않아요 — 위의 [AI 다시 연결 필요]를 눌러 '
                'OpenCode·Claude·ChatGPT·Claude 구독 중 하나를 연결하세요.')
        else:
            self.btn_apikey.setText('AI 꺼짐 · 연결하기')
            self.ai_input.setPlaceholderText('AI가 꺼져 있어요 — 위의 [AI 꺼짐 · 연결하기]에서 연결하면 쓸 수 있어요.')
        self.btn_apikey.setIconSize(QSize(9, 9) if not legacy else QSize(12, 12))

    # ── AI 설정 메뉴 ──
    def _show_ai_settings_menu(self):
        has_notes = bool(doc_style.load_style_notes())
        choice = ai_connect.show_menu(self, self.btn_apikey, _style(),
                                      extra=[('style', ('✓  ' if has_notes else '     ') + '나만의 작성 규칙…')])
        if choice is None:
            return
        if choice == 'style':
            self._edit_style_notes()
        elif choice == 'none':
            ai_connect.turn_off()
            self.statusBar().showMessage('AI를 껐습니다 — 자동 초안은 뼈대만 넣습니다.')
        else:
            ai_connect.connect(self, choice, self.statusBar().showMessage)

    def _edit_style_notes(self):
        """나만의 작성 규칙 — 자동 초안·AI 대화 때마다 AI에 함께 보낸다."""
        from PyQt5.QtWidgets import QPlainTextEdit
        dlg = QDialog(self)
        dlg.setWindowTitle('나만의 작성 규칙')
        dlg.setStyleSheet(_style())
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(24, 22, 24, 18)
        layout.setSpacing(10)
        head = QLabel('나만의 작성 규칙')
        head.setObjectName('DialogTitle')
        layout.addWidget(head)
        hint = QLabel('자동 초안과 AI 대화 때마다 AI에게 함께 알려 줍니다. 한 줄에 하나씩 적어 보세요.\n'
                      '저장한 공문도 비슷한 제목이면 예시로 함께 보내므로, 공문 자체의 말투는 저절로 따라갑니다.')
        hint.setWordWrap(True)
        layout.addWidget(hint)
        edit = QPlainTextEdit()
        edit.setPlaceholderText('예)\n'
                                "일시는 '2026. 5. 13.(수) 13:00 ~ 16:30' 형식으로\n"
                                "대상은 '전교생' 대신 '1~3학년 전체 학생'으로\n"
                                "'협조 바랍니다' 대신 '협조하여 주시기 바랍니다'로")
        edit.setPlainText(doc_style.load_style_notes())
        edit.setMinimumSize(px(460), px(180))
        layout.addWidget(edit, 1)
        count = QLabel()
        layout.addWidget(count)

        def update_count():
            n = len(edit.toPlainText().strip())
            count.setText(f'{n} / {doc_style.STYLE_NOTES_MAX}자' + ('  — 넘는 부분은 저장되지 않습니다' if n > doc_style.STYLE_NOTES_MAX else ''))
        edit.textChanged.connect(update_count)
        update_count()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton('취소')
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok = QPushButton('저장')
        btn_ok.setObjectName('Primary')
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)
        keep_light(dlg)
        if dlg.exec_() == QDialog.Accepted:
            doc_style.save_style_notes(edit.toPlainText())
            self.statusBar().showMessage('작성 규칙을 저장했습니다.' if doc_style.load_style_notes()
                                         else '작성 규칙을 비웠습니다.')

    def _busy(self, msg):
        self.statusBar().showMessage(msg)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

    def _idle(self):
        QApplication.restoreOverrideCursor()

    def _setup_keyed_ai(self, mode: str):
        ai_connect.connect(self, mode, self.statusBar().showMessage)

    def _setup_claude_code_ai(self):
        ai_connect.connect(self, 'claude_code', self.statusBar().showMessage)

    def _setup_claude_ai(self):
        self._setup_keyed_ai('claude')

    def _setup_openai_ai(self):
        self._setup_keyed_ai('openai')

    def _setup_opencode_ai(self):
        ai_connect.connect(self, 'opencode', self.statusBar().showMessage)

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

        # 저장한 공문은 다음 자동 초안 때 '내 공문' 예시로 쓰인다 (doc_style)
        save_official_document(title, doc_number, content, self._current_doc_type())
        self.statusBar().showMessage('저장되었습니다.')
        QTimer.singleShot(1000, lambda: self.statusBar().showMessage(''))

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

    # 창 위치·크기·최대화·「받은 공문」 칸 폭. 바꿀 때마다 잠깐 뒤 저장 (창을 닫지 않고 앱·PC를 꺼도 남게)
    _saved_panel_w = 0

    def _restore_window_state(self):
        import json
        try:
            with open(self._STATE_FILE, 'r', encoding='utf-8') as f:
                s = json.load(f)
            self.setGeometry(*fit_to_screen(s['x'], s['y'], s['width'], s['height'], min_w=560, min_h=420))
            self._restore_maximized = bool(s.get('maximized'))
            self._saved_panel_w = int(s.get('panel_w') or 0)
        except Exception:
            self.setGeometry(*default_geometry(900, 680))   # 처음엔 화면 크기에 맞춘 크기로

    def _save_window_state(self):
        import json, os
        if not getattr(self, '_state_ready', False):
            return
        try:
            os.makedirs(os.path.dirname(self._STATE_FILE), exist_ok=True)
            g = self.normalGeometry() if self.isMaximized() else self.geometry()
            if self.panel.isVisible():
                self._saved_panel_w = self.splitter.sizes()[0]
            data = {'x': g.x(), 'y': g.y(),
                    'width': g.width() - (self._panel_extra if self.panel.isVisible() else 0),
                    'height': g.height(), 'maximized': self.isMaximized(),
                    'panel_w': self._saved_panel_w}
            tmp = self._STATE_FILE + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(data, f)
            os.replace(tmp, self._STATE_FILE)
        except Exception:
            pass

    def _schedule_state_save(self, *_):
        t = getattr(self, '_state_timer', None)
        if t is None:
            t = self._state_timer = QTimer(self)
            t.setSingleShot(True)
            t.setInterval(800)
            t.timeout.connect(self._save_window_state)
        t.start()

    def reset_layout(self):
        """「화면 배치 기본값으로」: 창 크기·위치·받은 공문 칸 폭을 처음 상태로 (쓰던 글은 그대로)."""
        self.showNormal()
        self._show_panel(False)
        self._saved_panel_w = 0
        self.setGeometry(*default_geometry(900, 680))
        self._save_window_state()

    def showEvent(self, event):
        super().showEvent(event)
        if not getattr(self, '_state_ready', False):
            if getattr(self, '_restore_maximized', False):
                self.showMaximized()
            QTimer.singleShot(0, lambda: setattr(self, '_state_ready', True))

    def moveEvent(self, event):
        super().moveEvent(event)
        self._schedule_state_save()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._schedule_state_save()

    def closeEvent(self, event):
        t = getattr(self, '_state_timer', None)
        if t is not None:
            t.stop()
        self._save_window_state()
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
            if self._saved_panel_w:                      # 지난번에 끌어서 맞춘 칸 폭
                w = max(need, min(self._saved_panel_w, int(screen.width() * 0.5)))
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
        """AI 대화·자동 초안에 함께 보낼 받은 공문 정리 (받은 공문 칸에 내용이 있을 때)."""
        return self.panel.context() if self.panel.use_as_context() else ''

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
