"""ui/profile_dialog.py — 「학교 정보·개인정보 보호」 설정 창."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox, QDialog, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
    QPushButton, QVBoxLayout,
)

from privacy import parse_name_list
from school_profile import (PROFILE_FIELDS, load_profile, save_profile,
                            privacy_enabled, set_privacy_enabled)
from .common import pr_font
from .theme import T, tc, tct

_STYLE = f"""
    QDialog {{ background: {T.BG_WARM}; }}
    QLabel {{ color: {T.TEXT_WARM}; background: transparent; }}
    QLabel#Title {{ color: {T.INK}; font-size: 13pt; font-weight: 700; }}
    QLabel#Hint {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QLineEdit, QPlainTextEdit {{
        background: {tc("white")}; border: 1px solid {T.SURFACE_WARM}; border-radius: 4px;
        padding: 3px 6px; color: {T.INK};
    }}
    QLineEdit:focus, QPlainTextEdit:focus {{ border: 1px solid {T.ACCENT}; }}
    QPushButton {{
        background: {T.SURFACE_WARM}; color: {T.TEXT_ON_NOTE}; border: none; border-radius: 5px;
        padding: 6px 14px;
    }}
    QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
    QPushButton#Primary {{ background: {T.ACCENT}; color: {tct("white")}; font-weight: 700; }}
    QPushButton#Primary:hover {{ background: {T.ACCENT_HOVER}; }}
"""


class ProfileDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window | Qt.WindowStaysOnTopHint)
        self.setWindowTitle('학교 정보 · 개인정보 보호')
        self.setStyleSheet(_STYLE)
        self.setFont(pr_font(10))
        self.setFixedWidth(460)  # 폭 고정 → 줄바꿈 안내문 높이가 일정하게 계산됨

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 14)
        root.setSpacing(8)
        title = QLabel('🏫 학교 정보')
        title.setObjectName('Title')
        root.addWidget(title)
        hint = QLabel('공문 초안·회신 공문을 만들 때 자동으로 들어갑니다. 비워 둔 칸은 쓰지 않습니다.')
        hint.setObjectName('Hint')
        hint.setWordWrap(True)
        root.addWidget(hint)

        p = load_profile()
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setVerticalSpacing(6)
        self._edits = {}
        for key, label, placeholder in PROFILE_FIELDS:
            ed = QLineEdit(p.get(key, ''))
            ed.setPlaceholderText(placeholder)
            form.addRow(label, ed)
            self._edits[key] = ed
        root.addLayout(form)

        root.addSpacing(6)
        title2 = QLabel('🛡️ 개인정보 보호')
        title2.setObjectName('Title')
        root.addWidget(title2)
        self.chk_privacy = QCheckBox('외부 AI로 보내기 전에 개인정보 가리기 (권장)')
        self.chk_privacy.setChecked(privacy_enabled())
        root.addWidget(self.chk_privacy)
        hint2 = QLabel('켜 두면 캡처한 공문은 이미지 대신 PC에서 읽은 글자만 보내고, 주민등록번호·전화번호·'
                       '이메일·아래 이름을 가립니다. 위 연락처와 담당자 이름은 가리지 않습니다. '
                       'PC 안에서 도는 AI(Ollama)에는 적용하지 않습니다.')
        hint2.setObjectName('Hint')
        hint2.setWordWrap(True)
        root.addWidget(hint2)
        self.ed_names = QPlainTextEdit(p.get('mask_names', ''))
        self.ed_names.setPlaceholderText('가릴 이름 (학생 명단 등) — 쉼표나 줄바꿈으로 구분\n예: 김철수, 이영희')
        self.ed_names.setFixedHeight(70)
        root.addWidget(self.ed_names)
        self.lbl_count = QLabel('')
        self.lbl_count.setObjectName('Hint')
        root.addWidget(self.lbl_count)
        self.ed_names.textChanged.connect(self._update_count)
        self._update_count()

        btns = QHBoxLayout()
        btns.addStretch()
        b_cancel = QPushButton('취소')
        b_ok = QPushButton('저장')
        b_ok.setObjectName('Primary')
        b_ok.setDefault(True)
        btns.addWidget(b_cancel)
        btns.addWidget(b_ok)
        root.addLayout(btns)
        b_cancel.clicked.connect(self.reject)
        b_ok.clicked.connect(self._save)
        self.resize(self.width(), self.sizeHint().height())  # 최소 높이가 아닌 권장 높이로 열기

    def _update_count(self):
        n = len(parse_name_list(self.ed_names.toPlainText()))
        self.lbl_count.setText(f'등록된 이름 {n}명' if n else '')

    def _save(self):
        data = {k: ed.text() for k, ed in self._edits.items()}
        data['mask_names'] = self.ed_names.toPlainText()
        save_profile(data)
        set_privacy_enabled(self.chk_privacy.isChecked())
        self.accept()
