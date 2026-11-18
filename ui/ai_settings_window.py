"""「AI 연결」 창 — 업무창 메뉴 → 설정 → AI 연결 (빠른 찾기 · AI 대화 창에서도).

지금 쓰는 AI 한 줄, 구독 계정 연동 2개 · API 연동 3개 목록([연결] / 쓰는 중이면 ✓ [설정]), [AI 끄기].
실제 연결(키 입력 · 구독 로그인)은 ui/ai_connect.connect가 띄우는 창 그대로 쓴다.
공문 작성 · 예산 창 위쪽 AI 단추는 지금처럼 바로 펼쳐지는 메뉴(ai_connect.show_menu).
"""
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from . import ai_connect

GROUPS = (('구독 계정 연동', ('claude_code', 'codex')),
          ('API 연동', ('opencode', 'claude', 'openai')))
_DOT = {'on': '#16A34A', 'legacy': '#D97706', 'off': '#94A3B8'}


class AiSettingsWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        from document_editor import _style, C
        from ui.theme import keep_light
        from ui.common import px
        self._C = C
        self.setWindowTitle('AI 연결')
        self.setStyleSheet(_style() + f"""
            QFrame#AiRow {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px; }}
            QFrame#AiRow[current="true"] {{ border: 1px solid {C.ACCENT}; background: {C.ACCENT_SOFT}; }}
            QLabel#AiGroup {{ color: {C.MUTED}; font-weight: 700; }}
            QLabel#AiName {{ background: transparent; font-weight: 600; color: {C.TEXT}; }}
            QLabel#AiCheck {{ background: transparent; color: {C.ACCENT}; font-weight: 800; }}
        """)
        self.setFixedWidth(px(380))
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 16)
        v.setSpacing(8)

        head = QLabel('AI 연결')
        head.setObjectName('DialogTitle')
        v.addWidget(head)
        self.lbl_now = QLabel()
        self.lbl_now.setTextFormat(Qt.RichText)
        self.lbl_now.setStyleSheet(f'color: {C.TEXT};')
        v.addWidget(self.lbl_now)

        self.rows = {}
        for title, keys in GROUPS:
            v.addSpacing(6)
            g = QLabel(title)
            g.setObjectName('AiGroup')
            v.addWidget(g)
            for key in keys:
                row = QFrame()
                row.setObjectName('AiRow')
                h = QHBoxLayout(row)
                h.setContentsMargins(12, 6, 8, 6)
                h.setSpacing(6)
                name = QLabel(ai_connect.MENU_LABELS[key])
                name.setObjectName('AiName')
                check = QLabel('✓')
                check.setObjectName('AiCheck')
                btn = QPushButton('연결')
                btn.setCursor(Qt.PointingHandCursor)
                btn.setMinimumWidth(px(64))
                btn.clicked.connect(lambda _, k=key: self._connect(k))
                h.addWidget(name)
                h.addWidget(check)
                h.addStretch(1)
                h.addWidget(btn)
                v.addWidget(row)
                self.rows[key] = (row, check, btn)

        v.addSpacing(10)
        bottom = QHBoxLayout()
        self.btn_off = QPushButton('AI 끄기')
        self.btn_off.setCursor(Qt.PointingHandCursor)
        self.btn_off.clicked.connect(self._turn_off)
        btn_close = QPushButton('닫기')
        btn_close.setObjectName('Primary')
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setMinimumHeight(32)
        btn_close.clicked.connect(self.accept)
        bottom.addWidget(self.btn_off)
        bottom.addStretch(1)
        bottom.addWidget(btn_close)
        v.addLayout(bottom)
        keep_light(self)
        ai_connect.on_change(self._on_changed)    # 다른 창에서 바꿔도 맞춤
        self.refresh()

    def refresh(self):
        from ai_client import load_ai_mode
        C = self._C
        text, state = ai_connect.status()
        if state == 'off':
            text = 'AI 꺼짐'
        self.lbl_now.setText(f'<span style="color:{_DOT[state]}">●</span>&nbsp; '
                             f'<span style="color:{C.MUTED}">지금</span>&nbsp; <b>{text}</b>')
        mode = load_ai_mode()
        for key, (row, check, btn) in self.rows.items():
            cur = key == mode
            row.setProperty('current', 'true' if cur else 'false')
            row.style().unpolish(row)
            row.style().polish(row)
            check.setVisible(cur)
            btn.setText('설정' if cur else '연결')
        self.btn_off.setEnabled(mode != 'none')

    def _on_changed(self):
        try:
            self.refresh()
        except RuntimeError:          # 이미 닫힌 창
            pass

    def _connect(self, key):
        ai_connect.connect(self, key)
        self.refresh()

    def _turn_off(self):
        ai_connect.turn_off()
        self.refresh()


def open_ai_settings(parent=None):
    """AI 연결 창을 띄운다 (닫을 때까지 기다림)."""
    dlg = AiSettingsWindow(parent)
    dlg.exec_()
    return dlg
