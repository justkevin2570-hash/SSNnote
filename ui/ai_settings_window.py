"""「AI 연결」 창 — 업무창 메뉴 → 설정 → AI 연결 (빠른 찾기 · AI 대화 창에서도).

지금 쓰는 AI 한 줄, 구독 계정 연동 2개 · API 연동 3개 목록([연결] / 쓰는 중이면 ✓ [설정]), [AI 끄기].
실제 연결(키 입력 · 구독 로그인)은 ui/ai_connect.connect가 띄우는 창 그대로 쓴다.
공문 작성 · 예산 창 위쪽 AI 단추는 지금처럼 바로 펼쳐지는 메뉴(ai_connect.show_menu).
튜토리얼(AI 단계) 중에 열리면 맨 위에 유주 안내 띠(GuideBar)가 붙어 칸을 ①~⑥ 하나씩 설명한다.
"""
from PyQt5.QtCore import QRect, QRectF, Qt
from PyQt5.QtGui import QColor, QPainter, QPen
from PyQt5.QtWidgets import QDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from . import ai_connect

GROUPS = (('구독 계정 연동', ('claude_code', 'codex')),
          ('API 연동', ('opencode', 'claude', 'openai')))
_DOT = {'on': '#16A34A', 'legacy': '#D97706', 'off': '#94A3B8'}

# 튜토리얼 안내 띠: (이름, 설명, 강조할 칸, 유주 그림)
GUIDE = (
    ('지금', '지금 쓰는 AI예요. 초록 <span style="color:#16A34A">●</span>은 연결됨, '
             '회색 <span style="color:#94A3B8">●</span>은 꺼짐이에요.', 'now', 4),
    ('구독 계정 연동', 'Claude · ChatGPT 유료 구독이 있으면 이걸 골라요. '
                   '[연결] → 설치 → 로그인 → 연결 확인 → 저장. 추가 요금은 없어요.', 'sub', 2),
    ('API 연동', 'API 키가 있으면 골라요. [연결] → 키 붙여넣기 → 모델 고르기 → 저장. 쓴 만큼 요금이 나가요.', 'api', 3),
    ('✓ 설정', '지금 쓰는 AI에는 ✓가 붙고 [연결]이 [설정]으로 바뀌어요. [설정]에서 모델을 바꿀 수 있어요.', 'current', 5),
    ('AI 끄기', 'AI를 끄면 AI 기능만 쉬어요. 앱은 그대로 쓸 수 있어요.', 'off', 6),
    ('닫기', '다 봤으면 닫아요. 언제든 「…」 → 설정 → AI 연결에서 다시 열 수 있어요.', 'close', 1),
)
_NUM = '①②③④⑤⑥'
_PURPLE = '#7C3AED'


class _GuideRing(QWidget):
    """안내 중인 칸을 두르는 보라 테두리 (마우스는 아래로 그대로 통과)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.target = QRect()

    def paintEvent(self, _):
        if self.target.isEmpty():
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QPen(QColor(_PURPLE), 2.5))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(QRectF(self.target).adjusted(-5, -4, 5, 4), 10, 10)
        p.end()


class GuideBar(QFrame):
    """AI 연결 창 맨 위 유주 안내 띠: 「① 지금 1 / 6」 · 설명 · [이전] [다음] (마지막은 [다 봤어요])."""

    def __init__(self, dlg):
        super().__init__(dlg)
        from ui.common import px
        self.dlg = dlg
        self.page = 0
        self.seen_all = False
        self.setObjectName('GuideBar')
        self.setStyleSheet(f"""
            QFrame#GuideBar {{ background: #F5F3FF; border: 1.5px solid {_PURPLE}; border-radius: 10px; }}
            QFrame#GuideBar QLabel {{ background: transparent; border: none; }}
            QFrame#GuideBar QLabel#GuidePill {{ background: {_PURPLE}; color: white; border-radius: 9px;
                                font-weight: 700; padding: 1px 8px; }}
            QFrame#GuideBar QLabel#GuideCount {{ color: #6D28D9; font-weight: 600; }}
            QFrame#GuideBar QLabel#GuideText {{ color: #2E1065; }}
            QPushButton#GuideNext {{ background: {_PURPLE}; color: white; border: none; border-radius: 6px;
                                     padding: 5px 14px; font-weight: 700; }}
            QPushButton#GuideNext:hover {{ background: #6D28D9; }}
        """)
        outer = QHBoxLayout(self)
        outer.setContentsMargins(8, 10, 12, 10)
        outer.setSpacing(8)
        self.yuju = QLabel()
        self.yuju.setFixedWidth(px(56))
        self.yuju.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
        outer.addWidget(self.yuju, 0, Qt.AlignBottom)
        v = QVBoxLayout()
        v.setSpacing(6)
        top = QHBoxLayout()
        self.pill = QLabel()
        self.pill.setObjectName('GuidePill')
        self.count = QLabel()
        self.count.setObjectName('GuideCount')
        top.addWidget(self.pill)
        top.addStretch(1)
        top.addWidget(self.count)
        v.addLayout(top)
        self.text = QLabel()
        self.text.setObjectName('GuideText')
        self.text.setTextFormat(Qt.RichText)
        self.text.setWordWrap(True)
        self.text.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        v.addWidget(self.text)
        nav = QHBoxLayout()
        nav.addStretch(1)
        self.btn_prev = QPushButton('이전')
        self.btn_next = QPushButton('다음')
        self.btn_next.setObjectName('GuideNext')
        for b in (self.btn_prev, self.btn_next):
            b.setCursor(Qt.PointingHandCursor)
            b.setAutoDefault(False)
            nav.addWidget(b)
        v.addLayout(nav)
        v.addStretch(1)
        outer.addLayout(v, 1)
        self.btn_prev.clicked.connect(lambda: self.go(self.page - 1))
        self.btn_next.clicked.connect(self._on_next)
        self.ring = _GuideRing(dlg)

    def _text_width(self):
        from ui.common import px
        m = self.dlg.layout().contentsMargins()
        return max(px(120), self.dlg.width() - m.left() - m.right() - 8 - 12 - 3 - px(56) - 8)

    def go(self, page):
        from ui.onboarding import _yuju_pixmap
        from ui.common import px
        self.page = max(0, min(page, len(GUIDE) - 1))
        name, text, part, yn = GUIDE[self.page]
        last = self.page == len(GUIDE) - 1
        if last:
            self.seen_all = True
        self.pill.setText(f'{_NUM[self.page]}  {name}')
        self.count.setText(f'{self.page + 1} / {len(GUIDE)}')
        self.text.setText(text)
        tw = self._text_width()
        self.text.setFixedWidth(tw)
        self.text.setFixedHeight(self.text.heightForWidth(tw))
        self.btn_prev.setEnabled(self.page > 0)
        self.btn_next.setText('다 봤어요' if last else '다음')
        pm = _yuju_pixmap(yn, px(64), self.devicePixelRatioF() or 1.0)
        self.yuju.setPixmap(pm)
        self.dlg.adjustSize()
        self.place_ring()

    def _on_next(self):
        if self.page >= len(GUIDE) - 1:
            self.seen_all = True
            self.dlg.accept()
        else:
            self.go(self.page + 1)

    def place_ring(self):
        self.ring.setGeometry(self.dlg.rect())
        self.ring.target = self.dlg.guide_rect(GUIDE[self.page][2])
        if self.dlg.isVisible() and not self.ring.isVisible():
            self.ring.show()           # 창이 뜨는 도중에 만들어졌으면 저절로 안 보임
        self.ring.raise_()
        self.ring.update()


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
        self.group_labels = {}
        self.guide = None
        for title, keys in GROUPS:
            v.addSpacing(6)
            g = QLabel(title)
            g.setObjectName('AiGroup')
            v.addWidget(g)
            self.group_labels[keys] = g
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
        self.btn_close = btn_close = QPushButton('닫기')
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
        if getattr(self, 'guide', None) is not None:
            self.guide.place_ring()

    # ── 튜토리얼 안내 ─────────────────────────────────────────
    def start_guide(self):
        """맨 위에 유주 안내 띠를 붙이고 ①부터 보여 준다 (한 번만)."""
        if self.guide is None:
            self.guide = GuideBar(self)
            self.layout().insertWidget(0, self.guide)
            self.layout().insertSpacing(1, 4)
        self.guide.go(0)
        return self.guide

    def guide_rect(self, part) -> QRect:
        """안내할 칸의 자리 (창 좌표)."""
        def at(w):
            return QRect(w.mapTo(self, w.rect().topLeft()), w.size()) if w.isVisible() else QRect()
        if part == 'now':
            return at(self.lbl_now)
        if part in ('sub', 'api'):
            keys = GROUPS[0][1] if part == 'sub' else GROUPS[1][1]
            r = at(self.group_labels[keys])
            for k in keys:
                r = r.united(at(self.rows[k][0]))
            return r
        if part == 'current':
            from ai_client import load_ai_mode
            key = load_ai_mode()
            row = self.rows.get(key, self.rows[GROUPS[0][1][0]])[0]
            return at(row)
        if part == 'off':
            return at(self.btn_off)
        if part == 'close':
            return at(self.btn_close)
        return QRect()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self.guide is not None:
            self.guide.place_ring()

    def showEvent(self, e):
        super().showEvent(e)
        if self.guide is not None:
            self.guide.place_ring()

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
