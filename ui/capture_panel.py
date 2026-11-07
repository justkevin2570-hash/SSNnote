"""ui/capture_panel.py — 공문 작성 창 왼쪽의 「받은 공문」 칸.

캡처한 공문을 AI로 읽어, 제목·공문번호는 오른쪽 공문 작성 칸(제목·관련 번호)에 넣고(analyzed 시그널),
이 칸에는 요약(핵심 내용·기한)만 보여준다. [요약 복사]는 '제목 (관련 번호)' + 요약을 복사한다.
모양(QSS)은 공문 작성 창에서 물려받는다.

이어서 캡처: 제목·본문과 공문번호가 한 화면에 안 잡히면 두 번 캡처한다. [새 공문]을 누르기 전까지
캡처는 같은 공문에 덧붙고, 모은 캡처를 합쳐(이미지는 위아래로 이어 붙이고, 보호 모드면 읽은 글자를
이어 붙여) 다시 분석한다. 사용자가 요약을 직접 고쳤으면 다시 분석해도 덮어쓰지 않는다.
"""

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QFrame, QHBoxLayout, QLabel, QPlainTextEdit, QPushButton, QVBoxLayout,
)

import doc_ai
from privacy import describe

# 칸을 닫아도 분석 스레드가 끝날 때까지 참조를 유지 (실행 중인 QThread가 파괴되면 앱이 종료됨)
_ALIVE_WORKERS = set()


def _icon(name, color):
    try:
        import qtawesome as qta
        return qta.icon(name, color=color)
    except Exception:
        from PyQt5.QtGui import QIcon
        return QIcon()


def _merge(old: dict, new: dict) -> dict:
    """이어서 캡처: 새 결과에서 빈 값은 지금 값을 유지 (예: 공문번호만 캡처해도 제목·요약은 남음)."""
    out = dict(old)
    for k, v in new.items():
        if v or k not in out:
            out[k] = v
    out['needs_reply'] = bool(old.get('needs_reply')) or bool(new.get('needs_reply'))
    return out


class CapturePanel(QFrame):
    """시그널로 결과를 넘긴다.

    analyzed(str, str)    — 분석이 끝남: (제목, 공문번호) → 오른쪽 제목·관련 번호 칸에 넣는다
    collapse_requested()  — 칸 접기
    message(str)          — 상태 표시줄에 보일 안내

    header_source: () -> (제목, 관련 번호). 요약 복사·AI 참고에 쓸 머리말 (공문 작성 창이 넣어 줌).
    """
    collapse_requested = pyqtSignal()
    message = pyqtSignal(str)
    analyzed = pyqtSignal(str, str)

    MUTED = '#64748B'
    ACCENT = '#4F46E5'
    MAX_CAPTURES = 4          # 한 공문에 모을 수 있는 캡처 수 (넘으면 가장 오래된 것부터 뺌)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('Panel')
        self._pixmaps = []        # 이 공문의 캡처들
        self._ocr_texts = []      # 캡처별 PC에서 읽은 글자 (None = 아직 안 읽음)
        self._gen = 0             # 분석 요청 번호 (늦게 온 이전 결과 무시용)
        self._analysis = {}       # 지금까지 합친 분석 결과
        self._summary_edited = False
        self._filling = False
        self._worker = None
        self._closed = False
        self._has_content = False
        self.header_source = None
        self._build()
        self.clear()

    # ── 화면 구성 ────────────────────────────────────────────────
    def _small(self, text, icon=None, tip=''):
        b = QPushButton(text)
        b.setObjectName('Ghost')
        if icon:
            b.setIcon(_icon(icon, self.MUTED))
            b.setIconSize(QSize(12, 12))
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        return b

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(6)
        self._inner = self

        head = QHBoxLayout()
        title = QLabel('받은 공문')
        title.setObjectName('SectionTitle')
        head.addWidget(title)
        head.addStretch()
        self.btn_new = self._small('새 공문', 'fa5s.plus', '이 칸을 비우고 다른 공문을 새로 캡처')
        self.btn_new.clicked.connect(self.new_document)
        head.addWidget(self.btn_new)
        self.btn_collapse = self._small('', 'fa5s.angle-double-left', '받은 공문 칸 접기')
        self.btn_collapse.setFixedSize(28, 26)
        self.btn_collapse.clicked.connect(self.collapse_requested)
        head.addWidget(self.btn_collapse)
        root.addLayout(head)

        # 캡처 그림 (모은 순서대로 1, 2, …)
        self.thumb_row = QHBoxLayout()
        self.thumb_row.setSpacing(6)
        self.thumbs = []
        for i in range(self.MAX_CAPTURES):
            t = QLabel()
            t.setObjectName('Thumb')
            t.setFixedSize(60, 44)
            t.setAlignment(Qt.AlignCenter)
            t.setToolTip(f'캡처 {i + 1}')
            t.hide()
            self.thumbs.append(t)
            self.thumb_row.addWidget(t)
        self.thumb_row.addStretch()
        root.addLayout(self.thumb_row)
        self.lbl_status = QLabel('')
        self.lbl_status.setObjectName('Status')
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        root.addWidget(self.lbl_status)

        lb = QLabel('요약')
        lb.setObjectName('FieldLabel')
        root.addWidget(lb)
        self.ed_summary = QPlainTextEdit()
        self.ed_summary.setPlaceholderText('공문을 캡처하면 핵심 내용·기한이 여기에 정리됩니다.\n'
                                           '(Ctrl+Shift+X 또는 [공문 캡처])')
        self.ed_summary.textChanged.connect(self._on_summary_edited)
        root.addWidget(self.ed_summary, 1)

        self.btn_copy = QPushButton('요약 복사')
        self.btn_copy.setIcon(_icon('fa5.copy', self.MUTED))
        self.btn_copy.setMinimumHeight(32)
        self.btn_copy.setToolTip('제목 (관련 번호)와 요약을 함께 복사 — 메신저·회의 공유용')
        self.btn_copy.clicked.connect(self._copy_summary)
        root.addWidget(self.btn_copy)

    def content_min_width(self) -> int:
        """칸 안의 글자·버튼이 잘리지 않는 최소 폭 (배율·글꼴에 따라 달라짐)."""
        from ui.common import px
        return max(self.minimumSizeHint().width(), px(300))

    # ── 상태 ────────────────────────────────────────────────────
    def has_content(self) -> bool:
        return self._has_content

    def use_as_context(self) -> bool:
        return self._has_content and bool(self.summary().strip() or self._analysis.get('title'))

    def capture_count(self) -> int:
        return len(self._pixmaps)

    def summary(self) -> str:
        return self.ed_summary.toPlainText()

    def clear(self):
        self._gen += 1            # 진행 중인 분석 결과는 버린다
        self._has_content = False
        self._pixmaps, self._ocr_texts = [], []
        self._analysis = {}
        self._filling = True
        self.ed_summary.clear()
        self._filling = False
        self._summary_edited = False
        self._show_thumbs()
        self._set_status('')
        self._set_busy(False)
        self.btn_new.hide()

    def new_document(self):
        """[새 공문]: 모은 캡처와 칸을 비운다. 다음 캡처부터 새 공문."""
        self.clear()
        self._set_status('비웠습니다. 새 공문을 캡처하세요 (Ctrl+Shift+X).')
        self.message.emit('받은 공문 칸을 비웠습니다.')

    def _set_busy(self, busy: bool):
        self.ed_summary.setReadOnly(busy)
        self.btn_copy.setEnabled(not busy)

    def _set_status(self, text, state=''):
        self.lbl_status.setText(text)
        self.lbl_status.setVisible(bool(text))
        self.lbl_status.setProperty('state', state)
        self.lbl_status.style().unpolish(self.lbl_status)
        self.lbl_status.style().polish(self.lbl_status)

    def _on_summary_edited(self):
        if not self._filling:
            self._summary_edited = True    # 직접 고친 요약은 다시 분석해도 유지

    # ── 캡처 모으기 ──────────────────────────────────────────────
    def add_capture(self, pixmap: QPixmap):
        """캡처를 이 공문에 덧붙이고 모은 캡처 전체를 다시 분석한다 ([새 공문] 전까지 같은 공문)."""
        if len(self._pixmaps) >= self.MAX_CAPTURES:
            self._pixmaps.pop(0)
            self._ocr_texts.pop(0)
            self.message.emit(f'캡처는 {self.MAX_CAPTURES}장까지 모읍니다 — 가장 먼저 찍은 것을 뺐습니다.')
        self._pixmaps.append(pixmap)
        self._ocr_texts.append(None)
        self._has_content = True
        self.btn_new.show()
        self._show_thumbs()
        self._analyze()

    start = add_capture        # 예전 이름

    def _show_thumbs(self):
        for i, t in enumerate(self.thumbs):
            if i < len(self._pixmaps) and not self._pixmaps[i].isNull():
                t.setPixmap(self._pixmaps[i].scaled(60, 44, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                t.show()
            else:
                t.clear()
                t.hide()

    def _combined_pixmap(self) -> QPixmap:
        """모은 캡처를 위아래로 이어 붙인 한 장 (AI에 보낼 이미지)."""
        pms = [p for p in self._pixmaps if p is not None and not p.isNull()]
        if len(pms) <= 1:
            return pms[0] if pms else QPixmap()
        gap = 16
        w = max(p.width() for p in pms)
        h = sum(p.height() for p in pms) + gap * (len(pms) - 1)
        out = QPixmap(w, h)
        out.fill(QColor('white'))
        painter = QPainter(out)
        y = 0
        for p in pms:
            painter.drawPixmap(0, y, p)
            y += p.height() + gap
        painter.end()
        return out

    def _combined_text(self) -> str:
        return '\n\n'.join(t for t in self._ocr_texts if t)

    # ── 분석 ────────────────────────────────────────────────────
    def _analyze(self):
        """개인정보 보호 모드면 아직 안 읽은 캡처의 글자를 PC에서 읽은 뒤, 모은 글자를 AI로 보낸다.
        보호 모드가 꺼져 있으면 이어 붙인 이미지를 보낸다."""
        from ai_client import get_adapter
        from school_profile import privacy_enabled
        self._gen += 1
        gen = self._gen
        self._set_busy(True)
        if get_adapter() is not None and privacy_enabled():
            self._ocr_pending(gen)
        else:
            self._run_ai(gen, use_text=False)

    def _ocr_pending(self, gen):
        todo = [i for i, t in enumerate(self._ocr_texts) if t is None]
        if not todo:
            self._run_ai(gen, use_text=True)
            return
        try:
            import winrt  # noqa: F401  (Windows OCR)
        except ImportError:
            self._fail('Windows OCR을 쓸 수 없어 글자를 읽지 못했습니다.\n'
                       '메뉴 → 설정 → 개인정보 보호 모드를 끄면 이미지로 분석합니다.', gen)
            return
        from capture import run_ocr
        i = todo[0]
        n = len(self._pixmaps)
        self._set_status('PC에서 글자를 읽는 중… (개인정보 보호 모드)' + (f' — 캡처 {i + 1}/{n}' if n > 1 else ''))
        pixmap = self._pixmaps[i]

        def done(text):
            if self._closed or gen != self._gen:
                return
            if i < len(self._ocr_texts) and self._pixmaps[i] is pixmap:
                self._ocr_texts[i] = text.strip()
            self._ocr_pending(gen)

        def fail(msg):
            if not self._closed and gen == self._gen:
                self._fail(msg, gen)
        w = run_ocr(pixmap, done, fail)
        _ALIVE_WORKERS.add(w)
        w.finished.connect(lambda *_: _ALIVE_WORKERS.discard(w))

    def _run_ai(self, gen, use_text: bool):
        n = len(self._pixmaps)
        joined = f' (캡처 {n}장을 합쳐서)' if n > 1 else ''
        self._set_status(f'AI가 공문을 읽는 중…{joined}')
        text = self._combined_text() if use_text else ''
        if use_text and not text:
            self._fail('글자를 읽지 못했습니다.', gen)
            return
        w = doc_ai.DocAnalysisWorker(pixmap=None if use_text else self._combined_pixmap(), ocr_text=text)
        _ALIVE_WORKERS.add(w)
        w.finished.connect(lambda: _ALIVE_WORKERS.discard(w))   # QThread 종료 시 정리
        w.progress.connect(lambda m, g=gen: g == self._gen and not self._closed and self._set_status(m + joined))
        w.analyzed.connect(lambda a, info, g=gen: self._on_done(a, info, g))
        w.failed.connect(lambda m, g=gen: self._fail(m, g))
        self._worker = w
        w.start()

    def _on_done(self, a: dict, info: dict, gen=None):
        if self._closed or (gen is not None and gen != self._gen):
            return
        self.fill(a, merge=True)
        self._set_busy(False)
        docno = doc_ai.related_line(self._analysis)
        self.analyzed.emit(self._analysis.get('title', '').strip(), docno)
        n = len(self._pixmaps)
        msg = f'캡처 {n}장을 합쳐서 분석했어요.' if n > 1 else '분석 완료. 제목·관련 번호는 오른쪽 칸에 넣었어요.'
        if self._analysis.get('needs_reply'):
            msg += '\n회신이 필요한 공문으로 보여요.'
        if not docno:
            msg += '\n공문번호를 찾지 못했어요. 번호 부분만 한 번 더 캡처하면 합쳐서 읽어요.'
        elif '(' not in docno:
            msg += '\n공문번호의 날짜를 찾지 못했어요. 날짜 부분을 한 번 더 캡처하거나 직접 넣어 주세요.'
        if info.get('mode') == 'text':
            masked = describe(info.get('masked') or {})
            msg += ('\n🛡 개인정보 보호: 이미지는 보내지 않고 글자만 보냈습니다.'
                    + (f' ({masked} 가림)' if masked else ''))
        self._set_status(msg, 'done')

    def _fail(self, msg: str, gen=None):
        if self._closed or (gen is not None and gen != self._gen):
            return
        self._set_busy(False)
        self._set_status(f'분석 실패: {msg}\n다시 캡처하거나 요약을 직접 적어도 됩니다.', 'error')
        text = self._combined_text()
        if text and not self._analysis.get('title'):
            first = next((l.strip() for l in text.splitlines() if l.strip()), '')
            if first:
                self.analyzed.emit(first[:80], '')     # 읽은 첫 줄을 제목 칸에 (비어 있을 때만 들어감)

    # ── 값 채우기 / 읽기 ─────────────────────────────────────────
    def fill(self, a: dict, merge: bool = False):
        """분석 결과를 담고 요약 칸을 채운다. merge=True(이어서 캡처)면 지금까지의 결과와 합치고,
        사용자가 고친 요약은 그대로 둔다."""
        self._has_content = True
        self.btn_new.show()
        self._analysis = _merge(self._analysis, a) if merge else dict(a)
        if merge and self._summary_edited:
            return
        self._filling = True
        try:
            self.ed_summary.setPlainText(doc_ai.summary_body(self._analysis))
        finally:
            self._filling = False
        self._summary_edited = False

    def header(self):
        """(제목, 관련 번호): 공문 작성 창의 칸 값, 없으면 분석 결과."""
        if self.header_source is not None:
            title, rel = self.header_source()
            if title or rel:
                return title, rel
        return self._analysis.get('title', ''), doc_ai.related_line(self._analysis) if self._analysis else ''

    def context(self) -> str:
        """AI 대화·자동 초안에 함께 보낼 받은 공문 정리."""
        a = self._analysis
        return doc_ai.context_block(a.get('title', ''), doc_ai.related_line(a) if a else '', self.summary())

    # ── 동작 ────────────────────────────────────────────────────
    def _copy_summary(self):
        title, rel = self.header()
        text = doc_ai.summary_text(title, rel, self.summary())
        if text:
            QApplication.clipboard().setText(text)
            self.message.emit('요약을 복사했습니다.')
        else:
            self.message.emit('복사할 요약이 없습니다.')

    def shutdown(self):
        """창을 닫을 때: 아직 도는 분석이 닫힌 창을 건드리지 않도록 연결만 끊는다."""
        self._closed = True
        if self._worker is not None:
            for sig in (self._worker.progress, self._worker.analyzed, self._worker.failed):
                try:
                    sig.disconnect()
                except TypeError:
                    pass
