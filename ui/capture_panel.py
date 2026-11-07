"""ui/capture_panel.py — 공문 작성 창 왼쪽의 「받은 공문」 칸.

캡처한 공문을 AI로 읽어 업무명·공문번호(날짜 포함)·기한·요약·할 일을 보여주고,
[업무로 등록] [회신 공문 쓰기] [요약 복사], 그리고 제목·관련 번호 칸에 넣기를 할 수 있다.
AI가 실패해도 칸을 직접 채워 등록할 수 있다. 모양(QSS)은 공문 작성 창에서 물려받는다.

이어서 캡처: 제목·본문과 공문번호가 한 화면에 안 잡히면 두 번 캡처한다. [새 공문]을 누르기 전까지
캡처는 같은 공문에 덧붙고, 모은 캡처를 합쳐(이미지는 위아래로 이어 붙이고, 보호 모드면 읽은 글자를
이어 붙여) 다시 분석한다. 사용자가 직접 고친 칸은 다시 분석해도 덮어쓰지 않는다.
"""

from PyQt5.QtCore import QDate, QLocale, QSize, Qt, QTime, pyqtSignal
from PyQt5.QtGui import QColor, QPainter, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QCheckBox, QDateEdit, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QPushButton, QScrollArea, QTimeEdit, QVBoxLayout, QWidget,
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


class CapturePanel(QFrame):
    """시그널로 결과를 넘긴다.

    register_requested(dict)  — 업무로 등록 (collect() 결과)
    reply_requested(dict)     — 회신 공문 쓰기
    use_title(str) / use_ref(str) — 공문 작성 창의 제목 / 관련 번호 칸에 넣기
    collapse_requested()      — 칸 접기
    message(str)              — 상태 표시줄에 보일 안내
    analyzed(str, str)        — 분석이 끝남: (제목, 공문번호) → 오른쪽 제목·관련 번호 칸에도 넣는다
    """
    register_requested = pyqtSignal(dict)
    reply_requested = pyqtSignal(dict)
    use_title = pyqtSignal(str)
    use_ref = pyqtSignal(str)
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
        self._user_edited = set() # 사용자가 직접 고친 칸 — 다시 분석해도 유지
        self._filling = False
        self._worker = None
        self._closed = False
        self._has_content = False
        self._needs_reply = False
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

    def _label(self, text):
        lb = QLabel(text)
        lb.setObjectName('FieldLabel')
        return lb

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        inner = QWidget()
        inner.setObjectName('PanelInner')
        self._inner = inner
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)
        footer = QWidget()
        footer.setObjectName('PanelFooter')
        foot = QVBoxLayout(footer)
        foot.setContentsMargins(16, 10, 16, 12)
        foot.setSpacing(6)
        outer.addWidget(footer)

        root = QVBoxLayout(inner)
        root.setContentsMargins(16, 14, 16, 10)
        root.setSpacing(6)

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

        # 업무명 + 제목 칸에 넣기
        root.addWidget(self._label('업무명 (받은 공문 제목)'))
        self.ed_title = QLineEdit()
        self.ed_title.setPlaceholderText('업무명')
        root.addWidget(self.ed_title)
        # 공문번호 + 관련 칸에 넣기
        root.addWidget(self._label('공문번호'))
        self.ed_docno = QLineEdit()
        self.ed_docno.setPlaceholderText('예: 중등교육과-12345(2026. 11. 3.)')
        self.ed_docno.editingFinished.connect(self._tidy_docno)
        root.addWidget(self.ed_docno)
        put = QHBoxLayout()
        put.setSpacing(4)
        self.btn_to_title = self._small('제목 칸에', 'fa5s.level-up-alt', '업무명을 오른쪽 제목 칸에 넣기')
        self.btn_to_ref = self._small('관련 번호 칸에', 'fa5s.level-up-alt', '공문번호를 오른쪽 관련 번호 칸에 넣기')
        put.addWidget(self.btn_to_title)
        put.addWidget(self.btn_to_ref)
        put.addStretch()
        root.addLayout(put)

        # 기한
        self.chk_deadline = QCheckBox('기한')
        self.chk_deadline.setObjectName('FieldCheck')
        root.addWidget(self.chk_deadline)
        dl = QHBoxLayout()
        dl.setSpacing(6)
        self.ed_date = QDateEdit(QDate.currentDate())
        self.ed_date.setCalendarPopup(True)
        self.ed_date.setLocale(QLocale(QLocale.Korean, QLocale.SouthKorea))   # 요일을 (월)(화)… 로
        self.ed_date.setDisplayFormat('yyyy. M. d. (ddd)')
        self.chk_time = QCheckBox('시각')
        self.ed_time = QTimeEdit(QTime(17, 0))
        self.ed_time.setDisplayFormat('HH:mm')
        self.ed_time.setButtonSymbols(QTimeEdit.NoButtons)
        self.ed_time.setFixedWidth(64)
        dl.addWidget(self.ed_date, 1)
        dl.addWidget(self.chk_time)
        dl.addWidget(self.ed_time)
        self.chk_deadline.toggled.connect(self._sync_deadline_enabled)
        self.chk_time.toggled.connect(self._sync_deadline_enabled)
        root.addLayout(dl)

        root.addWidget(self._label('요약'))
        self.ed_summary = QPlainTextEdit()
        self.ed_summary.setPlaceholderText('핵심 내용 (한 줄에 하나)')
        self.ed_summary.setFixedHeight(66)
        root.addWidget(self.ed_summary)
        root.addWidget(self._label('할 일'))
        self.ed_todos = QPlainTextEdit()
        self.ed_todos.setPlaceholderText('해야 할 일 (한 줄에 하나)')
        self.ed_todos.setFixedHeight(76)
        root.addWidget(self.ed_todos)

        self.lbl_privacy = QLabel('')
        self.lbl_privacy.setObjectName('Hint')
        self.lbl_privacy.setWordWrap(True)
        self.lbl_privacy.hide()
        root.addWidget(self.lbl_privacy)

        self.chk_context = QCheckBox('AI 대화·자동 초안에서 이 공문 참고')
        self.chk_context.setChecked(True)
        self.chk_context.setToolTip('켜 두면 오른쪽 AI 대화와 자동 초안이 이 공문의 요약·할 일을 함께 보고 씁니다')
        foot.addWidget(self.chk_context)

        root.addStretch()
        self.btn_register = QPushButton('업무로 등록')
        self.btn_register.setObjectName('Primary')
        self.btn_register.setIcon(_icon('fa5s.check', 'white'))
        self.btn_register.setIconSize(QSize(12, 12))
        self.btn_register.setMinimumHeight(34)
        foot.addWidget(self.btn_register)
        row = QHBoxLayout()
        row.setSpacing(6)
        self.btn_reply = QPushButton('회신 공문 쓰기')
        self.btn_reply.setIcon(_icon('fa5s.reply', self.MUTED))
        self.btn_reply.setToolTip('오른쪽에 회신 제목·관련 번호를 넣고 AI가 본문 초안을 씁니다')
        self.btn_copy = QPushButton('요약 복사')
        self.btn_copy.setIcon(_icon('fa5.copy', self.MUTED))
        for b in (self.btn_reply, self.btn_copy):
            b.setMinimumHeight(32)
            row.addWidget(b, 1)
        foot.addLayout(row)
        extra = QHBoxLayout()
        extra.setSpacing(4)
        self.btn_retry = self._small('다시 분석', 'fa5s.redo-alt')
        self.btn_copy_raw = self._small('원문 복사', 'fa5.file-alt', 'PC에서 읽은 글자 그대로 복사')
        extra.addWidget(self.btn_retry)
        extra.addWidget(self.btn_copy_raw)
        extra.addStretch()
        root.addLayout(extra)

        self.btn_to_title.clicked.connect(lambda: self.use_title.emit(self.ed_title.text().strip()))
        self.btn_to_ref.clicked.connect(lambda: self.use_ref.emit(self._tidy_docno()))
        self.btn_register.clicked.connect(self._register)
        self.btn_reply.clicked.connect(lambda: self.reply_requested.emit(self.collect()))
        self.btn_copy.clicked.connect(lambda: self._copy(doc_ai.summary_text(self.collect()), '요약을 복사했습니다.'))
        self.btn_copy_raw.clicked.connect(lambda: self._copy(self._combined_text(), '원문 글자를 복사했습니다.'))
        self.btn_retry.clicked.connect(self._retry)
        self.ed_title.textChanged.connect(lambda *_: self._edited('title'))
        self.ed_docno.textChanged.connect(lambda *_: self._edited('doc_number'))
        self.ed_summary.textChanged.connect(lambda: self._edited('summary'))
        self.ed_todos.textChanged.connect(lambda: self._edited('todos'))
        for w in (self.chk_deadline, self.chk_time):
            w.toggled.connect(lambda *_: self._edited('deadline'))
        self.ed_date.dateChanged.connect(lambda *_: self._edited('deadline'))
        self.ed_time.timeChanged.connect(lambda *_: self._edited('deadline'))

    def content_min_width(self) -> int:
        """칸 안의 글자·버튼이 잘리지 않는 최소 폭 (배율·글꼴에 따라 달라짐)."""
        footer = self.findChild(QWidget, 'PanelFooter')
        need = self._inner.minimumSizeHint().width() + 10          # 세로 스크롤바 자리
        if footer is not None:
            need = max(need, footer.minimumSizeHint().width())
        return need

    # ── 상태 ────────────────────────────────────────────────────
    def has_content(self) -> bool:
        return self._has_content

    def use_as_context(self) -> bool:
        return self._has_content and self.chk_context.isChecked() and bool(self.ed_title.text().strip())

    def capture_count(self) -> int:
        return len(self._pixmaps)

    def clear(self):
        self._gen += 1            # 진행 중인 분석 결과는 버린다
        self._has_content = False
        self._pixmaps, self._ocr_texts = [], []
        self._filling = True
        for w in (self.ed_title, self.ed_docno):
            w.clear()
        self.ed_summary.clear()
        self.ed_todos.clear()
        self.chk_deadline.setChecked(False)
        self.chk_time.setChecked(False)
        self.lbl_privacy.hide()
        self.btn_retry.hide()
        self.btn_copy_raw.hide()
        self._filling = False
        self._user_edited = set()
        self._needs_reply = False
        self._show_thumbs()
        self._set_status('')
        self._set_busy(False)
        self._reset_register()
        self.btn_new.hide()

    def new_document(self):
        """[새 공문]: 모은 캡처와 칸을 비운다. 다음 캡처부터 새 공문."""
        self.clear()
        self._set_status('비웠습니다. 새 공문을 캡처하세요 (Ctrl+Shift+X).')
        self.message.emit('받은 공문 칸을 비웠습니다.')

    def _fields(self):
        return (self.ed_title, self.ed_docno, self.ed_summary, self.ed_todos, self.chk_deadline, self.chk_time)

    def _set_busy(self, busy: bool):
        for w in self._fields():
            w.setEnabled(not busy)
        self._sync_deadline_enabled()
        for b in (self.btn_register, self.btn_reply, self.btn_copy, self.btn_to_title, self.btn_to_ref):
            b.setEnabled(not busy)
        if not busy:
            self._sync_buttons()

    def _set_status(self, text, state=''):
        self.lbl_status.setText(text)
        self.lbl_status.setProperty('state', state)
        self.lbl_status.style().unpolish(self.lbl_status)
        self.lbl_status.style().polish(self.lbl_status)

    def _sync_deadline_enabled(self):
        on = self.chk_deadline.isChecked() and self.chk_deadline.isEnabled()
        self.ed_date.setEnabled(on)
        self.chk_time.setEnabled(on)
        self.ed_time.setEnabled(on and self.chk_time.isChecked())

    def _sync_buttons(self):
        has_title = bool(self.ed_title.text().strip())
        self.btn_register.setEnabled(has_title and not getattr(self, '_registered', False))
        self.btn_reply.setEnabled(has_title)
        self.btn_to_title.setEnabled(has_title)
        self.btn_to_ref.setEnabled(bool(self.ed_docno.text().strip()))

    def _reset_register(self):
        self._registered = False
        self.btn_register.setText('업무로 등록')
        self._sync_buttons()

    def _edited(self, field=''):
        """사용자가 고친 칸은 기억해 두고(다시 분석해도 유지), 등록한 뒤 고치면 다시 등록할 수 있게."""
        if self._filling:
            return
        if field:
            self._user_edited.add(field)
        if getattr(self, '_registered', False):
            self._reset_register()
        else:
            self._sync_buttons()

    def _tidy_docno(self) -> str:
        """공문번호를 '중등교육과-12345(2026. 11. 3.)' 형식으로 맞춘다."""
        v = doc_ai.format_doc_number(self.ed_docno.text())
        if v != self.ed_docno.text():
            self.ed_docno.setText(v)
        return v

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
    def _retry(self):
        self._analyze()

    def _analyze(self):
        """개인정보 보호 모드면 아직 안 읽은 캡처의 글자를 PC에서 읽은 뒤, 모은 글자를 AI로 보낸다.
        보호 모드가 꺼져 있으면 이어 붙인 이미지를 보낸다."""
        from ai_client import get_adapter
        from school_profile import privacy_enabled
        self._gen += 1
        gen = self._gen
        self.btn_retry.hide()
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
            self.btn_copy_raw.setVisible(bool(self._combined_text()))
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
        self.analyzed.emit(self.ed_title.text().strip(), self.ed_docno.text().strip())
        n = len(self._pixmaps)
        msg = f'캡처 {n}장을 합쳐서 분석했어요.' if n > 1 else '분석 완료. 고칠 곳은 고쳐서 쓰세요.'
        if self._needs_reply:
            msg += '\n회신이 필요한 공문으로 보여요.'
        docno = self.ed_docno.text().strip()
        if not docno:
            msg += '\n공문번호를 찾지 못했어요. 번호 부분만 한 번 더 캡처하면 합쳐서 읽어요.'
        elif '(' not in docno:
            msg += '\n공문번호의 날짜를 찾지 못했어요. 날짜 부분을 한 번 더 캡처하거나 직접 넣어 주세요.'
        self._set_status(msg, 'done')
        if info.get('mode') == 'text':
            masked = describe(info.get('masked') or {})
            self.lbl_privacy.setText('🛡 개인정보 보호: 이미지는 보내지 않고 글자만 보냈습니다.'
                                     + (f' ({masked} 가림)' if masked else ''))
            self.lbl_privacy.show()

    def _fail(self, msg: str, gen=None):
        if self._closed or (gen is not None and gen != self._gen):
            return
        self._set_busy(False)
        self._set_status(f'분석 실패: {msg}\n칸을 직접 채워 쓸 수도 있습니다.', 'error')
        self.btn_retry.show()
        text = self._combined_text()
        if text and not self.ed_title.text().strip():
            first = next((l.strip() for l in text.splitlines() if l.strip()), '')
            self._filling = True
            self.ed_title.setText(first[:80])
            self._filling = False

    # ── 값 채우기 / 읽기 ─────────────────────────────────────────
    def fill(self, a: dict, merge: bool = False):
        """분석 결과를 칸에 넣는다. merge=True(이어서 캡처)면 사용자가 고친 칸은 그대로 두고,
        새 결과가 비어 있는 칸은 지금 값을 유지한다 (예: 공문번호만 캡처해도 제목·요약은 남음)."""
        self._has_content = True
        self.btn_new.show()
        keep = self._user_edited if merge else set()
        self._filling = True
        try:
            def put(field, value, setter, current):
                if field in keep or (merge and not value and current):
                    return
                setter(value)
            put('title', a.get('title', ''), self.ed_title.setText, self.ed_title.text().strip())
            put('doc_number', doc_ai.related_line(a), self.ed_docno.setText, self.ed_docno.text().strip())
            put('summary', '\n'.join(a.get('summary', [])), self.ed_summary.setPlainText,
                self.ed_summary.toPlainText().strip())
            put('todos', '\n'.join(a.get('todos', [])), self.ed_todos.setPlainText, self.ed_todos.toPlainText().strip())
            self._needs_reply = bool(a.get('needs_reply', False)) or (merge and self._needs_reply)
            d = QDate.fromString(a.get('deadline', ''), 'yyyy-MM-dd')
            if 'deadline' not in keep and (d.isValid() or not (merge and self.chk_deadline.isChecked())):
                self.chk_deadline.setChecked(d.isValid())
                if d.isValid():
                    self.ed_date.setDate(d)
                t = QTime.fromString(a.get('deadline_time', ''), 'HH:mm')
                self.chk_time.setChecked(t.isValid())
                if t.isValid():
                    self.ed_time.setTime(t)
        finally:
            self._filling = False
        self._sync_deadline_enabled()
        if not (merge and getattr(self, '_registered', False)):    # 이미 등록했으면 '등록됨' 유지 (중복 방지)
            self._reset_register()
        else:
            self._sync_buttons()

    def collect(self) -> dict:
        """사용자가 고친 값까지 반영한 분석 결과. doc_number는 날짜까지 붙은 형식."""
        lines = lambda w: [l.strip() for l in w.toPlainText().splitlines() if l.strip()]  # noqa: E731
        has_dl = self.chk_deadline.isChecked()
        return {
            'title': self.ed_title.text().strip(),
            'doc_number': doc_ai.format_doc_number(self.ed_docno.text()),
            'doc_date': '',
            'deadline': self.ed_date.date().toString('yyyy-MM-dd') if has_dl else '',
            'deadline_time': self.ed_time.time().toString('HH:mm') if has_dl and self.chk_time.isChecked() else '',
            'summary': lines(self.ed_summary),
            'todos': lines(self.ed_todos),
            'needs_reply': self._needs_reply,
        }

    # ── 동작 ────────────────────────────────────────────────────
    def _register(self):
        a = self.collect()
        if not a['title']:
            return
        self.register_requested.emit(a)
        self._registered = True
        self.btn_register.setText('업무로 등록됨')
        self.btn_register.setEnabled(False)

    def _copy(self, text, msg):
        if text:
            QApplication.clipboard().setText(text)
            self.message.emit(msg)

    def shutdown(self):
        """창을 닫을 때: 아직 도는 분석이 닫힌 창을 건드리지 않도록 연결만 끊는다."""
        self._closed = True
        if self._worker is not None:
            for sig in (self._worker.progress, self._worker.analyzed, self._worker.failed):
                try:
                    sig.disconnect()
                except TypeError:
                    pass
