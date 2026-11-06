"""ui/capture_panel.py — 공문 작성 창 왼쪽의 「받은 공문」 칸.

캡처한 공문을 AI로 읽어 업무명·공문번호(날짜 포함)·기한·요약·할 일을 보여주고,
[업무로 등록] [회신 공문 쓰기] [요약 복사], 그리고 제목·관련 번호 칸에 넣기를 할 수 있다.
AI가 실패해도 칸을 직접 채워 등록할 수 있다. 모양(QSS)은 공문 작성 창에서 물려받는다.
"""

from PyQt5.QtCore import QDate, QLocale, QSize, Qt, QTime, pyqtSignal
from PyQt5.QtGui import QPixmap
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
    """
    register_requested = pyqtSignal(dict)
    reply_requested = pyqtSignal(dict)
    use_title = pyqtSignal(str)
    use_ref = pyqtSignal(str)
    collapse_requested = pyqtSignal()
    message = pyqtSignal(str)

    MUTED = '#64748B'
    ACCENT = '#4F46E5'

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('Panel')
        self._pixmap = None
        self._ocr_text = ''
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
        self.btn_collapse = self._small('', 'fa5s.angle-double-left', '받은 공문 칸 접기')
        self.btn_collapse.setFixedSize(28, 26)
        self.btn_collapse.clicked.connect(self.collapse_requested)
        head.addWidget(self.btn_collapse)
        root.addLayout(head)

        top = QHBoxLayout()
        top.setSpacing(10)
        self.thumb = QLabel()
        self.thumb.setObjectName('Thumb')
        self.thumb.setFixedSize(84, 62)
        self.thumb.setAlignment(Qt.AlignCenter)
        top.addWidget(self.thumb)
        self.lbl_status = QLabel('')
        self.lbl_status.setObjectName('Status')
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        top.addWidget(self.lbl_status, 1)
        root.addLayout(top)

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
        self.btn_copy_raw.clicked.connect(lambda: self._copy(self._ocr_text, '원문 글자를 복사했습니다.'))
        self.btn_retry.clicked.connect(self._retry)
        for w in (self.ed_title, self.ed_docno):
            w.textChanged.connect(self._edited)
        for w in (self.ed_summary, self.ed_todos):
            w.textChanged.connect(self._edited)
        for w in (self.chk_deadline, self.chk_time):
            w.toggled.connect(self._edited)
        self.ed_date.dateChanged.connect(self._edited)
        self.ed_time.timeChanged.connect(self._edited)

    # ── 상태 ────────────────────────────────────────────────────
    def has_content(self) -> bool:
        return self._has_content

    def use_as_context(self) -> bool:
        return self._has_content and self.chk_context.isChecked() and bool(self.ed_title.text().strip())

    def clear(self):
        self._has_content = False
        self._ocr_text = ''
        for w in (self.ed_title, self.ed_docno):
            w.clear()
        self.ed_summary.clear()
        self.ed_todos.clear()
        self.chk_deadline.setChecked(False)
        self.chk_time.setChecked(False)
        self.lbl_privacy.hide()
        self.btn_retry.hide()
        self.btn_copy_raw.hide()
        self.thumb.clear()
        self._set_status('')
        self._sync_deadline_enabled()
        self._reset_register()

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
        if not getattr(self, '_registered', False):
            self.btn_register.setEnabled(has_title)
        self.btn_reply.setEnabled(has_title)
        self.btn_to_title.setEnabled(has_title)
        self.btn_to_ref.setEnabled(bool(self.ed_docno.text().strip()))

    def _reset_register(self):
        self._registered = False
        self.btn_register.setText('업무로 등록')
        self._sync_buttons()

    def _edited(self, *_):
        """등록한 뒤 내용을 고치면 다시 등록할 수 있게."""
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

    # ── 분석 ────────────────────────────────────────────────────
    def start(self, pixmap: QPixmap):
        """새 캡처 분석. 개인정보 보호 모드면 PC에서 글자를 먼저 읽고, 아니면 이미지를 바로 보낸다."""
        self.clear()
        self._pixmap = pixmap
        self._has_content = True
        if pixmap is not None and not pixmap.isNull():
            self.thumb.setPixmap(pixmap.scaled(84, 62, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self._analyze()

    def _retry(self):
        self._analyze()

    def _analyze(self):
        from ai_client import get_adapter
        from school_profile import privacy_enabled
        self.btn_retry.hide()
        self._set_busy(True)
        adapter = get_adapter()
        if adapter is not None and privacy_enabled() and not self._ocr_text:
            self._set_status('PC에서 글자를 읽는 중… (개인정보 보호 모드)')
            self._run_ocr()
        else:
            self._run_ai()

    def _run_ocr(self):
        try:
            import winrt  # noqa: F401  (Windows OCR)
        except ImportError:
            self._fail('Windows OCR을 쓸 수 없어 글자를 읽지 못했습니다.\n'
                       '메뉴 → 🛡️ 개인정보 보호 모드를 끄면 이미지로 분석합니다.')
            return
        from capture import run_ocr

        def done(text):
            if self._closed:
                return
            self._ocr_text = text.strip()
            self.btn_copy_raw.setVisible(bool(self._ocr_text))
            self._run_ai()

        def fail(msg):
            if not self._closed:
                self._fail(msg)
        w = run_ocr(self._pixmap, done, fail)
        _ALIVE_WORKERS.add(w)
        w.finished.connect(lambda *_: _ALIVE_WORKERS.discard(w))

    def _run_ai(self):
        self._set_status('AI가 공문을 읽는 중…')
        w = doc_ai.DocAnalysisWorker(pixmap=None if self._ocr_text else self._pixmap, ocr_text=self._ocr_text)
        _ALIVE_WORKERS.add(w)
        w.finished.connect(lambda: _ALIVE_WORKERS.discard(w))   # QThread 종료 시 정리
        w.progress.connect(self._set_status)
        w.analyzed.connect(self._on_done)
        w.failed.connect(self._fail)
        self._worker = w
        w.start()

    def _on_done(self, a: dict, info: dict):
        if self._closed:
            return
        self.fill(a)
        self._set_busy(False)
        msg = '분석 완료. 고칠 곳은 고쳐서 쓰세요.'
        if a.get('needs_reply'):
            msg += '\n회신이 필요한 공문으로 보여요.'
        if a.get('doc_number') and '(' not in a['doc_number']:
            msg += '\n공문번호의 날짜를 찾지 못했습니다. 직접 넣어 주세요.'
        self._set_status(msg, 'done')
        if info.get('mode') == 'text':
            masked = describe(info.get('masked') or {})
            self.lbl_privacy.setText('🛡 개인정보 보호: 이미지는 보내지 않고 글자만 보냈습니다.'
                                     + (f' ({masked} 가림)' if masked else ''))
            self.lbl_privacy.show()

    def _fail(self, msg: str):
        if self._closed:
            return
        self._set_busy(False)
        self._set_status(f'분석 실패: {msg}\n칸을 직접 채워 쓸 수도 있습니다.', 'error')
        self.btn_retry.show()
        if self._ocr_text and not self.ed_title.text().strip():
            first = next((l.strip() for l in self._ocr_text.splitlines() if l.strip()), '')
            self.ed_title.setText(first[:80])

    # ── 값 채우기 / 읽기 ─────────────────────────────────────────
    def fill(self, a: dict):
        self._has_content = True
        self.ed_title.setText(a.get('title', ''))
        self.ed_docno.setText(doc_ai.related_line(a))
        self.ed_summary.setPlainText('\n'.join(a.get('summary', [])))
        self.ed_todos.setPlainText('\n'.join(a.get('todos', [])))
        self._needs_reply = a.get('needs_reply', False)
        d = QDate.fromString(a.get('deadline', ''), 'yyyy-MM-dd')
        self.chk_deadline.setChecked(d.isValid())
        if d.isValid():
            self.ed_date.setDate(d)
        t = QTime.fromString(a.get('deadline_time', ''), 'HH:mm')
        self.chk_time.setChecked(t.isValid())
        if t.isValid():
            self.ed_time.setTime(t)
        self._sync_deadline_enabled()
        self._reset_register()

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
        self.btn_register.setText('✓ 업무로 등록됨')
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
