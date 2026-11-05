"""ui/capture_dialog.py — 「AI 공문 캡처」 결과 창.

캡처한 공문을 AI로 분석해 업무명·마감일·공문번호·요약·할 일·붙임을 채워 보여주고,
[업무로 등록] [회신 공문 작성] [요약 복사] [원문 복사] 중 하나를 고르게 한다.
AI가 실패해도 칸을 직접 채워 업무로 등록할 수 있다.
"""

from PyQt5.QtCore import Qt, QDate, QLocale, QTime
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (
    QApplication, QCheckBox, QDateEdit, QDialog, QFormLayout, QHBoxLayout, QLabel,
    QLineEdit, QPlainTextEdit, QPushButton, QTimeEdit, QVBoxLayout, QWidget,
)

import doc_ai
from privacy import describe
from .common import pr_font
from .theme import T

# 창을 닫아도 분석 스레드가 끝날 때까지 참조를 유지 (실행 중인 QThread가 파괴되면 앱이 종료됨)
_ALIVE_WORKERS = set()

_STYLE = f"""
    QDialog {{ background: {T.BG_WARM}; }}
    QLabel {{ color: {T.TEXT_WARM}; background: transparent; }}
    QLabel#Title {{ color: {T.INK}; font-size: 13pt; font-weight: 700; }}
    QLabel#Status {{ color: {T.TEXT_WARM_MUTED}; }}
    QLabel#Status[state="error"] {{ color: {T.DANGER}; }}
    QLabel#Status[state="done"] {{ color: {T.TEXT_ON_NOTE}; }}
    QLabel#Privacy {{ color: {T.TEXT_WARM_MUTED}; font-size: 9pt; }}
    QLineEdit, QPlainTextEdit, QDateEdit, QTimeEdit {{
        background: white; border: 1px solid {T.SURFACE_WARM}; border-radius: 4px;
        padding: 3px 6px; color: {T.INK};
    }}
    QLineEdit:focus, QPlainTextEdit:focus, QDateEdit:focus, QTimeEdit:focus {{ border: 1px solid {T.ACCENT}; }}
    QLineEdit:disabled, QPlainTextEdit:disabled {{ background: {T.BG_WARM_ALT}; color: {T.TEXT_WARM_MUTED}; }}
    QPushButton {{
        background: {T.SURFACE_WARM}; color: {T.TEXT_ON_NOTE}; border: none; border-radius: 5px;
        padding: 6px 12px;
    }}
    QPushButton:hover {{ background: {T.SURFACE_WARM_HOVER}; }}
    QPushButton:pressed {{ background: {T.SURFACE_WARM_PRESSED}; }}
    QPushButton:disabled {{ color: {T.TEXT_WARM_MUTED}; }}
    QPushButton#Primary {{ background: {T.ACCENT}; color: white; font-weight: 700; }}
    QPushButton#Primary:hover {{ background: {T.ACCENT_HOVER}; }}
"""


class CaptureResultDialog(QDialog):
    """on_register(name, deadline, notes, related_no) / on_reply(analysis) 콜백으로 결과를 넘긴다."""

    def __init__(self, pixmap: QPixmap, on_register, on_reply=None, parent=None, auto_start=True):
        super().__init__(parent, Qt.Window | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setWindowTitle('AI 공문 캡처')
        self.setStyleSheet(_STYLE)
        self.setFont(pr_font(10))
        self.setMinimumWidth(480)
        self._pixmap = pixmap
        self._on_register = on_register
        self._on_reply = on_reply
        self._ocr_text = ''
        self._worker = None
        self._ocr_worker = None
        self._closed = False
        self._build()
        if auto_start:
            self.start()

    # ── 화면 구성 ────────────────────────────────────────────────
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 14)
        root.setSpacing(10)

        head = QHBoxLayout()
        thumb = QLabel()
        if self._pixmap is not None and not self._pixmap.isNull():
            thumb.setPixmap(self._pixmap.scaled(96, 72, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        thumb.setFixedSize(96, 72)
        thumb.setAlignment(Qt.AlignCenter)
        thumb.setStyleSheet(f'background: white; border: 1px solid {T.SURFACE_WARM}; border-radius: 4px;')
        head.addWidget(thumb)
        titles = QVBoxLayout()
        lbl_title = QLabel('공문 분석')
        lbl_title.setObjectName('Title')
        self.lbl_status = QLabel('준비 중…')
        self.lbl_status.setObjectName('Status')
        self.lbl_status.setWordWrap(True)
        titles.addWidget(lbl_title)
        titles.addWidget(self.lbl_status)
        titles.addStretch()
        head.addLayout(titles, 1)
        root.addLayout(head)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(7)

        self.ed_title = QLineEdit()
        self.ed_title.setPlaceholderText('업무명 (공문 제목)')
        form.addRow('업무명', self.ed_title)

        dl = QWidget()
        dl_lay = QHBoxLayout(dl)
        dl_lay.setContentsMargins(0, 0, 0, 0)
        self.chk_deadline = QCheckBox('기한')
        self.ed_date = QDateEdit(QDate.currentDate())
        self.ed_date.setCalendarPopup(True)
        self.ed_date.setLocale(QLocale(QLocale.Korean, QLocale.SouthKorea))  # 요일을 (월)(화)… 로
        self.ed_date.setDisplayFormat('yyyy. M. d. (ddd)')
        self.chk_time = QCheckBox('시각')
        self.ed_time = QTimeEdit(QTime(17, 0))
        self.ed_time.setDisplayFormat('HH:mm')
        for w in (self.chk_deadline, self.ed_date, self.chk_time, self.ed_time):
            dl_lay.addWidget(w)
        dl_lay.addStretch()
        self.chk_deadline.toggled.connect(self._sync_deadline_enabled)
        self.chk_time.toggled.connect(self._sync_deadline_enabled)
        form.addRow('마감', dl)

        self.ed_docno = QLineEdit()
        self.ed_docno.setPlaceholderText('예: 경기도교육청 교육과정과-12345')
        form.addRow('공문번호', self.ed_docno)
        self.ed_docdate = QLineEdit()
        self.ed_docdate.setPlaceholderText('시행일 YYYY-MM-DD')
        self.ed_sender = QLineEdit()
        self.ed_sender.setPlaceholderText('발신 기관')
        row2 = QWidget()
        r2 = QHBoxLayout(row2)
        r2.setContentsMargins(0, 0, 0, 0)
        r2.addWidget(self.ed_sender, 2)
        r2.addWidget(self.ed_docdate, 1)
        form.addRow('발신·시행일', row2)

        self.ed_summary = QPlainTextEdit()
        self.ed_summary.setPlaceholderText('핵심 내용 (한 줄에 하나)')
        self.ed_summary.setFixedHeight(64)
        form.addRow('요약', self.ed_summary)
        self.ed_todos = QPlainTextEdit()
        self.ed_todos.setPlaceholderText('해야 할 일 (한 줄에 하나)')
        self.ed_todos.setFixedHeight(80)
        form.addRow('할 일', self.ed_todos)
        self.ed_attach = QLineEdit()
        self.ed_attach.setPlaceholderText('붙임 파일 (쉼표로 구분)')
        form.addRow('붙임', self.ed_attach)
        root.addLayout(form)

        self.lbl_privacy = QLabel('')
        self.lbl_privacy.setObjectName('Privacy')
        self.lbl_privacy.setWordWrap(True)
        self.lbl_privacy.hide()
        root.addWidget(self.lbl_privacy)

        btns = QHBoxLayout()
        self.btn_reply = QPushButton('✉ 회신 공문 작성')
        self.btn_copy = QPushButton('요약 복사')
        self.btn_copy_raw = QPushButton('원문 복사')
        self.btn_retry = QPushButton('다시 분석')
        self.btn_close = QPushButton('닫기')
        self.btn_register = QPushButton('업무로 등록')
        self.btn_register.setObjectName('Primary')
        self.btn_register.setDefault(True)
        for b in (self.btn_reply, self.btn_copy, self.btn_copy_raw, self.btn_retry):
            btns.addWidget(b)
        btns.addStretch()
        btns.addWidget(self.btn_close)
        btns.addWidget(self.btn_register)
        root.addLayout(btns)

        self.btn_reply.setVisible(self._on_reply is not None)
        self.btn_copy_raw.hide()
        self.btn_retry.hide()
        self.btn_reply.clicked.connect(self._reply)
        self.btn_copy.clicked.connect(lambda: self._copy(doc_ai.summary_text(self.collect()), '요약을 복사했습니다.'))
        self.btn_copy_raw.clicked.connect(lambda: self._copy(self._ocr_text, '원문 글자를 복사했습니다.'))
        self.btn_retry.clicked.connect(self.start)
        self.btn_close.clicked.connect(self.close)
        self.btn_register.clicked.connect(self._register)
        self.ed_title.textChanged.connect(self._sync_buttons)
        self._sync_deadline_enabled()
        self._sync_buttons()

    def _fields(self):
        return (self.ed_title, self.ed_docno, self.ed_docdate, self.ed_sender, self.ed_summary,
                self.ed_todos, self.ed_attach, self.chk_deadline, self.chk_time)

    def _set_busy(self, busy: bool):
        for w in self._fields():
            w.setEnabled(not busy)
        self._sync_deadline_enabled()
        for b in (self.btn_register, self.btn_reply, self.btn_copy):
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
        self.btn_register.setEnabled(has_title)
        self.btn_reply.setEnabled(has_title)

    # ── 분석 ────────────────────────────────────────────────────
    def start(self):
        """개인정보 보호 모드 + 외부 AI면 글자를 먼저 읽고, 아니면 이미지를 바로 분석."""
        from ai_client import get_adapter, is_local_adapter
        from school_profile import privacy_enabled
        self.btn_retry.hide()
        self._set_busy(True)
        adapter = get_adapter()
        need_text = adapter is not None and privacy_enabled() and not is_local_adapter(adapter)
        if need_text and not self._ocr_text:
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
        self._ocr_worker = w

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
        self.fill(a)
        self._set_busy(False)
        msg = '분석 완료 — 내용을 확인하고 고칠 곳이 있으면 고친 뒤 등록하세요.'
        if a.get('needs_reply'):
            msg += ' (회신이 필요한 공문으로 보입니다)'
        self._set_status(msg, 'done')
        if info.get('mode') == 'text':
            masked = describe(info.get('masked') or {})
            self.lbl_privacy.setText('🛡 개인정보 보호: 이미지는 보내지 않고 글자만 보냈습니다.'
                                     + (f' ({masked} 가림)' if masked else ''))
            self.lbl_privacy.show()

    def _fail(self, msg: str):
        self._set_busy(False)
        self._set_status(f'분석 실패: {msg}\n칸을 직접 채워 등록할 수도 있습니다.', 'error')
        self.btn_retry.show()
        if self._ocr_text and not self.ed_title.text().strip():
            first = next((l.strip() for l in self._ocr_text.splitlines() if l.strip()), '')
            self.ed_title.setText(first[:80])

    # ── 값 채우기 / 읽기 ─────────────────────────────────────────
    def fill(self, a: dict):
        self.ed_title.setText(a.get('title', ''))
        self.ed_docno.setText(a.get('doc_number', ''))
        self.ed_docdate.setText(a.get('doc_date', ''))
        self.ed_sender.setText(a.get('sender', ''))
        self.ed_summary.setPlainText('\n'.join(a.get('summary', [])))
        self.ed_todos.setPlainText('\n'.join(a.get('todos', [])))
        self.ed_attach.setText(', '.join(a.get('attachments', [])))
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

    def collect(self) -> dict:
        """사용자가 고친 값까지 반영한 분석 결과."""
        lines = lambda w: [l.strip() for l in w.toPlainText().splitlines() if l.strip()]  # noqa: E731
        has_dl = self.chk_deadline.isChecked()
        return {
            'title': self.ed_title.text().strip(),
            'doc_number': self.ed_docno.text().strip(),
            'doc_date': doc_ai.normalize_date(self.ed_docdate.text()),
            'sender': self.ed_sender.text().strip(),
            'deadline': self.ed_date.date().toString('yyyy-MM-dd') if has_dl else '',
            'deadline_time': self.ed_time.time().toString('HH:mm') if has_dl and self.chk_time.isChecked() else '',
            'summary': lines(self.ed_summary),
            'todos': lines(self.ed_todos),
            'attachments': [x.strip() for x in self.ed_attach.text().split(',') if x.strip()],
            'needs_reply': getattr(self, '_needs_reply', False),
        }

    # ── 동작 ────────────────────────────────────────────────────
    def _register(self):
        a = self.collect()
        if not a['title']:
            return
        self._on_register(a['title'], doc_ai.task_deadline(a), doc_ai.task_notes(a), doc_ai.related_line(a))
        self.close()

    def _reply(self):
        a = self.collect()
        if a['title'] and self._on_reply:
            self._on_reply(a)

    def _copy(self, text, msg):
        if text:
            QApplication.clipboard().setText(text)
            self._set_status(msg, 'done')

    def closeEvent(self, event):
        self._closed = True
        # 아직 도는 분석이 닫힌 창을 건드리지 않도록 연결만 끊는다 (스레드는 _ALIVE_WORKERS가 유지)
        if self._worker is not None:
            for sig in (self._worker.progress, self._worker.analyzed, self._worker.failed):
                try:
                    sig.disconnect()
                except TypeError:
                    pass
        super().closeEvent(event)
