import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QWidget, QLabel, QLineEdit,
    QComboBox, QMessageBox, QInputDialog
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from ai_client import AiStreamThread, load_api_key, save_api_key, load_ai_mode, save_ai_mode
from db import save_official_document, get_official_documents


class DocumentEditorWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ai_thread = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("AI 공문 도우미")
        self._restore_window_state()  # 저장된 위치/크기 복원, 없으면 기본값

        _font = QFont('Malgun Gothic', 12)
        _input_h = 36  # 기본 ~28px의 130%

        central = QWidget()
        central.setFont(_font)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # ── 헤더: 제목 + 유형 ──
        header = QHBoxLayout()
        self.title_input = QLineEdit()
        self.title_input.setMinimumHeight(_input_h)
        self.type_combo = QComboBox()
        self.type_combo.addItems(["계획", "품의", "보관"])
        self.type_combo.setFixedWidth(104)
        self.type_combo.setMinimumHeight(_input_h)
        self.type_combo.setEditable(True)
        self.type_combo.lineEdit().setReadOnly(True)
        self.type_combo.lineEdit().setAlignment(Qt.AlignCenter)
        header.addWidget(QLabel("공문 제목"))
        header.addWidget(self.title_input)
        header.addWidget(self.type_combo)

        # ── 관련 공문번호 ──
        ref_row = QHBoxLayout()
        self.ref_input = QLineEdit()
        self.ref_input.setMinimumHeight(_input_h)
        btn_move_to_title = QPushButton("공문 제목으로 이동")
        btn_move_to_title.setMinimumHeight(_input_h)
        btn_move_to_title.setStyleSheet("font-family: 'Malgun Gothic'; font-size: 11pt; padding-left: 10px; padding-right: 10px;")
        btn_move_to_title.clicked.connect(self._move_ref_to_title)
        ref_row.addWidget(QLabel("관련 번호"))
        ref_row.addWidget(self.ref_input)
        ref_row.addWidget(btn_move_to_title)

        # ── 본문 편집 영역 ──
        self.editor = QTextEdit()
        self.editor.setFont(_font)
        self.editor.setPlaceholderText(
            "AI 초안 생성 버튼을 누르거나, 아래 채팅창에 요청을 입력하세요."
        )

        # ── AI 채팅바 ──
        chat_row = QHBoxLayout()
        self.ai_input = QLineEdit()
        self.ai_input.setPlaceholderText("AI에게 요청 (예: '붙임 문구 추가해줘', '더 간결하게 다듬어줘')")
        self.ai_input.setMinimumHeight(_input_h)
        self.ai_input.returnPressed.connect(self.ask_ai)
        self.btn_send = QPushButton("전송")
        self.btn_send.setFixedWidth(78)
        self.btn_send.setMinimumHeight(_input_h)
        self.btn_send.clicked.connect(self.ask_ai)
        chat_row.addWidget(self.ai_input)
        chat_row.addWidget(self.btn_send)

        # ── 컨트롤바 ──
        ctrl_row = QHBoxLayout()
        _ctrl_btn_style = "font-family: 'Malgun Gothic'; font-size: 12pt; padding-left: 18px; padding-right: 18px;"
        self.btn_draft = QPushButton("AI 초안 생성")
        self.btn_draft.setMinimumHeight(_input_h)
        self.btn_draft.setStyleSheet(_ctrl_btn_style)
        self.btn_draft.clicked.connect(self.generate_draft)
        self.btn_save = QPushButton("저장")
        self.btn_save.setMinimumHeight(_input_h)
        self.btn_save.setStyleSheet(_ctrl_btn_style)
        self.btn_save.clicked.connect(self.save_document)
        self.btn_save_exit = QPushButton("저장 후 나가기")
        self.btn_save_exit.setMinimumHeight(_input_h)
        self.btn_save_exit.setStyleSheet(_ctrl_btn_style)
        self.btn_save_exit.clicked.connect(self.save_and_close)
        self.btn_apikey = QPushButton("AI 설정")
        self.btn_apikey.setMinimumHeight(_input_h)
        self.btn_apikey.setStyleSheet(_ctrl_btn_style)
        self.btn_apikey.clicked.connect(self._show_ai_settings_menu)
        ctrl_row.addWidget(self.btn_draft)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self.btn_apikey)
        ctrl_row.addWidget(self.btn_save)
        ctrl_row.addWidget(self.btn_save_exit)

        # ── 전체 배치 ──
        layout.addLayout(header)
        layout.addLayout(ref_row)
        layout.addWidget(self.editor, stretch=1)
        layout.addLayout(chat_row)
        layout.addLayout(ctrl_row)

        self.statusBar().showMessage("준비 완료")
        self._apply_ai_mode()

    # ── AI 초안 생성 ──
    def generate_draft(self):
        title = self.title_input.text().strip()
        if not title:
            self.statusBar().showMessage("제목을 먼저 입력하세요.")
            return

        doc_type = self.type_combo.currentText()
        ref = self.ref_input.text().strip()
        ref_line = f"관련: {ref}\n" if ref else ""

        prompt = (
            f"제목: {title}\n"
            f"유형: {doc_type}\n"
            f"{ref_line}"
            "위 내용으로 공문 본문(내용/붙임)을 작성해주세요. "
            "개조식, '~함', '~바람', '~고자 함' 문체로 작성하세요."
        )
        self.editor.clear()
        self._start_ai(prompt)

    # ── 자유 AI 채팅 ──
    def ask_ai(self):
        user_text = self.ai_input.text().strip()
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
    def _start_ai(self, prompt):
        if self.ai_thread and self.ai_thread.isRunning():
            return

        self.statusBar().showMessage("AI가 생각 중...")
        self._set_controls_enabled(False)

        self.ai_thread = AiStreamThread(prompt)
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
        self.btn_draft.setEnabled(enabled and ai_available)
        self.ai_input.setEnabled(enabled and ai_available)

    def _apply_ai_mode(self):
        """현재 AI 모드에 따라 AI 관련 버튼 활성/비활성."""
        mode = load_ai_mode()
        ai_available = mode != 'none'
        self.btn_draft.setEnabled(ai_available)
        self.btn_send.setEnabled(ai_available)
        self.ai_input.setEnabled(ai_available)
        if not ai_available:
            self.ai_input.setPlaceholderText('AI 이용 안함 — AI 설정에서 변경하세요.')
        else:
            self.ai_input.setPlaceholderText("AI에게 요청 (예: '붙임 문구 추가해줘', '더 간결하게 다듬어줘')")

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
            QMessageBox.information(self, '내부 AI', '내부 AI 설정은 추후 지원 예정입니다.')
        elif action == act_none:
            save_ai_mode('none')
            self._apply_ai_mode()
            self.statusBar().showMessage('AI 이용 안함 — OCR(winrt)만 사용합니다.')

    def _setup_external_ai(self):
        current = load_api_key()
        key, ok = QInputDialog.getText(
            self, '외부 AI 설정', 'Gemini API 키를 입력하세요:',
            text=current
        )
        if ok and key.strip():
            save_api_key(key.strip())
            save_ai_mode('external')
            self._apply_ai_mode()
            self.statusBar().showMessage('외부 AI(Gemini) 설정 완료.')

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
        doc_type = self.type_combo.currentText()

        existing = get_official_documents()
        for doc in existing:
            if (doc.get('title', '') == title and
                    doc.get('doc_number', '') == doc_number and
                    doc.get('content', '') == content):
                self.statusBar().showMessage('같은 내용이 저장되어 있습니다.')
                QTimer.singleShot(1000, lambda: self.statusBar().showMessage(''))
                return

        save_official_document(title, doc_number, content, doc_type)
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
        mode   = load_ai_mode()

        if mode == 'external' and load_api_key():
            from ai_client import AiImageWorker
            self._img_worker = AiImageWorker(pixmap)
            self._img_worker.finished.connect(lambda r: self._on_ai_result(r, target))
            self._img_worker.error.connect(lambda msg: self._ocr_fallback(pixmap, target))
            self._img_worker.start()
            self.statusBar().showMessage('AI가 공문을 분석 중...')
        else:
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentEditorWindow()
    window.show()
    sys.exit(app.exec_())
