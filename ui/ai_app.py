"""ai_actions가 앱(열린 창 · 예산 창 · 공문 작성 창 · 설정)을 다룰 때 쓰는 어댑터.

AI 대화 창 · 예산 창 · 공문 작성 창이 모두 이것을 써서 같은 일을 할 수 있다.
host: 기준 업무창 (새 업무를 넣을 창, 새 창을 띄울 위치). 없으면 find_host()로 찾는다.
"""


def find_host():
    """기준 업무창: AI 대화가 붙어 있는 업무창 → 보이는 업무창(메모장 아님) → 아무 업무창."""
    try:
        from . import ai_chat_window
        w = ai_chat_window.instance(create=False)
        if w is not None and w.host is not None:
            w.host.window_id                                   # 지워진 창이면 RuntimeError
            return w.host
    except (RuntimeError, AttributeError):
        pass
    from PyQt5.QtWidgets import QApplication
    cands = []
    for top in QApplication.topLevelWidgets():
        if type(top).__name__ == 'MemoWindow':
            try:
                cands.append((not top.isVisible(), bool(getattr(top, '_memo_mode', False)), top.window_id, top))
            except RuntimeError:
                pass
    cands.sort(key=lambda c: c[:3])
    return cands[0][3] if cands else None


class AppAdapter:
    def __init__(self, host=None, chat=None, on_clear=None):
        self.host = host if host is not None else find_host()
        self._chat = chat
        self._on_clear = on_clear

    # ── 창 ──────────────────────────────────────────────────────
    @property
    def host_window_id(self):
        try:
            return getattr(self.host, 'window_id', None)
        except RuntimeError:
            return None

    @property
    def chat(self):
        """AI 대화 창 (없으면 만든다 — chat.* 기능용)."""
        if self._chat is None:
            from . import ai_chat_window
            self._chat = ai_chat_window.instance()
        return self._chat

    def windows(self):
        try:
            return [w for w in list(getattr(self.host, '_open_windows', None) or [self.host]) if w is not None]
        except RuntimeError:
            return []

    def window(self, wid):
        for w in self.windows():
            try:
                if w.window_id == wid:
                    return w
            except RuntimeError:
                pass
        return None

    def refresh_tasks(self):
        for w in self.windows():
            try:
                if not getattr(w, '_memo_mode', False):
                    w._refresh_tasks()
            except RuntimeError:
                pass

    def refresh_budget(self):
        try:
            import budget
            if getattr(budget, '_instance', None) is not None:
                budget._instance.refresh()
        except RuntimeError:
            pass

    def new_window(self, kind):
        h = self.host
        if h is None:
            return None
        before = set(id(w) for w in self.windows())
        h.on_new(offset_from=h, memo_mode=(kind == 'memo'), on_toggle_hotkey=h._on_toggle_hotkey,
                 on_shortcut_change=h._on_shortcut_change, get_shortcut_enabled=h._get_shortcut_enabled)
        return next((w for w in self.windows() if id(w) not in before), None)

    def delete_window(self, wid):
        import db
        w = self.window(wid)
        db.delete_window(wid)
        if w is not None:
            w._deleted = True
            w.close()

    def set_memo_text(self, w, text):
        w.memo_editor.setPlainText(text)
        w._apply_memo_line_spacing()
        w._flush_memo_save()

    def style_window(self, w, color, font, size, gap):
        if color:
            w._apply_color(color)
            w.save_state()
        if font:
            w._set_font(font)
        if size is not None:
            w._set_scale(size)
        if gap is not None:
            w._set_row_gap(gap)

    # ── 창 열기 ─────────────────────────────────────────────────
    def open(self, what, title='', category_id=None):
        h = self.host
        if what == 'calendar':
            h._show_monthly_calendar()
        elif what == 'dashboard':
            h._open_dashboard()
        elif what == 'history':
            h._show_history(0)
        elif what == 'documents':
            h._show_history(1)
        elif what == 'document':
            ed = self._editor(open_it=True)
            if ed is not None and title:
                ed.title_input.setText(title)
        elif what == 'budget':
            if h is not None:
                h._open_budget_window()
            else:
                import budget
                budget.open_budget_window()
            import budget
            if category_id and getattr(budget, '_instance', None) is not None:
                budget._instance._set_tab(category_id)

    # ── 공문 작성 창 ─────────────────────────────────────────────
    def _editor(self, open_it=False):
        from document_editor import DocumentEditorWindow
        ed = DocumentEditorWindow._instance
        if ed is None and open_it:
            if self.host is not None:
                self.host._open_document_editor()
            else:
                ed = DocumentEditorWindow()
                ed.show()
            ed = DocumentEditorWindow._instance or ed
        return ed

    def document_state(self):
        """열려 있는 공문 작성 창의 제목·관련 번호·본문 (없으면 None)."""
        ed = self._editor()
        if ed is None:
            return None
        try:
            return {'title': ed.title_input.text(), 'ref': ed.ref_input.text(), 'body': ed.editor.toPlainText()}
        except RuntimeError:
            return None

    def write_document(self, title=None, ref=None, body=None):
        """공문 작성 창에 쓰기 (없으면 연다). 본문은 한 번의 편집이라 Ctrl+Z로도 되돌린다. 전 값을 돌려준다."""
        ed = self._editor(open_it=True)
        if ed is None:
            return None
        old = {'title': ed.title_input.text(), 'ref': ed.ref_input.text(), 'body': ed.editor.toPlainText()}
        if title is not None:
            ed.title_input.setText(title)
        if ref is not None:
            ed.ref_input.setText(ref)
        if body is not None:
            from document_editor import _correct_document_format
            ed._replace_body(_correct_document_format(body))
            ed._last_draft = ed.editor.toPlainText()
        return old

    # ── 설정 ────────────────────────────────────────────────────
    def get_alarm(self):
        try:
            return self.host._get_alarm_interval() if self.host._get_alarm_interval else 180
        except (AttributeError, RuntimeError):
            return 180

    def set_alarm(self, minutes):
        self.host._set_alarm_interval(minutes)

    def get_auto_update(self):
        import updater
        return updater.auto_enabled()

    def get_autostart(self):
        try:
            from autostart import is_enabled
            return bool(is_enabled())
        except Exception:
            return False

    def set_autostart(self, on):
        from autostart import set_enabled
        set_enabled(bool(on))

    def clear_chat_later(self):
        """chat.clear: 지금 대화하는 창의 대화를 지운다."""
        from PyQt5.QtCore import QTimer
        fn = self._on_clear or (lambda: self.chat._clear_chat())
        QTimer.singleShot(0, fn)
