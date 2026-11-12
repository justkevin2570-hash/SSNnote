"""AI 제안 확인 칸 — 예산 창 · 공문 작성 창이 쓴다 (AI 대화 창과 같은 모양 · 같은 동작).

  · 노란 칸: 바꾸기 [넣기] [취소] (할 수 없는 제안은 회색 줄)
  · 빨간 칸: 지우기 [지우기] [취소]
  · 보기·열기(now)는 show_reply에서 바로 실행
  · 실행한 뒤 undo() 한 번 → ai_actions.undo
say(글, undo=bool): 창이 결과를 알리는 방법 (상태 줄 · 답 칸 등). undo=True면 「되돌리기」를 걸어 둘 것.
"""
import html

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QVBoxLayout, QWidget


class AiConfirmPanel(QWidget):
    def __init__(self, app_factory, say, parent=None, on_done=None):
        super().__init__(parent)
        self._app_factory = app_factory
        self._say = say
        self._on_done = on_done or (lambda: None)        # 실행·되돌리기 뒤 (예: 예산 창 다시 그리기)
        self._pending, self._pending_del, self._journal = [], [], None
        self.setObjectName('AiConfirm')
        self.setStyleSheet("""
            QWidget#AiConfirm QLabel { background: transparent; }
            QFrame#Proposal { background: #FFF8E1; border: 1px solid #F2D48A; border-radius: 8px; }
            QFrame#ProposalDel { background: #FFF1F2; border: 1px solid #FECDD3; border-radius: 8px; }
            QWidget#AiConfirm QPushButton { background: #FFFFFF; border: 1px solid #D9D4C3; border-radius: 7px;
                                            padding: 4px 12px; color: #1E293B; }
            QWidget#AiConfirm QPushButton#Apply { background: #D4B800; border: 1px solid #D4B800; color: white;
                                                  font-weight: 700; }
            QWidget#AiConfirm QPushButton#Danger { background: #E11D48; border: 1px solid #E11D48; color: white;
                                                   font-weight: 700; }
        """)
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(6)
        self.proposal, self.lbl_proposal, self.btn_apply, self.btn_skip = self._box(
            'Proposal', '넣기', 'Apply', self.apply, self.hide_changes)
        self.proposal_del, self.lbl_delete, self.btn_delete, _ = self._box(
            'ProposalDel', '지우기', 'Danger', self.apply_deletes, self.hide_deletes)
        v.addWidget(self.proposal)
        v.addWidget(self.proposal_del)
        self.hide()

    def _box(self, name, ok_text, ok_name, on_ok, on_cancel):
        f = QFrame()
        f.setObjectName(name)
        f.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        lay = QVBoxLayout(f)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(6)
        lbl = QLabel()
        lbl.setWordWrap(True)
        lbl.setTextFormat(Qt.RichText)
        lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
        lay.addWidget(lbl)
        row = QHBoxLayout()
        row.addStretch()
        ok = QPushButton(ok_text)
        ok.setObjectName(ok_name)
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(on_ok)
        cancel = QPushButton('취소')
        cancel.setCursor(Qt.PointingHandCursor)
        cancel.clicked.connect(on_cancel)
        row.addWidget(ok)
        row.addWidget(cancel)
        lay.addLayout(row)
        f.hide()
        return f, lbl, ok, cancel

    # ── 답 보여 주기 ─────────────────────────────────────────────
    def show_reply(self, reply, now_ok=True):
        """보기·열기는 바로 실행하고, 바꾸기·지우기는 확인 칸에. 바로 실행한 것의 설명 목록을 돌려준다."""
        import ai_actions
        done = []
        now = reply.get('now') or []
        if now and now_ok:
            try:
                ai_actions.execute(now, self._app_factory())
                done = [it['desc'] for it in now]
                self._on_done()
            except Exception as e:
                self._say(f'실행하지 못했어요: {e}', False)
        self.show_changes(reply.get('changes') or [], reply.get('rejected') or [])
        self.show_deletes(reply.get('deletes') or [])
        return done

    def show_changes(self, changes, rejected=()):
        self._pending = list(changes)
        if not changes and not rejected:
            self.proposal.hide()
            self._sync()
            return
        lines = [f'• {html.escape(c["desc"])}' for c in changes]
        lines += [f'<span style="color:#94A3B8">• {html.escape(r["desc"])} — {html.escape(r["reason"])}'
                  f' (할 수 없어요)</span>' for r in rejected]
        head = '이렇게 할까요?' if changes else '할 수 있는 제안이 없어요'
        self.lbl_proposal.setText(f'<b>{head}</b><br>' + '<br>'.join(lines))
        self.btn_apply.setVisible(bool(changes))
        self.btn_skip.setText('취소' if changes else '닫기')
        self.proposal.show()
        self._sync()

    def show_deletes(self, deletes):
        self._pending_del = list(deletes)
        if not deletes:
            self.proposal_del.hide()
            self._sync()
            return
        lines = '<br>'.join(f'• {html.escape(d["desc"])}' for d in deletes)
        self.lbl_delete.setText(f'<b style="color:#BE123C">지울까요?</b><br>{lines}<br>'
                                '<span style="color:#64748B">지운 뒤에도 되돌리기를 한 번 할 수 있어요.</span>')
        self.proposal_del.show()
        self._sync()

    def hide_changes(self):
        self._pending = []
        self.proposal.hide()
        self._sync()

    def hide_deletes(self):
        self._pending_del = []
        self.proposal_del.hide()
        self._sync()

    def clear(self):
        self.hide_changes()
        self.hide_deletes()
        self._journal = None

    def _sync(self):
        self.setVisible(not self.proposal.isHidden() or not self.proposal_del.isHidden())

    # ── 실행 · 되돌리기 ──────────────────────────────────────────
    def _run(self, items, word):
        import ai_actions
        try:
            self._journal = ai_actions.execute(items, self._app_factory())
        except Exception as e:
            self._say(f'실행하지 못했어요: {e}', False)
            return
        self._on_done()
        self._say(f'✓ {len(items)}건을 {word}.', True)

    def apply(self):
        if self._pending:
            items = self._pending
            self.hide_changes()
            self._run(items, '했어요')

    def apply_deletes(self):
        if self._pending_del:
            items = self._pending_del
            self.hide_deletes()
            self._run(items, '지웠어요')

    def can_undo(self):
        return bool(self._journal)

    def undo(self):
        import ai_actions
        if not self._journal:
            return False
        journal, self._journal = self._journal, None
        try:
            ai_actions.undo(journal, self._app_factory())
        except Exception as e:
            self._say(f'되돌리지 못했어요: {e}', False)
            return False
        self._on_done()
        self._say('↩ 하기 전으로 되돌렸어요.', False)
        return True
