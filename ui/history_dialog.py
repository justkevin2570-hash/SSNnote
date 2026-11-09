"""ui/history_dialog.py — 「지난 기록」 창: 지난 업무 · 작성 공문 두 칸.

지난 업무: 삭제한 업무를 마감 달별로 묶어 보여주고, 되살리기(마감이 지나지 않은 것만)·완전 삭제·메모 보기.
작성 공문: 공문 작성 창에서 「저장 후 나가기」로 저장한 공문 — 목록 + 내용 보기, 제목·본문 복사, 삭제.
처음엔 「지난 업무」로 열리고, 그다음부터는 마지막에 본 칸으로 열린다.
"""

from collections import defaultdict
from datetime import date

from PyQt5.QtCore import QSettings, QSize, Qt
from PyQt5.QtGui import QFont, QFontMetrics
from PyQt5.QtWidgets import (
    QApplication, QDialog, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QLineEdit, QMessageBox, QPlainTextEdit, QPushButton, QScrollArea, QSplitter, QStackedWidget,
    QVBoxLayout, QWidget,
)

from .calendar_view import P, _ghost_button, _icon_button, _ui_font
from .common import px
from .dates import calc_dday

TAB_KEY = 'history_dlg/tab'
GEO_KEY = 'history_dlg/geometry'


def _short_date(iso: str) -> str:
    try:
        d = date.fromisoformat(iso[:10])
        return f'{d.month}. {d.day}.' if d.year == date.today().year else f'{d.year}. {d.month}. {d.day}.'
    except (ValueError, TypeError):
        return ''


class _Segment(QWidget):
    """위쪽 전환 버튼 [지난 업무 3] [작성 공문 5]."""

    def __init__(self, labels, on_change, parent=None):
        super().__init__(parent)
        self.setObjectName('Segment')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(f"QWidget#Segment {{ background: {P.CHIP}; border-radius: 9px; }}")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(3, 3, 3, 3)
        lay.setSpacing(2)
        self._buttons = []
        for i, text in enumerate(labels):
            b = QPushButton(text)
            b.setCursor(Qt.PointingHandCursor)
            b.setFont(_ui_font(10, QFont.DemiBold))
            b.clicked.connect(lambda _, i=i: on_change(i))
            lay.addWidget(b)
            self._buttons.append(b)

    def set_label(self, i, text):
        self._buttons[i].setText(text)

    def select(self, idx):
        for i, b in enumerate(self._buttons):
            if i == idx:
                b.setStyleSheet(f"QPushButton {{ background: {P.BG}; color: {P.TEXT}; border: none; border-radius: 7px;"
                                f" padding: 6px 16px; }}")
            else:
                b.setStyleSheet(f"QPushButton {{ background: transparent; color: {P.MUTED}; border: none;"
                                f" border-radius: 7px; padding: 6px 16px; }}"
                                f"QPushButton:hover {{ color: {P.TEXT}; }}")


class HistoryDialog(QDialog):
    def __init__(self, parent, window_id, on_restore, tab=None):
        """on_restore(record): 업무를 되살린 뒤 메모 창을 새로 그린다 (메모 창이 넘겨줌)."""
        super().__init__(parent, Qt.Dialog | Qt.WindowCloseButtonHint | Qt.WindowMaximizeButtonHint)
        self.setWindowTitle('지난 기록')
        self.setStyleSheet(
            f"QDialog {{ background: {P.BG}; }} QLabel {{ background: transparent; color: {P.TEXT}; }}"
            f"QScrollArea {{ background: transparent; border: none; }}"
            f"QScrollBar:vertical {{ background: transparent; width: 8px; margin: 2px; }}"
            f"QScrollBar::handle:vertical {{ background: {P.LINE}; border-radius: 3px; min-height: 30px; }}"
            f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}")
        self._window_id = window_id
        self._on_restore = on_restore
        self.setMinimumSize(px(520), px(460))
        self.resize(px(600), px(560))
        geo = QSettings('SSNnote', 'SSNnote').value(GEO_KEY)
        if geo:
            self.restoreGeometry(geo)
        self._build()
        if tab is None:
            tab = int(QSettings('SSNnote', 'SSNnote').value(TAB_KEY, 0) or 0)
        self.show_tab(tab)

    def done(self, r):
        s = QSettings('SSNnote', 'SSNnote')
        s.setValue(GEO_KEY, self.saveGeometry())
        s.setValue(TAB_KEY, self._stack.currentIndex())
        super().done(r)

    # ── 화면 ────────────────────────────────────────────────────
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 16)
        root.setSpacing(14)
        head = QHBoxLayout()
        title = QLabel('지난 기록')
        title.setFont(_ui_font(15, QFont.Bold))
        head.addWidget(title)
        head.addStretch()
        self._seg = _Segment(['지난 업무', '작성 공문'], self.show_tab)
        head.addWidget(self._seg)
        root.addLayout(head)

        self._stack = QStackedWidget()
        self._task_page = QWidget()
        self._doc_page = QWidget()
        self._stack.addWidget(self._task_page)
        self._stack.addWidget(self._doc_page)
        root.addWidget(self._stack, 1)
        self._build_task_page()
        self._build_doc_page()

        foot = QHBoxLayout()
        self._hint = QLabel('')
        self._hint.setFont(_ui_font(8.5))
        self._hint.setStyleSheet(f'color: {P.FAINT};')
        foot.addWidget(self._hint)
        foot.addStretch()
        close = _ghost_button('닫기')
        close.clicked.connect(self.accept)
        foot.addWidget(close)
        root.addLayout(foot)

    def show_tab(self, idx):
        idx = 0 if idx not in (0, 1) else idx
        self._stack.setCurrentIndex(idx)
        self._seg.select(idx)
        self._hint.setText('마감이 지나지 않은 업무만 되살릴 수 있어요.' if idx == 0
                           else '공문 작성 창에서 「저장 후 나가기」로 저장한 공문이에요.')

    # ── 지난 업무 ───────────────────────────────────────────────
    def _build_task_page(self):
        lay = QVBoxLayout(self._task_page)
        lay.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        inner = QWidget()
        inner.setStyleSheet('background: transparent;')
        self._task_list = QVBoxLayout(inner)
        self._task_list.setContentsMargins(0, 0, 6, 0)
        self._task_list.setSpacing(2)
        scroll.setWidget(inner)
        lay.addWidget(scroll)
        self._reload_tasks()

    def _reload_tasks(self):
        from db import get_task_history
        while self._task_list.count():
            it = self._task_list.takeAt(0)
            if it.widget():
                it.widget().hide()
                it.widget().deleteLater()
        records = get_task_history(self._window_id)
        self._seg.set_label(0, f'지난 업무  {len(records)}' if records else '지난 업무')
        if not records:
            self._task_list.addWidget(self._empty('ph.clock-counter-clockwise', '지난 업무가 없어요',
                                                  '업무를 삭제하면 여기에 모여요. 실수로 지웠다면 되살릴 수 있어요.'))
            self._task_list.addStretch()
            return
        cur = date.today().strftime('%Y-%m')
        groups = defaultdict(list)
        for r in records:
            ym = r['deadline'][:7] if r.get('deadline') and r['deadline'][:7] < cur else cur
            groups[ym].append(r)
        for ym in sorted(groups, reverse=True):          # 이번 달(마감 없음 포함) → 지난달 순
            self._task_list.addWidget(self._month_group(ym, groups[ym], expanded=(ym == cur)))
        self._task_list.addStretch()

    def _month_group(self, ym, items, expanded):
        import qtawesome as qta
        box = QWidget()
        bl = QVBoxLayout(box)
        bl.setContentsMargins(0, 4, 0, 6)
        bl.setSpacing(1)
        y, m = ym.split('-')
        head = QPushButton(f'  {y}년 {int(m)}월')
        head.setCursor(Qt.PointingHandCursor)
        head.setFont(_ui_font(9.5, QFont.DemiBold))
        head.setIconSize(QSize(12, 12))
        head.setStyleSheet(f"QPushButton {{ background: transparent; border: none; color: {P.MUTED};"
                           f" text-align: left; padding: 4px 2px; }} QPushButton:hover {{ color: {P.TEXT}; }}")
        count = QLabel(str(len(items)))
        count.setFont(_ui_font(9, QFont.DemiBold))
        count.setStyleSheet(f'color: {P.FAINT};')
        hr = QHBoxLayout()
        hr.setSpacing(4)
        hr.addWidget(head)
        hr.addWidget(count)
        hr.addStretch()
        bl.addLayout(hr)
        body = QWidget()
        bdl = QVBoxLayout(body)
        bdl.setContentsMargins(0, 0, 0, 0)
        bdl.setSpacing(1)
        for r in items:
            bdl.addWidget(self._task_row(r, box, count))
        bl.addWidget(body)

        def set_open(on):
            body.setVisible(on)
            head.setIcon(qta.icon('ph.caret-down-bold' if on else 'ph.caret-right-bold', color=P.MUTED))
        head.clicked.connect(lambda: set_open(body.isHidden()))
        set_open(expanded)
        box._count = count
        return box

    def _task_row(self, r, group, count_lbl):
        import qtawesome as qta
        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(0, 0, 0, 0)
        wl.setSpacing(0)
        row = QWidget()
        row.setObjectName('HRow')
        row.setAttribute(Qt.WA_StyledBackground, True)
        row.setStyleSheet(f"QWidget#HRow {{ background: transparent; border-radius: 8px; }}"
                          f"QWidget#HRow:hover {{ background: {P.HOVER}; }}")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(10, 7, 6, 7)
        rl.setSpacing(8)
        texts = QVBoxLayout()
        texts.setSpacing(1)
        name = QLabel()
        nf = _ui_font(10.5, QFont.Medium)
        name.setFont(nf)
        name.setText(QFontMetrics(nf).elidedText(r['name'], Qt.ElideRight, px(330)))
        name.setToolTip(r['name'])
        name.setTextInteractionFlags(Qt.TextSelectableByMouse)
        texts.addWidget(name)
        sub = []
        overdue = False
        if r.get('deadline'):
            dday = calc_dday(r['deadline'])
            overdue = dday.startswith('D+')
            sub.append(f"마감 {_short_date(r['deadline'])}{' ' + r['deadline'][11:16] if len(r['deadline']) > 10 else ''} · {dday}")
        else:
            sub.append('마감일 없음')
        sub.append(f"삭제 {_short_date(r.get('cleared_at', ''))}")
        sl = QLabel(' · '.join(sub))
        sl.setFont(_ui_font(8.5))
        sl.setStyleSheet(f'color: {P.MUTED};')
        texts.addWidget(sl)
        rl.addLayout(texts, 1)

        note = (r.get('notes') or '').strip()
        related = (r.get('related_no') or '').strip()
        panel = None
        if note or related:
            panel = QLabel(('공문: ' + related + ('\n' if note else '') if related else '') + note)
            panel.setWordWrap(True)
            panel.setFont(_ui_font(9.5))
            panel.setTextInteractionFlags(Qt.TextSelectableByMouse)
            panel.setStyleSheet(f'color: {P.TEXT2}; background: {P.SURFACE}; border: 1px solid {P.LINE};'
                                f' border-radius: 8px; padding: 8px 10px; margin: 0 6px 4px 10px;')
            panel.hide()
            b_note = _icon_button('ph.notepad', '메모 보기', 28)
            b_note.clicked.connect(lambda: panel.setVisible(panel.isHidden()))
            rl.addWidget(b_note)
        if not overdue:
            b_restore = QPushButton('되살리기')
            b_restore.setIcon(qta.icon('ph.arrow-counter-clockwise', color=P.TEXT2))
            b_restore.setCursor(Qt.PointingHandCursor)
            b_restore.setFont(_ui_font(9, QFont.DemiBold))
            b_restore.setStyleSheet(f"QPushButton {{ background: transparent; border: 1px solid {P.LINE};"
                                    f" border-radius: 7px; padding: 4px 10px; color: {P.TEXT}; }}"
                                    f"QPushButton:hover {{ background: {P.CHIP}; }}")
            b_restore.setToolTip('업무 목록으로 되돌리기')
            b_restore.clicked.connect(lambda: self._restore(r, wrap, group, count_lbl))
            rl.addWidget(b_restore)
        b_del = _icon_button('ph.trash', '기록에서 완전히 지우기', 28)
        b_del.clicked.connect(lambda: self._delete(r, wrap, group, count_lbl))
        rl.addWidget(b_del)
        wl.addWidget(row)
        if panel is not None:
            wl.addWidget(panel)
        return wrap

    def _gone(self, wrap, group, count_lbl):
        wrap.hide()
        n = int(count_lbl.text() or '1') - 1
        count_lbl.setText(str(n))
        if n <= 0:
            group.hide()
        from db import get_task_history
        left = len(get_task_history(self._window_id))
        self._seg.set_label(0, f'지난 업무  {left}' if left else '지난 업무')

    def _restore(self, r, wrap, group, count_lbl):
        self._on_restore(r)
        self._gone(wrap, group, count_lbl)

    def _delete(self, r, wrap, group, count_lbl):
        from db import delete_task_history
        delete_task_history(r['id'])
        self._gone(wrap, group, count_lbl)

    # ── 작성 공문 ───────────────────────────────────────────────
    def _build_doc_page(self):
        lay = QVBoxLayout(self._doc_page)
        lay.setContentsMargins(0, 0, 0, 0)
        self._doc_empty = self._empty('ph.files', '저장한 공문이 없어요',
                                      '공문 작성 창에서 「저장 후 나가기」를 누르면 여기에 모여요.')
        lay.addWidget(self._doc_empty)
        self._doc_split = QSplitter(Qt.Horizontal)
        self._doc_split.setHandleWidth(1)
        self._doc_split.setStyleSheet(f"QSplitter::handle {{ background: {P.LINE}; }}")
        self._doc_list = QListWidget()
        self._doc_list.setFont(_ui_font(10))
        self._doc_list.setStyleSheet(
            f"QListWidget {{ background: transparent; border: none; outline: 0; }}"
            f"QListWidget::item {{ color: {P.TEXT}; padding: 8px 10px; border-radius: 8px; margin: 1px 6px 1px 0; }}"
            f"QListWidget::item:hover {{ background: {P.HOVER}; }}"
            f"QListWidget::item:selected {{ background: {P.CHIP}; color: {P.TEXT}; }}")
        self._doc_list.setMinimumWidth(px(170))
        self._doc_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._doc_list.setWordWrap(True)                 # 긴 제목은 줄바꿈 (가로 스크롤 없이)
        self._doc_list.setTextElideMode(Qt.ElideNone)
        self._doc_split.addWidget(self._doc_list)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(16, 0, 0, 0)
        rl.setSpacing(6)
        self._doc_title = QLabel()
        self._doc_title.setWordWrap(True)
        self._doc_title.setFont(_ui_font(12, QFont.Bold))
        self._doc_title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self._doc_meta = QLabel()
        self._doc_meta.setFont(_ui_font(9))
        self._doc_meta.setStyleSheet(f'color: {P.MUTED};')
        self._doc_meta.setTextInteractionFlags(Qt.TextSelectableByMouse)
        # 수정할 때만 보이는 입력 칸 (평소에는 숨김)
        self._doc_title_edit = QLineEdit()
        self._doc_title_edit.setFont(_ui_font(12, QFont.Bold))
        self._doc_title_edit.setPlaceholderText('공문 제목')
        self._doc_title_edit.hide()
        self._doc_num_edit = QLineEdit()
        self._doc_num_edit.setFont(_ui_font(9))
        self._doc_num_edit.setPlaceholderText('관련 공문번호 (선택)')
        self._doc_num_edit.hide()
        self._doc_body = QPlainTextEdit()
        self._doc_body.setReadOnly(True)
        bf = QFont('굴림체', 11)
        self._doc_body.setFont(bf)
        self._doc_body_style_view = (f"QPlainTextEdit {{ background: {P.SURFACE}; border: 1px solid {P.LINE};"
                                     f" border-radius: 10px; padding: 10px 12px; color: {P.TEXT}; }}")
        self._doc_body_style_edit = (f"QPlainTextEdit {{ background: {P.SURFACE}; border: 1px solid {P.TEXT2};"
                                     f" border-radius: 10px; padding: 10px 12px; color: {P.TEXT}; }}")
        self._doc_body.setStyleSheet(self._doc_body_style_view)
        acts = QHBoxLayout()
        acts.setSpacing(6)
        self._b_copy_title = _ghost_button('제목 복사')
        self._b_copy_body = _ghost_button('본문 복사')
        self._b_doc_edit = _ghost_button('수정')
        self._b_doc_cancel = _ghost_button('취소')
        self._b_doc_cancel.hide()
        self._b_doc_del = _icon_button('ph.trash', '이 공문 기록 지우기', 32)
        self._b_copy_title.clicked.connect(lambda: self._copy(self._cur_doc.get('title', ''), '제목'))
        self._b_copy_body.clicked.connect(lambda: self._copy(self._cur_doc.get('content', ''), '본문'))
        self._b_doc_edit.clicked.connect(self._on_edit_clicked)
        self._b_doc_cancel.clicked.connect(self._cancel_edit)
        self._b_doc_del.clicked.connect(self._delete_doc)
        acts.addWidget(self._b_copy_title)
        acts.addWidget(self._b_copy_body)
        acts.addWidget(self._b_doc_edit)
        acts.addWidget(self._b_doc_cancel)
        acts.addStretch()
        acts.addWidget(self._b_doc_del)
        self._editing = False
        rl.addWidget(self._doc_title)
        rl.addWidget(self._doc_title_edit)
        rl.addWidget(self._doc_meta)
        rl.addWidget(self._doc_num_edit)
        rl.addWidget(self._doc_body, 1)
        rl.addLayout(acts)
        self._doc_split.addWidget(right)
        self._doc_split.setStretchFactor(1, 1)
        self._doc_split.setSizes([px(200), px(380)])
        lay.addWidget(self._doc_split, 1)
        self._cur_doc = {}
        self._doc_list.currentItemChanged.connect(lambda *_: self._show_doc())
        self._reload_docs()

    def _reload_docs(self, keep_id=None):
        from db import get_official_documents
        docs = get_official_documents()
        self._seg.set_label(1, f'작성 공문  {len(docs)}' if docs else '작성 공문')
        self._doc_list.clear()
        self._doc_empty.setVisible(not docs)
        self._doc_split.setVisible(bool(docs))
        for doc in docs:
            it = QListWidgetItem(f"{doc.get('title') or '(제목 없음)'}\n{_short_date(doc.get('created_at', ''))}")
            it.setData(Qt.UserRole, doc)
            it.setToolTip(doc.get('title') or '')
            self._doc_list.addItem(it)
        if docs:
            row = next((i for i, d in enumerate(docs) if keep_id is not None and d.get('id') == keep_id), 0)
            self._doc_list.setCurrentRow(row)

    def _show_doc(self):
        it = self._doc_list.currentItem()
        if not it:
            return
        doc = it.data(Qt.UserRole)
        self._cur_doc = doc
        self._doc_title.setText(doc.get('title') or '(제목 없음)')
        meta = [f"저장 {_short_date(doc.get('created_at', ''))}"]
        if doc.get('doc_number'):
            meta.append(f"관련 {doc['doc_number']}")
        self._doc_meta.setText(' · '.join(meta))
        self._doc_body.setPlainText(doc.get('content', ''))

    # ── 수정 ──
    def _set_editing(self, on):
        self._editing = on
        self._doc_title.setVisible(not on)
        self._doc_meta.setVisible(not on)
        self._doc_title_edit.setVisible(on)
        self._doc_num_edit.setVisible(on)
        self._doc_body.setReadOnly(not on)
        self._doc_body.setStyleSheet(self._doc_body_style_edit if on else self._doc_body_style_view)
        self._b_doc_edit.setText('저장' if on else '수정')
        self._b_doc_cancel.setVisible(on)
        self._b_copy_title.setEnabled(not on)
        self._b_copy_body.setEnabled(not on)
        self._b_doc_del.setEnabled(not on)
        self._doc_list.setEnabled(not on)          # 고치는 중에는 다른 공문으로 넘어가지 않게

    def _on_edit_clicked(self):
        if self._editing:
            return self._save_edit()
        if not self._cur_doc:
            return
        self._doc_title_edit.setText(self._cur_doc.get('title') or '')
        self._doc_num_edit.setText(self._cur_doc.get('doc_number') or '')
        self._doc_body.setPlainText(self._cur_doc.get('content') or '')
        self._set_editing(True)
        self._doc_body.setFocus()
        self._hint.setText('고친 뒤 [저장]을 누르세요. 저장하면 다음 자동 초안의 예시에도 반영돼요.')

    def _save_edit(self):
        from db import update_official_document
        title = self._doc_title_edit.text().strip()
        number = self._doc_num_edit.text().strip()
        content = self._doc_body.toPlainText().rstrip()
        if not title:
            QMessageBox.warning(self, '공문 수정', '제목을 비울 수 없어요.')
            return
        if not content.strip():
            QMessageBox.warning(self, '공문 수정', '본문이 비어 있어요. 내용을 넣거나 [취소]를 누르세요.')
            return
        doc_id = self._cur_doc['id']
        update_official_document(doc_id, title, number, content)
        self._set_editing(False)
        self._reload_docs(keep_id=doc_id)
        self._hint.setText('공문을 저장했어요.')

    def _cancel_edit(self):
        d = self._cur_doc
        changed = (self._doc_title_edit.text().strip() != (d.get('title') or '')
                   or self._doc_num_edit.text().strip() != (d.get('doc_number') or '')
                   or self._doc_body.toPlainText().rstrip() != (d.get('content') or '').rstrip())
        if changed and QMessageBox.question(
                self, '수정 취소', '고친 내용을 버리고 되돌릴까요?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        self._set_editing(False)
        self._show_doc()
        self._hint.setText('')

    def _copy(self, text, what):
        if text:
            QApplication.clipboard().setText(text)
            self._hint.setText(f'{what}을 복사했어요.')

    def _delete_doc(self):
        from db import delete_official_document
        doc = self._cur_doc
        if not doc:
            return
        ok = QMessageBox.question(self, '공문 기록 지우기',
                                  f"「{doc.get('title') or '(제목 없음)'}」 기록을 지울까요?",
                                  QMessageBox.Yes | QMessageBox.No)
        if ok == QMessageBox.Yes:
            delete_official_document(doc['id'])
            self._cur_doc = {}
            self._reload_docs()

    # ── 공통 ────────────────────────────────────────────────────
    def _empty(self, icon, title, sub):
        import qtawesome as qta
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, px(50), 0, px(50))
        lay.setSpacing(6)
        ic = QLabel()
        ic.setPixmap(qta.icon(icon, color=P.FAINT).pixmap(34, 34))
        ic.setAlignment(Qt.AlignCenter)
        t = QLabel(title)
        t.setFont(_ui_font(11, QFont.DemiBold))
        t.setAlignment(Qt.AlignCenter)
        s = QLabel(sub)
        s.setFont(_ui_font(9))
        s.setStyleSheet(f'color: {P.MUTED};')
        s.setAlignment(Qt.AlignCenter)
        s.setWordWrap(True)
        for x in (ic, t, s):
            lay.addWidget(x)
        return w
