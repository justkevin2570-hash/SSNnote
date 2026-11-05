"""ui/document_row.py — 공문 목록의 한 줄(DocumentRow)."""
from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout,
    QPushButton
)
from PyQt5.QtCore import Qt
from db import (update_document)
from .common import _base_path
from .widgets import SmartDocLineEdit
from .theme import T

class DocumentRow(QWidget):
    """공문번호 한 항목 행: [공문제목 입력] [공문번호(읽기전용)] [삭제메뉴]"""

    def __init__(self, doc: dict, on_delete, on_paste=None):
        super().__init__()
        self._doc = doc
        self._on_delete = on_delete
        self.setStyleSheet('background: transparent;')

        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 2, 8, 2)
        lay.setSpacing(0)

        field_style = f"""
            QLineEdit {{
                background: rgba(255,255,255,0.5);
                border: 1px solid {T.ACCENT};
                border-radius: 4px;
                padding: 2px 6px;
                color: #333;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
        """
        readonly_style = f"""
            QLineEdit {{
                background: rgba(255,255,255,0.25);
                border: 1px solid #c0a800;
                border-radius: 4px;
                padding: 2px 6px;
                color: {T.TEXT_ON_NOTE};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
        """

        editing_style = """
            QLineEdit {
                background: rgba(255,255,255,0.9);
                border: 1px solid #a08800;
                border-radius: 4px;
                padding: 2px 6px;
                color: #111;
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
        """

        self.title_edit = SmartDocLineEdit(doc.get('title', ''))
        self.title_edit.setPlaceholderText('공문 제목 붙여넣는 곳')
        self.title_edit.setReadOnly(True)
        self.title_edit.setStyleSheet(field_style)
        self.title_edit.setCursor(Qt.PointingHandCursor)
        lay.addWidget(self.title_edit, 11)
        lay.addSpacing(0)

        import os
        from PyQt5.QtGui import QIcon, QPixmap
        from PyQt5.QtCore import QSize
        _edit_icon_path = os.path.join(_base_path(), 'assets', '수정 아이콘.png')
        btn_edit = QPushButton()
        btn_edit.setIcon(QIcon(QPixmap(_edit_icon_path).scaled(15, 16, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)))
        btn_edit.setIconSize(QSize(15, 16))
        btn_edit.setFixedSize(17, 18)
        btn_edit.setToolTip('')
        btn_edit.setCursor(Qt.PointingHandCursor)
        btn_edit.setStyleSheet("""
            QPushButton { background: transparent; border: none; padding: 0; }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; }
        """)
        lay.addWidget(btn_edit)
        lay.addSpacing(1)

        self.num_edit = SmartDocLineEdit(doc.get('doc_number', ''))
        self.num_edit.setReadOnly(True)
        self.num_edit.setPlaceholderText('공문 번호 붙여넣는 곳')
        self.num_edit.setStyleSheet(readonly_style)
        self.num_edit.setCursor(Qt.PointingHandCursor)
        lay.addWidget(self.num_edit, 11)

        btn_edit_num = QPushButton()
        btn_edit_num.setIcon(QIcon(QPixmap(_edit_icon_path).scaled(15, 16, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)))
        btn_edit_num.setIconSize(QSize(15, 16))
        btn_edit_num.setFixedSize(17, 18)
        btn_edit_num.setToolTip('')
        btn_edit_num.setCursor(Qt.PointingHandCursor)
        btn_edit_num.setStyleSheet("""
            QPushButton { background: transparent; border: none; padding: 0; }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; }
        """)
        lay.addWidget(btn_edit_num)
        lay.addSpacing(4)

        def _enter_edit():
            self.title_edit._editing = True
            self.title_edit.setReadOnly(False)
            self.title_edit.setStyleSheet(editing_style)
            self.title_edit.setCursor(Qt.IBeamCursor)
            self.title_edit.setFocus()
            self.title_edit.setCursorPosition(0)

        def _commit_edit():
            self.title_edit._editing = False
            self.title_edit.setReadOnly(True)
            self.title_edit.setStyleSheet(field_style)
            self.title_edit.setCursor(Qt.PointingHandCursor)
            self.title_edit.setCursorPosition(0)
            if doc.get('id') is not None:
                self._save()

        _original_title = doc.get('title', '')

        def _cancel_edit():
            self.title_edit.setText(_original_title)
            _commit_edit()

        self.title_edit._cancel_edit = _cancel_edit
        self.title_edit.returnPressed.connect(_commit_edit)
        btn_edit.clicked.connect(lambda: _commit_edit() if self.title_edit._editing else _enter_edit())

        def _enter_edit_num():
            self.num_edit._editing = True
            self.num_edit.setReadOnly(False)
            self.num_edit.setStyleSheet(editing_style)
            self.num_edit.setCursor(Qt.IBeamCursor)
            self.num_edit.setFocus()
            self.num_edit.setCursorPosition(0)

        def _commit_edit_num():
            self.num_edit._editing = False
            self.num_edit.setReadOnly(True)
            self.num_edit.setStyleSheet(readonly_style)
            self.num_edit.setCursor(Qt.PointingHandCursor)
            self.num_edit.setCursorPosition(0)
            if doc.get('id') is not None:
                self._save()

        _original_num = doc.get('doc_number', '')

        def _cancel_edit_num():
            self.num_edit.setText(_original_num)
            _commit_edit_num()

        self.num_edit._cancel_edit = _cancel_edit_num
        self.num_edit.returnPressed.connect(_commit_edit_num)
        btn_edit_num.clicked.connect(lambda: _commit_edit_num() if self.num_edit._editing else _enter_edit_num())

        if on_paste:
            self.title_edit._on_empty_click = lambda: on_paste('title', self.title_edit, self.num_edit)
            self.num_edit._on_empty_click = lambda: on_paste('number', self.title_edit, self.num_edit)

        if doc.get('id') is None:
            lay.addSpacing(22)  # 삭제 버튼(22px) 자리 확보 → 위 행과 폭 통일
            return

        import os
        from PyQt5.QtGui import QIcon, QPixmap
        from PyQt5.QtCore import QSize
        _del_icon_path = os.path.join(_base_path(), 'assets', '엑스아이콘.png')
        btn_del = QPushButton()
        btn_del.setIcon(QIcon(QPixmap(_del_icon_path).scaled(18, 18, Qt.KeepAspectRatio, Qt.SmoothTransformation)))
        btn_del.setIconSize(QSize(18, 18))
        btn_del.setFixedSize(22, 26)
        btn_del.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                padding: 0;
            }
            QPushButton:hover { background: rgba(0,0,0,0.08); border-radius: 3px; }
            QPushButton:pressed { background: rgba(0,0,0,0.18); border-radius: 3px; padding: 2px 0 0 2px; }
        """)
        btn_del.clicked.connect(lambda: self._on_delete(self._doc['id']))
        lay.addWidget(btn_del)

    def _save(self):
        update_document(self._doc['id'], self.title_edit.text(), self.num_edit.text())
