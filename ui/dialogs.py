"""ui/dialogs.py — 범용 대화상자 (사용자 지정 반복)."""
from PyQt5.QtWidgets import (
    QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDialog, QComboBox
)
from PyQt5.QtCore import Qt
from .theme import tc, tct

class CustomRecurrenceDialog(QDialog):
    """월/주/요일 사용자 설정 다이얼로그"""
    def __init__(self, parent=None, current=''):
        super().__init__(parent)
        self.setWindowTitle('사용자 설정 반복')
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setMinimumWidth(300)
        self.setStyleSheet(f"font-family: 'Malgun Gothic'; font-size: 10pt; background-color: {tc('#ffffff')}; color: {tct('#111111')};")
        
        layout = QVBoxLayout(self)
        
        # 월 선택
        row_m = QHBoxLayout()
        row_m.addWidget(QLabel('반복 월:'))
        self.combo_m = QComboBox()
        self.combo_m.addItem('매월', 0)
        for i in range(1, 13):
            self.combo_m.addItem(f'{i}월', i)
        row_m.addWidget(self.combo_m, 1)
        layout.addLayout(row_m)
        
        # 주 선택
        row_w = QHBoxLayout()
        row_w.addWidget(QLabel('반복 주:'))
        self.combo_w = QComboBox()
        for i in range(1, 5):
            self.combo_w.addItem(f'{i}째주', i)
        row_w.addWidget(self.combo_w, 1)
        layout.addLayout(row_w)
        
        # 요일 선택
        row_wd = QHBoxLayout()
        row_wd.addWidget(QLabel('반복 요일:'))
        self.combo_wd = QComboBox()
        weekdays = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
        for i, name in enumerate(weekdays):
            self.combo_wd.addItem(name, i)
        row_wd.addWidget(self.combo_wd, 1)
        layout.addLayout(row_wd)
        
        # 현재 값 적용
        if current.startswith('custom:'):
            try:
                _, parts = current.split(':')
                m, w, wd = map(int, parts.split('|'))
                idx_m = self.combo_m.findData(m)
                if idx_m >= 0: self.combo_m.setCurrentIndex(idx_m)
                idx_w = self.combo_w.findData(w)
                if idx_w >= 0: self.combo_w.setCurrentIndex(idx_w)
                idx_wd = self.combo_wd.findData(wd)
                if idx_wd >= 0: self.combo_wd.setCurrentIndex(idx_wd)
            except: pass

        # 버튼
        btns = QHBoxLayout()
        btn_ok = QPushButton('확인')
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton('취소')
        btn_cancel.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)

    def get_value(self):
        m = self.combo_m.currentData()
        w = self.combo_w.currentData()
        wd = self.combo_wd.currentData()
        return f'custom:{m}|{w}|{wd}'
