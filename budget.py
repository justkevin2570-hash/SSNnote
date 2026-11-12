# -*- coding: utf-8 -*-
"""budget.py — 예산 관리 창.

화면: 위 요약 한 줄(예산현액·집행액·잔액·집행률) / 왼쪽 세부항목 목록 / 오른쪽 세부항목별 표.
  · 집행액 칸은 눌러서 바로 고치고, 완료·수정·삭제는 행에 마우스를 올리면 나온다.
  · 완료한 항목은 흐리게 아래로 접힌다. 집행률 색: 포인트 색 → 80% 이상 주황 → 100% 이상 빨강.
가져오기: 엑셀 파일(끌어다 놓기) · 엑셀에서 복사한 표 · 예산 화면 캡처(우클릭 → 붙여넣기, Ctrl+V)
  → 세부항목별로 보여 주고(_ImportPreviewDialog) 고친 뒤 가져온다. 캡처 읽기는 budget_image.py.
데이터는 ssnnote.db의 budget_categories, budget_items 테이블에 저장.
"""

import os
import sys
import json

import qtawesome as qta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QScrollArea, QFrame, QTableWidget, QTableWidgetItem,
    QHeaderView, QDialog, QMessageBox, QComboBox, QAbstractItemView,
    QProgressBar, QFileDialog, QButtonGroup, QTextEdit, QPlainTextEdit,
    QMenu, QRadioButton, QProgressDialog, QLayout, QCompleter, QTextBrowser
)
from PyQt5.QtCore import Qt, QSize, QBuffer, QIODevice, QPoint, QRect, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QFontDatabase, QKeySequence, QImage, QPixmap

from db import (get_budget_categories, add_budget_category, rename_budget_category,
                delete_budget_category, get_budget_items, add_budget_item,
                update_budget_item, delete_budget_item, set_budget_item_spent,
                toggle_budget_item_completed, clear_budget_data,
                get_budget_expenses, add_budget_expense, save_budget_item_edit, delete_budget_expense)
from ui.theme import keep_light  # 예산·공문 창은 라이트 전용
from ui.common import px


# ── 예산 창 글꼴 ──────────────────────────────────────────────────
# 예산 창은 작은 글자가 많아서 Windows에서 글자를 픽셀에 맞춰 또렷하게 그려 주는(힌팅된) 맑은 고딕을 쓴다.
# 번들 Pretendard는 힌팅 정보가 없어 12px 안팎에서 흐려 보였다. 맑은 고딕이 없는 PC에서만 Pretendard로 보인다.
# (되돌리려면 이 목록의 순서를 'Pretendard', 'Malgun Gothic'으로 바꾸면 됨)
_UI_FONTS = ('Malgun Gothic', 'Pretendard', 'Apple SD Gothic Neo')
_UI_FONT_QSS = ', '.join(f"'{f}'" for f in _UI_FONTS)


# ── 번들 Pretendard 등록 (맑은 고딕이 없는 PC의 대체 글꼴) ───────────────
# Pretendard가 설치되지 않은 PC에서도 예산 창이 Pretendard로 보이도록
# assets의 정적 폰트(여러 굵기)를 앱 폰트로 등록한다. QSS의
# font-family: 'Pretendard' 가 이 등록된 패밀리를 사용한다.
_pretendard_loaded = False


def _base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _load_pretendard():
    """assets 안의 Pretendard 정적 폰트(.ttf/.otf)를 모두 등록한다. 1회만."""
    global _pretendard_loaded
    if _pretendard_loaded:
        return
    _pretendard_loaded = True
    for dirpath, _dirs, files in os.walk(os.path.join(_base_path(), 'assets')):
        for fn in sorted(files):
            low = fn.lower()
            if low.startswith('pretendard') and low.endswith(('.ttf', '.otf')):
                QFontDatabase.addApplicationFont(os.path.join(dirpath, fn))


# ── 색 (공문 작성 창과 같은 결: 회색 바탕 + 포인트 색 하나) ─────────────
class C:
    BG          = '#F6F7F9'
    SURFACE     = '#FFFFFF'
    BORDER      = '#E3E6EB'
    BORDER_HI   = '#CBD2DC'
    LINE        = '#EEF0F3'    # 표 줄 구분
    TEXT        = '#1E293B'
    BODY        = '#334155'
    MUTED       = '#4F5F76'    # 보조 글 (흰 바탕 대비 6.5:1)
    FAINT       = '#64748B'    # 흐린 글·완료 항목 (4.8:1 — 예전 #94A3B8은 2.6:1이라 흐려 보였음)
    HOVER       = '#F1F3F6'
    ACCENT      = '#4F46E5'
    ACCENT_HI   = '#4338CA'
    ACCENT_SOFT = '#EEF0FD'
    ACCENT_LINE = '#C7CCF7'
    NEG         = '#BE123C'
    WARN_BG     = '#FFF7E0'


# ── 전체 QSS ─────────────────────────────────────────────────────
STYLE = f"""
    QMainWindow, QDialog {{ background: {C.BG}; }}
    QWidget {{ font-family: {_UI_FONT_QSS}; font-size: 10pt; color: {C.TEXT}; }}
    QLabel {{ background: transparent; }}
    QScrollArea {{ border: none; background: transparent; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QFrame#Card {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 12px; }}
    QFrame#VSep {{ background: {C.BORDER}; border: none; }}
    QLabel#Title {{ font-size: 16pt; font-weight: 800; }}
    QLabel#Meta {{ color: {C.MUTED}; font-size: 10pt; }}
    QLabel#StatLabel {{ color: {C.MUTED}; font-size: 10pt; font-weight: 600; }}
    QLabel#StatValue {{ font-size: 15pt; font-weight: 800; }}
    QLabel#SectionTitle {{ font-size: 12pt; font-weight: 800; }}
    QLabel#SectionMeta {{ color: {C.MUTED}; font-size: 10pt; }}
    QLabel#Muted {{ color: {C.MUTED}; }}
    QLabel#EmptyTitle {{ font-size: 13pt; font-weight: 800; }}

    QPushButton {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px;
        padding: 6px 14px; color: {C.TEXT}; font-weight: 600;
    }}
    QPushButton:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton:pressed {{ background: {C.BORDER}; }}
    QPushButton:disabled {{ color: {C.FAINT}; background: {C.SURFACE}; border-color: {C.BORDER}; }}
    QPushButton::menu-indicator {{ subcontrol-position: right center; right: 8px; width: 8px; }}
    QPushButton#Primary {{ background: {C.ACCENT}; border: 1px solid {C.ACCENT}; color: white; font-weight: 700; }}
    QPushButton#Primary:hover {{ background: {C.ACCENT_HI}; border-color: {C.ACCENT_HI}; }}
    QPushButton#Primary:disabled {{ background: {C.ACCENT_LINE}; border-color: {C.ACCENT_LINE}; color: white; }}
    QPushButton#IconBtn {{ background: transparent; border: none; border-radius: 6px; padding: 0; }}
    QPushButton#IconBtn:hover {{ background: {C.HOVER}; }}
    QPushButton#IconBtn:disabled {{ background: transparent; }}
    QPushButton#RowBtn {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 7px;
        padding: 0 10px; font-size: 10pt; font-weight: 600; color: {C.BODY};
    }}
    QPushButton#RowBtn:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton#RowDone {{
        background: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 7px;
        padding: 0 10px; font-size: 10pt; font-weight: 700; color: #047857;
    }}
    QPushButton#RowDone:hover {{ background: #D1FAE5; border-color: #6EE7B7; }}
    QPushButton#Danger {{ background: {C.SURFACE}; border: 1px solid #FBCFD8; color: #E11D48; font-weight: 700; }}
    QPushButton#Danger:hover {{ background: #FFF1F3; border-color: #F9A8B8; }}
    QPushButton#Link {{ background: transparent; border: none; color: {C.MUTED}; padding: 4px 2px; text-align: left; }}
    QPushButton#Link:hover {{ color: {C.TEXT}; }}
    QPushButton#QuickBtn {{
        background: {C.BG}; border: 1px solid {C.BORDER}; border-radius: 11px; padding: 3px 11px;
        font-size: 10pt; font-weight: 600; color: {C.BODY};
    }}
    QPushButton#QuickBtn:hover {{ border-color: {C.ACCENT_LINE}; color: {C.ACCENT_HI}; background: {C.ACCENT_SOFT}; }}
    QTextBrowser#Answer {{
        background: {C.BG}; border: none; border-radius: 8px; padding: 8px 10px; color: {C.TEXT};
    }}
    QFrame#Proposal {{ background: {C.ACCENT_SOFT}; border: 1px solid {C.ACCENT_LINE}; border-radius: 8px; }}

    QFrame#Nav {{ background: transparent; border: none; border-radius: 10px; }}
    QFrame#Nav:hover {{ background: {C.HOVER}; }}
    QFrame#Nav[active="true"] {{ background: {C.ACCENT_SOFT}; }}
    QLabel#NavName {{ font-weight: 700; }}
    QFrame#Nav[active="true"] QLabel#NavName {{ color: {C.ACCENT_HI}; }}
    QLabel#NavMeta {{ color: {C.MUTED}; font-size: 10pt; }}

    QTableWidget {{
        background: {C.SURFACE}; border: none; gridline-color: transparent;
        selection-background-color: transparent; selection-color: {C.TEXT}; outline: 0;
    }}
    QTableWidget::item {{ border-bottom: 1px solid {C.LINE}; padding: 0 10px; }}
    QTableWidget::item:selected {{ background: transparent; color: {C.TEXT}; }}
    QTableWidget QLineEdit {{ border: 1px solid {C.ACCENT}; border-radius: 4px; padding: 2px 6px; }}
    QHeaderView::section {{
        background: {C.SURFACE}; color: {C.MUTED}; padding: 8px 10px; border: none;
        border-bottom: 1px solid {C.BORDER}; font-weight: 700; font-size: 10pt;
    }}
    QLineEdit, QComboBox {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px; padding: 6px 9px;
    }}
    QLineEdit:focus, QComboBox:focus {{ border: 1px solid {C.ACCENT}; }}
    QRadioButton {{ spacing: 8px; }}
    QProgressBar {{ background: {C.LINE}; border: none; border-radius: 3px; }}
    QProgressBar::chunk {{ background: {C.ACCENT}; border-radius: 3px; }}
    QMenu {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: 8px; padding: 6px; }}
    QMenu::item {{ padding: 7px 18px 7px 12px; border-radius: 5px; color: {C.TEXT}; }}
    QMenu::item:selected {{ background: {C.HOVER}; }}
    QMenu::item:disabled {{ color: {C.FAINT}; }}
    QMenu::separator {{ height: 1px; background: {C.BORDER}; margin: 5px 6px; }}
    QToolTip {{ background: {C.TEXT}; color: white; border: none; padding: 5px 8px; border-radius: 4px; }}
    QScrollBar:vertical {{ width: 10px; background: transparent; margin: 2px; }}
    QScrollBar::handle:vertical {{ background: {C.BORDER_HI}; border-radius: 4px; min-height: 24px; }}
    QScrollBar::handle:vertical:hover {{ background: {C.FAINT}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar:horizontal {{ height: 10px; background: transparent; margin: 2px; }}
    QScrollBar::handle:horizontal {{ background: {C.BORDER_HI}; border-radius: 4px; min-width: 24px; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
"""

_STATE_FILE = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'budget_state.json')

_ROW_H = 44

# 열: 원가통계비목 · 산출내역 · 예산현액 · 집행액 · 잔액 · 집행률 · [수정][완료]
# 넉넉할 때의 폭과, 좁아져도 이 아래로는 줄이지 않는 최소 폭(100% 배율 기준, _sizes()가 배율을 곱함).
# 집행률·버튼 칸은 글꼴 폭으로 계산한다. 남는 폭은 산출내역이 가져간다.
_COL_W = (150, 250, 118, 118, 132, 0)
_COL_MIN = (90, 110, 96, 96, 0, 0)
_COLS = ('원가통계비목', '산출내역', '예산현액', '잔액', '집행률', '')

# 집행률 수준: 정상(포인트 색) · 80% 이상(주황) · 100% 이상(빨강) · 완료(회색)
_LEVEL_BAR = {'ok': C.ACCENT, 'warn': '#D97706', 'over': '#E11D48', 'done': C.BORDER_HI}
_LEVEL_TEXT = {'ok': C.ACCENT_HI, 'warn': '#B45309', 'over': C.NEG, 'done': C.FAINT}
_LEVEL_SOFT = {'ok': C.ACCENT_SOFT, 'warn': '#FEF3E2', 'over': '#FDECEF', 'done': C.HOVER}


def _level(pct, done=False):
    if done:
        return 'done'
    if pct >= 100:
        return 'over'
    if pct >= 80:
        return 'warn'
    return 'ok'


def _bar_qss(color):
    return (f'QProgressBar {{ background: {C.LINE}; border: none; border-radius: 3px; }}'
            f'QProgressBar::chunk {{ background: {color}; border-radius: 3px; }}')


def _fmt(n):
    try:
        return f'{int(n):,}'
    except (TypeError, ValueError):
        return '0'


def _parse_amount(text):
    digits = ''.join(ch for ch in str(text) if ch.isdigit())
    return int(digits) if digits else 0


def _pct(spent, budget):
    return (spent / budget * 100) if budget > 0 else 0.0


def _attach_amount_format(edit):
    def on_change(text):
        digits = ''.join(ch for ch in text if ch.isdigit())
        formatted = f'{int(digits):,}' if digits else ''
        if formatted != text:
            edit.blockSignals(True)
            edit.setText(formatted)
            edit.blockSignals(False)
            edit.setCursorPosition(len(formatted))
    edit.textChanged.connect(on_change)


def _make_strike_font(font):
    f = QFont(font)
    f.setStrikeOut(True)
    return f


def _icon_btn(icon_name, tooltip, color=C.MUTED, size=16):
    btn = QPushButton()
    btn.setObjectName('IconBtn')
    btn.setIcon(qta.icon(icon_name, color=color))
    btn.setIconSize(QSize(size, size))
    btn.setFixedSize(28, 28)
    btn.setToolTip(tooltip)
    return btn


_EXCEL_EXTS = ('.xls', '.xlsx')


def _cell_text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def _to_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(round(value))
    text = str(value).strip().replace(',', '').replace(' ', '')
    negative = False
    if text.startswith('(') and text.endswith(')'):
        negative = True
        text = text[1:-1]
    if text in ('', '-', '–', '—'):
        return None
    try:
        num = int(round(float(text)))
    except ValueError:
        return None
    return -num if negative else num


def _read_excel_sheets(path):
    """엑셀 파일을 [(시트이름, rows)] 형태로 읽는다. (.xls: xlrd, .xlsx: openpyxl)"""
    ext = os.path.splitext(path)[1].lower()
    sheets = []
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            rows = [[sh.cell_value(r, c) for c in range(sh.ncols)]
                    for r in range(sh.nrows)]
            sheets.append((sh.name, rows))
    elif ext == '.xlsx':
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            for sh in wb.worksheets:
                rows = [list(row) for row in sh.iter_rows(values_only=True)]
                sheets.append((sh.title, rows))
        finally:
            wb.close()
    else:
        raise ValueError('지원하지 않는 파일 형식입니다. (.xls, .xlsx)')
    return sheets


def _parse_budget_sheet(rows):
    """사업관리카드(예산) 시트를 세부항목/항목 구조로 파싱. 헤더 없으면 None."""
    header_row = None
    col_cost = col_desc = col_budget = col_spent = None
    for r in range(min(len(rows), 30)):
        texts = [_cell_text(v).replace(' ', '') for v in rows[r]]
        budget_idx = next((i for i, t in enumerate(texts) if '예산현액' in t), None)
        desc_idx = next((i for i, t in enumerate(texts) if '산출내역' in t), None)
        if budget_idx is None or desc_idx is None:
            continue
        header_row = r
        col_budget = budget_idx
        col_desc = desc_idx
        col_cost = next((i for i, t in enumerate(texts)
                         if '원가통계비목' in t or '세부사업' in t), 0)
        # 하위 헤더 행(집행 세부 열)까지 합쳐서 집행 열을 찾는다
        merged = list(texts)
        for extra in rows[r + 1:r + 3]:
            for i in range(min(len(merged), len(extra))):
                merged[i] += _cell_text(extra[i]).replace(' ', '')
        col_spent = next((i for i, t in enumerate(merged) if '원인행위' in t), None)
        if col_spent is None:
            col_spent = next((i for i, t in enumerate(merged) if '지출결의' in t), None)
        if col_spent is None:
            col_spent = next((i for i, t in enumerate(merged) if '지급액' in t), None)
        break
    if header_row is None:
        return None

    def cell(row, idx):
        return row[idx] if idx is not None and idx < len(row) else None

    categories = []
    current = None
    for row in rows[header_row + 1:]:
        cost = _cell_text(cell(row, col_cost))
        desc = _cell_text(cell(row, col_desc))
        budget = _to_number(cell(row, col_budget))
        spent = _to_number(cell(row, col_spent)) or 0
        if not cost and not desc and budget is None:
            continue
        if any(key in cost for key in ('합계', '총계', '소계')):
            continue
        if not desc and budget is not None:
            # 세부항목(카테고리) 소계 행
            current = {'name': cost or '미분류', 'items': []}
            categories.append(current)
            continue
        if desc:
            if current is None:
                current = {'name': '미분류', 'items': []}
                categories.append(current)
            current['items'].append({
                'cost_type': cost or '일반',
                'description': desc,
                'budget': budget or 0,
                'spent': spent,
            })
    return categories


def parse_budget_excel(path):
    """엑셀 파일 → [{'name': 세부항목, 'items': [...]}]. 실패 시 ValueError."""
    try:
        sheets = _read_excel_sheets(path)
    except ImportError as e:
        raise ValueError(f'엑셀 읽기 모듈을 불러오지 못했습니다: {e}')
    for _name, rows in sheets:
        if not rows:
            continue
        categories = _parse_budget_sheet(rows)
        if categories:
            return categories
    raise ValueError("'예산현액'/'산출내역' 헤더를 찾지 못했습니다.")


def _parse_tsv(text):
    """엑셀 클립보드 텍스트(탭 구분, 따옴표 셀 처리)를 행렬로 변환."""
    rows = []
    row = []
    cur = []
    in_quotes = False
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if in_quotes:
            if ch == '"':
                if i + 1 < n and text[i + 1] == '"':
                    cur.append('"')
                    i += 1
                else:
                    in_quotes = False
            else:
                cur.append(ch)
        elif ch == '"' and not cur:
            in_quotes = True
        elif ch == '\t':
            row.append(''.join(cur))
            cur = []
        elif ch == '\n':
            row.append(''.join(cur))
            cur = []
            rows.append(row)
            row = []
        elif ch != '\r':
            cur.append(ch)
        i += 1
    if cur or row:
        row.append(''.join(cur))
        rows.append(row)
    return rows


def _parse_budget_rows_no_header(rows):
    """헤더 없이 데이터만 복사된 경우: A=원가통계비목, B=산출내역,
    C=예산현액, E=원인행위금액(B) 위치 기준으로 파싱."""
    categories = []
    current = None

    def cell(row, idx):
        return row[idx] if idx < len(row) else None

    for row in rows:
        cost = _cell_text(cell(row, 0))
        desc = _cell_text(cell(row, 1))
        budget = _to_number(cell(row, 2))
        spent = _to_number(cell(row, 4)) or 0
        if not cost and not desc and budget is None:
            continue
        if any(key in cost for key in ('합계', '총계', '소계')):
            continue
        if not desc and budget is not None:
            current = {'name': cost or '미분류', 'items': []}
            categories.append(current)
            continue
        if desc:
            if current is None:
                current = {'name': '미분류', 'items': []}
                categories.append(current)
            current['items'].append({
                'cost_type': cost or '일반',
                'description': desc,
                'budget': budget or 0,
                'spent': spent,
            })
    if not categories or not any(c['items'] for c in categories):
        return None
    return categories


def parse_budget_clipboard(text):
    """클립보드 텍스트(엑셀 셀 복사) → [{'name': 세부항목, 'items': [...]}]."""
    text = (text or '').replace('\ufeff', '')
    if '\t' not in text:
        raise ValueError('탭으로 구분된 표 형식이 아닙니다.\n'
                         '엑셀에서 셀 범위(A~E열)를 복사한 뒤 다시 시도해주세요.')
    rows = _parse_tsv(text)
    categories = _parse_budget_sheet(rows) or _parse_budget_rows_no_header(rows)
    if not categories or not any(c['items'] for c in categories):
        raise ValueError('예산 데이터를 찾지 못했습니다.\n'
                         '세부항목 소계 행 포함, A열~E열 범위를 복사했는지 확인해주세요.')
    return categories


# ── 표 열 폭 / 높이 자동 맞춤 ──────────────────────────────────────
_SIZES = None


def _sizes():
    """열 폭·버튼 크기 (화면 배율과 실제 글꼴 폭으로 계산, 한 번만).

    픽셀로 고정하면 Windows 125% 배율에서 글자만 커져 집행률 '100.0%'가 잘린다.
    """
    global _SIZES
    if _SIZES is None:
        # 실제 화면과 같은 스타일을 입힌 견본으로 잰다 (QFont만 만들어 재면 Windows에서 실제 글꼴과 폭이 달랐다)
        sample = QLabel('100.0%')
        sample.setStyleSheet(STYLE + 'QLabel { font-weight: 700; }')
        sample.ensurePolished()
        fm = sample.fontMetrics()
        btn = QPushButton('완료 취소')
        btn.setObjectName('RowBtn')
        btn.setStyleSheet(STYLE)
        btn.ensurePolished()
        fs = btn.fontMetrics()
        pct_lbl = fm.horizontalAdvance('100.0%') + px(6)
        pad = px(20)                       # 표 칸 안쪽 여백(QSS 'padding: 0 10px')만큼 칸 속 위젯이 좁아진다
        pct_col = px(6) + px(40) + px(6) + pct_lbl + px(10) + pad
        btn_w = max(px(62), fs.horizontalAdvance('완료') + px(36))          # [수정] [완료]
        btn_wide = max(btn_w, fs.horizontalAdvance('완료 취소') + px(36))    # 완료한 줄의 [완료 취소]
        btn_h = max(px(28), fs.height() + px(12))
        act_col = btn_w + btn_wide + px(6) + px(16) + pad
        # 금액 칸: 글꼴이 바뀌어도(맑은 고딕 ↔ Pretendard) 9자리 금액까지 안 잘리게 실제 글꼴로 잰다
        num_w = fm.horizontalAdvance('999,999,999') + pad + px(6)       # 넉넉할 때: 9자리 금액
        num_min = fm.horizontalAdvance('99,999,999') + pad              # 줄여도 이 아래로는 안 줄임: 8자리 금액
        w = [px(x) for x in _COL_W[:4]]
        m = [px(x) for x in _COL_MIN[:4]]
        for i in (2, 3):
            w[i], m[i] = max(w[i], num_w), max(m[i], num_min)
        w, m = tuple(w) + (max(px(_COL_W[4]), pct_col), act_col), tuple(m) + (pct_col, act_col)
        # 왼쪽 세부항목 목록: 가장 긴 보조 글과 퍼센트가 안 잘리게
        meta = QLabel('잔액 99,999,999원 · 2개 세부항목')
        meta.setObjectName('NavMeta')
        meta.setStyleSheet(STYLE)
        meta.ensurePolished()
        nav_pct = fm.horizontalAdvance('100%') + px(10)
        nav_w = max(px(220), meta.fontMetrics().horizontalAdvance(meta.text()) + px(40))
        _SIZES = {'w': w, 'min': m, 'pct_lbl': pct_lbl, 'btn_w': btn_w, 'btn_wide': btn_wide, 'btn_h': btn_h,
                  'row_h': max(px(_ROW_H), btn_h + px(12)), 'nav_w': nav_w, 'nav_pct': nav_pct}
    return _SIZES


def _cell_widgets_width(table, col):
    """그 칸에 들어간 위젯들이 바라는 폭 + 표 칸 안쪽 여백."""
    best = 0
    for r in range(table.rowCount()):
        w = table.cellWidget(r, col)
        if w is not None and not table.isRowHidden(r):
            best = max(best, w.sizeHint().width())
    return best + px(20) if best else 0


def _fit_columns(table):
    """표 폭에 맞춰 열 폭을 배분한다(넘치면 최소 폭까지 비례 축소, 남으면 산출내역이 가져감)."""
    hdr = table.horizontalHeader()
    avail = table.viewport().width()
    col_w, col_min = list(_sizes()['w']), list(_sizes()['min'])
    for col in (4, 5):                     # 집행률·버튼 칸: 화면에 실제로 그려질 크기(글꼴·배율 반영)에 맞춤
        need = _cell_widgets_width(table, col)
        col_min[col] = max(col_min[col], need)
        col_w[col] = max(col_w[col], need)
    n = len(col_w)
    if avail <= 0 or hdr.count() < n:
        return
    total = sum(col_w)
    if avail >= total:
        widths = list(col_w)
        widths[1] += avail - total
    else:
        slack = [max(0, col_w[i] - col_min[i]) for i in range(n)]
        pool = sum(slack)
        if pool <= 0:
            return
        cut = min(total - avail, pool)
        widths = [max(col_min[i], int(col_w[i] - cut * slack[i] / pool)) for i in range(n)]
        over = sum(widths) - avail
        while over > 0:
            i = max(range(n), key=lambda k: widths[k] - col_min[k])
            if widths[i] <= col_min[i]:
                break
            widths[i] -= 1
            over -= 1
    for i, wd in enumerate(widths):
        if hdr.sectionSize(i) != wd:
            hdr.resizeSection(i, wd)


class _FlowLayout(QLayout):
    """폭이 모자라면 다음 줄로 넘어가는 가로 배치 (원가통계비목 칩)."""

    def __init__(self, parent=None, spacing=6):
        super().__init__(parent)
        self._items = []
        self._sp = spacing

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, i):
        return self._items[i] if 0 <= i < len(self._items) else None

    def takeAt(self, i):
        return self._items.pop(i) if 0 <= i < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientations(Qt.Orientation(0))

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._layout(QRect(0, 0, width, 0), dry=True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._layout(rect)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for it in self._items:
            size = size.expandedTo(it.minimumSize())
        m = self.contentsMargins()
        return size + QSize(m.left() + m.right(), m.top() + m.bottom())

    def _layout(self, rect, dry=False):
        m = self.contentsMargins()
        r = rect.adjusted(m.left(), m.top(), -m.right(), -m.bottom())
        x, y, line_h = r.x(), r.y(), 0
        for it in self._items:
            hint = it.sizeHint()
            if x + hint.width() > r.right() + 1 and line_h > 0:
                x, y = r.x(), y + line_h + self._sp
                line_h = 0
            if not dry:
                it.setGeometry(QRect(QPoint(x, y), hint))
            x += hint.width() + self._sp
            line_h = max(line_h, hint.height())
        return y + line_h - rect.y() + m.bottom()


def _cost_totals(items):
    """원가통계비목별 합계 {비목: [예산, 집행, 항목 수]} (처음 나온 순서)."""
    out = {}
    for it in items:
        t = out.setdefault(it['cost_type'] or '일반', [0, 0, 0])
        t[0] += it['budget']
        t[1] += it['spent']
        t[2] += 1
    return out


def _cost_chip_text(cost, budget, spent):
    return f'{cost}   잔액 {_fmt(budget - spent)}원 · {_pct(spent, budget):.0f}%'


def _cost_chip_tip(cost, budget, spent, n):
    return (f'{cost} ({n}개 항목)\n예산 {_fmt(budget)}원\n집행 {_fmt(spent)}원\n'
            f'잔액 {_fmt(budget - spent)}원 · 집행률 {_pct(spent, budget):.1f}%')


def _cost_chip_qss(budget, spent, selector='QPushButton'):
    lvl = _level(_pct(spent, budget))
    neg = budget - spent < 0
    return (f'{selector} {{ background: {_LEVEL_SOFT[lvl]}; color: {C.NEG if neg else _LEVEL_TEXT[lvl]};'
            f' border: 1px solid transparent; border-radius: 11px; padding: 4px 11px;'
            f' font-weight: 600; font-size: 10pt; }}'
            f'{selector}:hover {{ border-color: {C.BORDER_HI}; }}'
            f'{selector}:checked {{ background: {C.SURFACE}; border: 1px solid {C.ACCENT}; }}')


class _BudgetTable(QTableWidget):
    """세부항목 하나의 표. 폭에 맞춰 열을 나누고 높이를 내용에 맞춘다(안쪽 스크롤 없음)."""

    def __init__(self):
        super().__init__()
        self.action_buttons = []           # 행마다 (수정, 완료)
        self.row_costs = []
        self.pct_labels = []

    def _sync_pct_labels(self):
        """집행률 글자 칸 폭을 창에 입혀진 실제 글꼴로 다시 잰다 (모든 줄을 같은 폭으로 → 막대가 나란히)."""
        if not self.pct_labels or not self.isVisible():
            return
        need = self.pct_labels[0].fontMetrics().horizontalAdvance('100.0%') + px(6)
        for lb in self.pct_labels:
            if lb.width() != need:
                lb.setFixedWidth(need)

    def sync_geometry(self):
        self._sync_pct_labels()
        _fit_columns(self)
        hdr = self.horizontalHeader()
        rows_h = sum(self.rowHeight(r) for r in range(self.rowCount()) if not self.isRowHidden(r))
        head_h = 0 if hdr.isHidden() else max(hdr.height(), hdr.sizeHint().height())
        cols = sum(hdr.sectionSize(c) for c in range(hdr.count()))
        hbar = self.horizontalScrollBar().sizeHint().height() if cols > self.viewport().width() > 0 else 0
        target = head_h + rows_h + hbar + 2
        if self.height() != target:
            self.setFixedHeight(target)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.sync_geometry()


class _NavItem(QFrame):
    """왼쪽 세부항목 목록의 한 줄: 이름 · 집행률 막대 · 잔액."""
    clicked = pyqtSignal()

    def __init__(self, name):
        super().__init__()
        self.setObjectName('Nav')
        self.setProperty('active', 'false')
        self.setCursor(Qt.PointingHandCursor)
        v = QVBoxLayout(self)
        v.setContentsMargins(10, 8, 10, 8)
        v.setSpacing(5)
        self.lbl_name = QLabel(name)
        self.lbl_name.setObjectName('NavName')
        self.lbl_name.setWordWrap(True)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(4)
        self.lbl_pct = QLabel()
        self.lbl_pct.setObjectName('NavMeta')
        self.lbl_pct.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lbl_pct.setFixedWidth(_sizes()['nav_pct'])
        row.addWidget(self.bar, 1)
        row.addWidget(self.lbl_pct)
        self.lbl_meta = QLabel()
        self.lbl_meta.setObjectName('NavMeta')
        v.addWidget(self.lbl_name)
        v.addLayout(row)
        v.addWidget(self.lbl_meta)

    def set_active(self, on):
        self.setProperty('active', 'true' if on else 'false')
        for w in (self, self.lbl_name):
            w.style().unpolish(w)
            w.style().polish(w)

    def set_values(self, budget, spent, extra=''):
        pct = _pct(spent, budget)
        lvl = _level(pct)
        self.bar.setValue(int(round(min(pct, 100))))
        self.bar.setStyleSheet(_bar_qss(_LEVEL_BAR[lvl]))
        self.lbl_pct.setText(f'{pct:.0f}%')
        self.lbl_pct.setStyleSheet(f'color: {_LEVEL_TEXT[lvl]}; font-weight: 700;')
        rem = budget - spent
        color = C.NEG if rem < 0 else C.MUTED
        self.lbl_meta.setText(f'<span style="color:{color}">잔액 {_fmt(rem)}원</span>{extra}')

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()


_IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
_IMAGE_METHOD = {
    'ai-image': 'AI가 이미지를 읽어 정리',
    'ai-text': 'PC 글자 인식 + AI 정리',
    'ocr': 'PC 글자 인식 (AI 없음) — 숫자를 꼭 확인하세요',
}


class BudgetWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        _load_pretendard()
        self.categories = []
        self.items = []
        self.active_tab = 'ALL'
        self._overview_open = False
        self._items_by_id = {}
        self._row_refs = {}
        self._sec_refs = {}
        self._nav_refs = {}
        self._hide_done = set()           # 완료 항목을 접어 둔 세부항목 (기본은 펼침)
        self._cost_filter = {}            # 세부항목별로 골라 본 원가통계비목 (없으면 전부)
        self._overview_chips = {}
        self._img_worker = None
        self._ask_worker = None
        self._ask_history = []            # [(질문, 답)] — 이어지는 질문에 쓰려고 최근 것만
        self._ask_q = ''
        self._pending_changes = []        # AI가 제안한 변경 (사용자가 [넣기]를 눌러야 반영)
        self._undo_snapshot = None        # 마지막 [넣기] 전 (집행액 {item_id: spent}, 새로 적은 사용 내역 id들) (되돌리기용, 1회분)
        self._expenses_by_item = {}       # 항목별 사용 내역 (최근 날짜 먼저)
        self._expense_ui = {}             # 이번 창을 여는 동안만 기억: 펼침 · 더 보기 (창을 닫으면 사라짐)
        self._restore_window_state()
        self._init_ui()
        self.refresh()
        keep_light(self)  # 예산 창은 라이트 전용 (다크 테마에서도 원래 모습 유지)

    # ── UI 구성 ────────────────────────────────────────────────────
    def _init_ui(self):
        self.setWindowTitle('예산 관리')
        self.setMinimumSize(px(1000), px(600))
        self.setStyleSheet(STYLE)
        self.setAcceptDrops(True)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(14)

        # 머리: 제목 · 요약 글 | 가져오기 ▾ · 항목 추가 · ⋯
        header = QHBoxLayout()
        header.setSpacing(8)
        tb = QVBoxLayout()
        tb.setSpacing(2)
        title = QLabel('예산 관리')
        title.setObjectName('Title')
        self.lbl_meta = QLabel()
        self.lbl_meta.setObjectName('Meta')
        tb.addWidget(title)
        tb.addWidget(self.lbl_meta)
        header.addLayout(tb)
        header.addStretch()

        self.btn_import = QPushButton('가져오기')
        self.btn_import.setIcon(qta.icon('fa5s.file-import', color=C.MUTED))
        self.btn_import.setIconSize(QSize(13, 13))
        self.btn_import.setToolTip('엑셀 파일, 엑셀에서 복사한 표, 예산 화면 캡처를 가져와 세부항목별로 정리')
        self.btn_import.clicked.connect(lambda: self._exec_menu(self._import_menu(), self.btn_import))
        self.btn_manage = QPushButton('세부항목 관리')
        self.btn_manage.setIcon(qta.icon('fa5s.folder-open', color=C.MUTED))
        self.btn_manage.setIconSize(QSize(11, 11))
        self.btn_manage.setToolTip('세부항목을 만들고 이름을 바꾸거나 지웁니다')
        self.btn_manage.clicked.connect(lambda: self._open_category_manager())
        self.btn_more = QPushButton()
        self.btn_more.setObjectName('IconBtn')
        self.btn_more.setIcon(qta.icon('fa5s.ellipsis-h', color=C.MUTED))
        self.btn_more.setFixedSize(34, 34)
        self.btn_more.setToolTip('원가통계비목 목록 · 전체 지우기')
        self.btn_more.clicked.connect(lambda: self._exec_menu(self._more_menu(), self.btn_more))
        for b in (self.btn_import, self.btn_manage):
            b.setFixedHeight(34)
            b.setCursor(Qt.PointingHandCursor)
        header.addWidget(self.btn_import)
        header.addWidget(self.btn_manage)
        header.addWidget(self.btn_more)
        root.addLayout(header)

        # 요약 한 줄: 예산현액 | 집행액 | 잔액 ……… 집행률 ▓▓▓░
        self.summary = QFrame()
        self.summary.setObjectName('Card')
        sl = QHBoxLayout(self.summary)
        sl.setContentsMargins(22, 14, 22, 14)
        sl.setSpacing(22)
        self.stat = {}
        for key, label in (('budget', '예산현액'), ('spent', '집행액'), ('remaining', '잔액')):
            box = QVBoxLayout()
            box.setSpacing(2)
            lb = QLabel(label)
            lb.setObjectName('StatLabel')
            val = QLabel('0원')
            val.setObjectName('StatValue')
            box.addWidget(lb)
            box.addWidget(val)
            sl.addLayout(box)
            self.stat[key] = val
            if key != 'remaining':
                sep = QFrame()
                sep.setObjectName('VSep')
                sep.setFixedWidth(1)
                sl.addWidget(sep)
        sl.addStretch()
        pct_box = QWidget()
        pct_box.setFixedWidth(260)
        pv = QVBoxLayout(pct_box)
        pv.setContentsMargins(0, 0, 0, 0)
        pv.setSpacing(8)
        top = QHBoxLayout()
        lb = QLabel('집행률')
        lb.setObjectName('StatLabel')
        self.lbl_pct = QLabel('0.0%')
        top.addWidget(lb)
        top.addStretch()
        top.addWidget(self.lbl_pct)
        self.bar_total = QProgressBar()
        self.bar_total.setRange(0, 100)
        self.bar_total.setTextVisible(False)
        self.bar_total.setFixedHeight(6)
        pv.addLayout(top)
        pv.addWidget(self.bar_total)
        sl.addWidget(pct_box)
        root.addWidget(self.summary)

        # 본문: 왼쪽 세부항목 목록 | 오른쪽 표
        self.body = QWidget()
        bl = QHBoxLayout(self.body)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(14)
        self.nav_card = QFrame()
        self.nav_card.setObjectName('Card')
        self.nav_card.setFixedWidth(_sizes()['nav_w'])
        nl = QVBoxLayout(self.nav_card)
        nl.setContentsMargins(8, 12, 8, 8)
        nl.setSpacing(4)
        cap = QLabel('  세부항목')
        cap.setObjectName('StatLabel')
        nl.addWidget(cap)
        nav_scroll = QScrollArea()
        nav_scroll.setWidgetResizable(True)
        nav_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        nav_box = QWidget()
        self.nav_layout = QVBoxLayout(nav_box)
        self.nav_layout.setContentsMargins(0, 0, 0, 0)
        self.nav_layout.setSpacing(2)
        nav_scroll.setWidget(nav_box)
        nl.addWidget(nav_scroll, 1)
        bl.addWidget(self.nav_card)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.sections_container = QWidget()
        self.sections_layout = QVBoxLayout(self.sections_container)
        self.sections_layout.setContentsMargins(0, 0, 4, 0)
        self.sections_layout.setSpacing(14)
        self.scroll.setWidget(self.sections_container)
        bl.addWidget(self.scroll, 1)
        root.addWidget(self.body, 1)

        self.empty = self._make_empty()
        root.addWidget(self.empty, 1)

        self.ask_card = self._make_ask_panel()
        root.addWidget(self.ask_card)

    # ── AI에게 묻기 ────────────────────────────────────────────────
    def _make_ask_panel(self):
        from budget_ai import QUICK
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(16, 12, 16, 12)
        v.setSpacing(8)

        self.ask_answer = QTextBrowser()
        self.ask_answer.setObjectName('Answer')
        self.ask_answer.setOpenLinks(False)
        self.ask_answer.hide()
        v.addWidget(self.ask_answer)

        self.ask_proposal = QFrame()
        self.ask_proposal.setObjectName('Proposal')
        pl = QHBoxLayout(self.ask_proposal)
        pl.setContentsMargins(12, 8, 8, 8)
        pl.setSpacing(8)
        self.lbl_proposal = QLabel()
        self.lbl_proposal.setTextFormat(Qt.RichText)
        self.lbl_proposal.setWordWrap(True)
        b_apply = QPushButton('넣기')
        b_apply.setObjectName('Primary')
        b_apply.setCursor(Qt.PointingHandCursor)
        b_apply.clicked.connect(self._apply_spend)
        self.btn_apply_spend = b_apply
        b_skip = QPushButton('취소')
        b_skip.setCursor(Qt.PointingHandCursor)
        b_skip.clicked.connect(self._hide_proposal)
        pl.addWidget(self.lbl_proposal, 1)
        pl.addWidget(b_apply)
        pl.addWidget(b_skip)
        self.ask_proposal.hide()
        v.addWidget(self.ask_proposal)

        quick = QHBoxLayout()
        quick.setSpacing(6)
        icon = QLabel()
        icon.setPixmap(qta.icon('fa5s.comment-dots', color=C.ACCENT).pixmap(15, 15))
        title = QLabel('AI에게 묻기')
        title.setStyleSheet(f'color: {C.ACCENT}; font-weight: 700;')
        quick.addWidget(icon)
        quick.addWidget(title)
        quick.addSpacing(6)
        self.quick_buttons = []
        for label, question in QUICK:
            b = QPushButton(label)
            b.setObjectName('QuickBtn')
            b.setCursor(Qt.PointingHandCursor)
            b.setToolTip(question)
            b.clicked.connect(lambda _, q=question: self._ask(q))
            quick.addWidget(b)
            self.quick_buttons.append(b)
        quick.addStretch()
        self.ask_status = QLabel()
        self.ask_status.setObjectName('Muted')
        quick.addWidget(self.ask_status)
        self.btn_undo_apply = QPushButton('되돌리기')
        self.btn_undo_apply.setObjectName('QuickBtn')
        self.btn_undo_apply.setCursor(Qt.PointingHandCursor)
        self.btn_undo_apply.setToolTip('방금 [넣기]로 바꾼 집행액을 전 값으로 되돌립니다')
        self.btn_undo_apply.clicked.connect(self._undo_apply)
        self.btn_undo_apply.hide()
        quick.addWidget(self.btn_undo_apply)
        v.addLayout(quick)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.ask_input = QLineEdit()
        self.ask_input.setFixedHeight(36)
        self.ask_input.returnPressed.connect(lambda: self._ask(self.ask_input.text()))
        self.ask_send = QPushButton()
        self.ask_send.setObjectName('Primary')
        self.ask_send.setIcon(qta.icon('fa5s.paper-plane', color='white'))
        self.ask_send.setFixedSize(40, 36)
        self.ask_send.setToolTip('보내기 (Enter)')
        self.ask_send.setCursor(Qt.PointingHandCursor)
        self.ask_send.clicked.connect(lambda: self._ask(self.ask_input.text()))
        row.addWidget(self.ask_input, 1)
        row.addWidget(self.ask_send)
        v.addLayout(row)
        return card

    def _ai_ready(self):
        try:
            from ai_client import load_ai_mode
            return load_ai_mode() != 'none'
        except Exception:
            return False

    def _update_ask_enabled(self):
        busy = self._ask_worker is not None and self._ask_worker.isRunning()
        ready = self._ai_ready()
        for w in [self.ask_input, self.ask_send] + self.quick_buttons:
            w.setEnabled(ready and not busy)
        self.ask_input.setPlaceholderText(
            '예산에 대해 묻거나 지시하세요 — 예: 협의회비 잔액 30만원으로 맞춰줘'
            if ready else 'AI가 꺼져 있어요 — 공문 작성 창 → [AI 설정]에서 연결하면 쓸 수 있어요')

    def _ask(self, question):
        q = (question or '').strip()
        if not q or (self._ask_worker is not None and self._ask_worker.isRunning()):
            return
        from budget_ai import BudgetAskWorker
        self._ask_q = q
        self.ask_input.clear()
        self._hide_proposal()
        self._set_answer(f'Q. {q}\n\n생각하는 중…')
        self.ask_status.setText('AI가 예산 표를 보는 중…')
        w = BudgetAskWorker(q, self.categories, self.items, self._ask_history)
        w.done.connect(self._on_ask_done)
        w.failed.connect(self._on_ask_failed)
        self._ask_worker = w
        w.finished.connect(self._update_ask_enabled)
        w.start()
        self._update_ask_enabled()

    def _on_ask_done(self, reply, masked):
        answer = reply.get('answer') or ''
        self._ask_history = (self._ask_history + [(self._ask_q, answer)])[-2:]
        text = f'Q. {self._ask_q}\n\n{answer}'
        if masked:
            try:
                from privacy import describe
                text += f'\n\n🛡 개인정보 {describe(masked)}을(를) 가리고 보냈어요.'
            except Exception:
                pass
        self._set_answer(text)
        self.ask_status.setText('')
        self._show_proposal(reply.get('changes') or [], reply.get('rejected') or [])

    def _on_ask_failed(self, msg):
        self._set_answer(f'Q. {self._ask_q}\n\n⚠ {msg}')
        self.ask_status.setText('')

    def _set_answer(self, text):
        """답 칸: 글 길이에 맞춰 높이를 정하되 170px까지 (넘으면 스크롤)."""
        a = self.ask_answer
        a.setPlainText(text)
        a.show()
        doc = a.document()
        doc.setTextWidth(max(200, a.viewport().width() or self.width() - 90))
        a.setFixedHeight(int(min(170, max(52, doc.size().height() + 20))))

    # ── 말로 지출 입력: AI가 제안 → [넣기]를 눌러야 반영 ─────────────
    def _show_proposal(self, changes, rejected=()):
        """AI가 제안한 변경을 표로 보여 준다. 제외된 제안은 회색 줄로 이유와 함께."""
        self._pending_changes = list(changes)
        self.btn_apply_spend.setEnabled(bool(changes))
        if not changes and not rejected:
            self.ask_proposal.hide()
            return

        def rem_html(r):
            return f'<span style="color:{C.NEG}">{_fmt(r)}원</span>' if r < 0 else f'{_fmt(r)}원'

        cats = {c['id']: c['name'] for c in self.categories}
        parts = []
        if changes:
            parts.append('<b>이렇게 바꿀까요?</b>')
            parts.append("<table width='100%' cellspacing='0' cellpadding='3'>"
                         f"<tr style='color:{C.MUTED}'><td>산출내역</td><td>세부항목</td>"
                         "<td>집행액 (기존 → 변경)</td><td>잔액 (기존 → 변경)</td></tr>")
            for ch in changes:
                it = self._items_by_id.get(ch['item_id']) or ch['item']
                delta = (f" <span style='color:{C.MUTED}'>(+{_fmt(ch['value'])}원"
                         + (f" · {ch['purpose']}" if ch.get('purpose') else '') + ")</span>"
                         if ch['kind'] == 'spend' else '')
                parts.append(f"<tr><td><b>{it['description']}</b></td><td>{cats.get(it['category_id'], '')}</td>"
                             f"<td>{_fmt(ch['old_spent'])} → <b style='color:{C.ACCENT}'>{_fmt(ch['new_spent'])}원</b>{delta}</td>"
                             f"<td>{rem_html(ch['old_remaining'])} → {rem_html(ch['new_remaining'])}</td></tr>")
            parts.append('</table>')
        for rj in rejected:
            name = rj['item']['description'] if rj.get('item') else '알 수 없는 항목'
            parts.append(f"<span style='color:{C.FAINT}'>제외: {name} — {rj['reason']}</span>")
        if changes:
            parts.append(f"<span style='color:{C.MUTED}'>실제 집행 내역과 맞는지 확인하세요.</span>")
        self.lbl_proposal.setText('<br>'.join(parts))
        self.ask_proposal.show()

    def _hide_proposal(self):
        self._pending_changes = []
        self.ask_proposal.hide()

    def _apply_spend(self):
        """[넣기]: 바꾸기 전 집행액을 보관한 뒤 반영한다. 쓴 곳이 있는 지출은 사용 내역에도 남긴다.
        되돌리기는 마지막 1회분만."""
        from budget_ai import target_spent
        from datetime import date
        snapshot, expense_ids, done = {}, [], 0
        for ch in self._pending_changes:
            cur = self._items_by_id.get(ch['item_id'])          # 지금 값 기준으로 계산
            if not cur or cur['completed']:
                continue
            new = target_spent(ch['kind'], ch['value'], cur)
            if new < 0 or new == cur['spent']:
                continue
            snapshot[cur['id']] = cur['spent']
            if ch['kind'] == 'spend':
                expense_ids.append(add_budget_expense(cur['id'], ch['value'], ch.get('purpose', ''),
                                                      date.today().isoformat()))
            else:
                set_budget_item_spent(cur['id'], new)
            done += 1
        self._hide_proposal()
        self._undo_snapshot = (snapshot, expense_ids) if snapshot else None
        self.refresh()
        self.btn_undo_apply.setVisible(bool(snapshot))
        self.ask_status.setText(f'{done}개 항목에 넣었어요.' if done else '넣을 항목이 없어요.')

    def _undo_apply(self):
        snapshot, expense_ids = self._undo_snapshot or ({}, [])
        self._undo_snapshot = None
        n = 0
        for eid in expense_ids:
            delete_budget_expense(eid, adjust_spent=False)     # 기록만 지우고, 집행액은 아래에서 맞춤
        for iid, old in snapshot.items():
            if iid in self._items_by_id:
                set_budget_item_spent(iid, old)
                n += 1
        self.btn_undo_apply.hide()
        self.refresh()
        self.ask_status.setText(f'{n}개 항목을 넣기 전 값으로 되돌렸어요.' if n else '되돌릴 항목이 없어요.')

    def _make_empty(self):
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(40, 40, 40, 40)
        v.setSpacing(10)
        v.addStretch()
        icon = QLabel()
        icon.setPixmap(qta.icon('fa5s.file-import', color=C.BORDER_HI).pixmap(44, 44))
        icon.setAlignment(Qt.AlignCenter)
        t = QLabel('예산 자료를 가져오세요')
        t.setObjectName('EmptyTitle')
        t.setAlignment(Qt.AlignCenter)
        s = QLabel('사업관리카드 엑셀 파일을 이 창에 끌어다 놓거나,\n'
                   '예산 화면을 캡처해 붙여넣으면(우클릭 → 붙여넣기 또는 Ctrl+V) 세부항목별로 정리해 드려요.')
        s.setObjectName('Muted')
        s.setAlignment(Qt.AlignCenter)
        row = QHBoxLayout()
        row.addStretch()
        for text, icon_name, fn in (('엑셀 파일 열기', 'fa5s.file-excel', lambda: self._choose_excel()),
                                    ('붙여넣기', 'fa5s.paste', lambda: self._paste()),
                                    ('직접 만들기', 'fa5s.folder-plus', lambda: self._open_category_manager())):
            b = QPushButton(text)
            b.setIcon(qta.icon(icon_name, color=C.MUTED))
            b.setIconSize(QSize(13, 13))
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn)
            row.addWidget(b)
        row.addStretch()
        for w in (icon, t, s):
            v.addWidget(w)
        v.addSpacing(8)
        v.addLayout(row)
        v.addStretch()
        return card

    # ── 메뉴 ───────────────────────────────────────────────────────
    def _menu(self):
        m = QMenu(self)
        m.setStyleSheet(STYLE)
        return m

    @staticmethod
    def _exec_menu(menu, button):
        menu.exec_(button.mapToGlobal(button.rect().bottomLeft()))

    def _import_menu(self):
        m = self._menu()
        m.addAction(qta.icon('fa5s.file-excel', color=C.MUTED), '엑셀 파일 열기…', lambda: self._choose_excel())
        a = m.addAction(qta.icon('fa5s.paste', color=C.MUTED), '붙여넣기 — 엑셀 표 또는 캡처 이미지\tCtrl+V',
                        lambda: self._paste())
        a.setEnabled(self._clipboard_kind() is not None)
        return m

    def _more_menu(self):
        m = self._menu()
        m.addAction('원가통계비목 목록…', lambda: _CostTypeDialog(self).exec_())
        m.addSeparator()
        a = m.addAction('전체 지우기…', lambda: self._reset_all())
        a.setEnabled(bool(self.categories or self.items))
        return m

    def contextMenuEvent(self, e):
        m = self._import_menu()
        m.addSeparator()
        a = m.addAction('항목 추가…', lambda: self._open_add_item())
        a.setEnabled(bool(self.categories))
        m.exec_(e.globalPos())

    # ── 데이터 새로고침 ────────────────────────────────────────────
    def refresh(self, keep_scroll=True):
        pos = self.scroll.verticalScrollBar().value() if keep_scroll else 0
        self.categories = get_budget_categories()
        self.items = get_budget_items()
        self._items_by_id = {it['id']: it for it in self.items}
        self._expenses_by_item = {}
        for ex in get_budget_expenses():
            self._expenses_by_item.setdefault(ex['item_id'], []).append(ex)
        if self.active_tab != 'ALL' and self.active_tab not in {c['id'] for c in self.categories}:
            self.active_tab = 'ALL'
        has = bool(self.categories)
        self.summary.setVisible(has)
        self.body.setVisible(has)
        self.empty.setVisible(not has)
        self.ask_card.setVisible(has)
        self._update_ask_enabled()
        self._rebuild_nav()
        self._rebuild_sections()
        self._update_summary()
        if keep_scroll:
            self.scroll.verticalScrollBar().setValue(pos)

    def _update_meta(self):
        if not self.categories:
            self.lbl_meta.setText('세부항목을 가져오거나 직접 만들어 시작하세요')
            return
        done = sum(1 for it in self.items if it['completed'])
        text = f'세부항목 {len(self.categories)}개 · 항목 {len(self.items)}개'
        self.lbl_meta.setText(text + (f' · 완료 {done}개' if done else ''))

    # ── 왼쪽 목록 ──────────────────────────────────────────────────
    def _rebuild_nav(self):
        while self.nav_layout.count():
            w = self.nav_layout.takeAt(0).widget()
            if w:
                w.setParent(None)
                w.deleteLater()
        self._nav_refs = {}
        for key, name in [('ALL', '전체')] + [(c['id'], c['name']) for c in self.categories]:
            nav = _NavItem(name)
            nav.clicked.connect(lambda k=key: self._set_tab(k))
            self.nav_layout.addWidget(nav)
            self._nav_refs[key] = nav
        self.nav_layout.addStretch()
        self._update_nav()

    def _update_nav(self):
        for key, nav in self._nav_refs.items():
            if key == 'ALL':
                b = sum(it['budget'] for it in self.items)
                s = sum(it['spent'] for it in self.items)
                nav.set_values(b, s, f' · {len(self.categories)}개 세부항목')
            else:
                b, s = self._cat_totals(key)
                nav.set_values(b, s)
            nav.set_active(key == self.active_tab)

    def _set_tab(self, tab):
        if tab == self.active_tab:
            return
        self.active_tab = tab
        self._update_nav()
        self._rebuild_sections()
        self.scroll.verticalScrollBar().setValue(0)

    # ── 오른쪽 표 ──────────────────────────────────────────────────
    def _rebuild_sections(self):
        self._row_refs = {}
        self._sec_refs = {}
        while self.sections_layout.count():
            w = self.sections_layout.takeAt(0).widget()
            if w:
                w.setParent(None)
                w.deleteLater()
        self._overview_chips = {}
        target = self.categories if self.active_tab == 'ALL' else \
            [c for c in self.categories if c['id'] == self.active_tab]
        if self.active_tab == 'ALL' and self.items:
            self.sections_layout.addWidget(self._make_cost_overview())
        for cat in target:
            cat_items = [it for it in self.items if it['category_id'] == cat['id']]
            self.sections_layout.addWidget(self._make_section(cat, cat_items))
        self.sections_layout.addStretch()

    def _make_section(self, cat, cat_items):
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(1, 0, 1, 8)
        v.setSpacing(0)

        head = QWidget()
        hl = QHBoxLayout(head)
        hl.setContentsMargins(18, 14, 12, 8)
        hl.setSpacing(10)
        name = QLabel(cat['name'])
        name.setObjectName('SectionTitle')
        name.setWordWrap(True)
        meta = QLabel()
        meta.setObjectName('SectionMeta')
        meta.setTextFormat(Qt.RichText)
        chip = QLabel()
        add = _icon_btn('fa5s.plus', '이 세부항목에 항목 추가', color=C.MUTED, size=12)
        add.clicked.connect(lambda _, cid=cat['id']: self._open_add_item(cid))
        drop = _icon_btn('fa5s.trash-alt', '이 세부항목 지우기 (안의 항목과 사용 내역도 함께 지워요)',
                         color='#E11D48', size=12)
        drop.clicked.connect(lambda _, cid=cat['id']: self._delete_category(cid))
        hl.addWidget(name, 1)
        hl.addWidget(meta)
        hl.addWidget(chip)
        hl.addWidget(add)
        hl.addWidget(drop)
        v.addWidget(head)

        # 원가통계비목별 잔액 (누르면 그 비목의 줄만 보기)
        chips = {}
        if cat_items:
            chip_box = QWidget()
            flow = _FlowLayout(chip_box)
            flow.setContentsMargins(18, 0, 18, 10)
            for cost in _cost_totals(cat_items):
                b = QPushButton()
                b.setCheckable(True)
                b.setCursor(Qt.PointingHandCursor)
                b.clicked.connect(lambda _, cid=cat['id'], c=cost: self._toggle_cost_filter(cid, c))
                flow.addWidget(b)
                chips[cost] = b
            v.addWidget(chip_box)
        tables = []

        active = [it for it in cat_items if not it['completed']]
        done = [it for it in cat_items if it['completed']]
        if not cat_items:
            lbl = QLabel('아직 항목이 없어요. 오른쪽 위 + 로 추가하세요.')
            lbl.setObjectName('Muted')
            lbl.setContentsMargins(18, 4, 18, 8)
            v.addWidget(lbl)
        elif active:
            tables.append(self._make_table(active))
            v.addWidget(tables[-1])
        else:
            lbl = QLabel('모든 항목을 완료했어요.')
            lbl.setObjectName('Muted')
            lbl.setContentsMargins(18, 4, 18, 4)
            v.addWidget(lbl)
        if done:
            open_ = cat['id'] not in self._hide_done
            wrap = QWidget()
            wl = QHBoxLayout(wrap)
            wl.setContentsMargins(14, 6, 14, 0)
            toggle = QPushButton()
            toggle.setObjectName('Link')
            toggle.setCursor(Qt.PointingHandCursor)
            wl.addWidget(toggle)
            wl.addStretch()
            dtable = self._make_table(done, finished=True)
            tables.append(dtable)

            def set_open(on, cid=cat['id'], t=dtable, b=toggle, n=len(done)):
                (self._hide_done.discard if on else self._hide_done.add)(cid)
                t.setVisible(on)
                b.setText(f'{"▾" if on else "▸"}  완료 {n}건')
                if on:
                    t.sync_geometry()
            toggle.clicked.connect(lambda _, c=cat['id'], f=set_open: f(c in self._hide_done))
            set_open(open_)
            v.addWidget(wrap)
            v.addWidget(dtable)
        self._sec_refs[cat['id']] = {'meta': meta, 'chip': chip, 'chips': chips, 'tables': tables}
        self._update_section(cat['id'])
        self._apply_cost_filter(cat['id'])
        return card

    # ── 원가통계비목 ───────────────────────────────────────────────
    def _toggle_cost_filter(self, cat_id, cost):
        cur = self._cost_filter.get(cat_id)
        self._cost_filter[cat_id] = None if cur == cost else cost
        self._apply_cost_filter(cat_id)

    def _apply_cost_filter(self, cat_id):
        refs = self._sec_refs.get(cat_id)
        if not refs:
            return
        cost = self._cost_filter.get(cat_id)
        if cost not in refs['chips']:
            cost = self._cost_filter[cat_id] = None
        for c, b in refs['chips'].items():
            b.setChecked(c == cost)
        for t in refs['tables']:
            for i, rc in enumerate(t.row_costs):
                hidden = cost is not None and rc != cost
                t.setRowHidden(2 * i, hidden)
                t.setRowHidden(2 * i + 1, hidden or not t.detail_open[i])
            t.sync_geometry()

    def _make_cost_overview(self):
        """'전체'를 볼 때 맨 위: 세부항목을 가리지 않고 원가통계비목끼리 합친 잔액."""
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(18, 14, 18, 14)
        v.setSpacing(8)
        head = QHBoxLayout()
        head.setSpacing(6)
        toggle = QPushButton()
        toggle.setCursor(Qt.PointingHandCursor)
        toggle.setFlat(True)
        toggle.setStyleSheet(f'QPushButton {{ text-align: left; border: none; background: transparent; '
                             f'padding: 0; color: {C.TEXT}; font-size: 12pt; font-weight: 800; }}')
        head.addWidget(toggle)
        head.addStretch()
        v.addLayout(head)
        sub = QLabel('세부항목을 가리지 않고 같은 비목끼리 합친 금액이에요.')
        sub.setObjectName('SectionMeta')
        box = QWidget()
        flow = _FlowLayout(box)
        flow.setContentsMargins(0, 2, 0, 0)
        self._overview_chips = {}
        for cost in _cost_totals(self.items):
            lb = QLabel()
            lb.setObjectName('CostChip')
            flow.addWidget(lb)
            self._overview_chips[cost] = lb
        v.addWidget(sub)
        v.addWidget(box)

        def set_open(on):
            self._overview_open = on
            sub.setVisible(on)
            box.setVisible(on)
            toggle.setText(f'{"▾" if on else "▸"}  원가통계비목별 합계')

        toggle.clicked.connect(lambda: set_open(not self._overview_open))
        set_open(self._overview_open)
        self._update_cost_overview()
        return card

    def _update_cost_overview(self):
        totals = _cost_totals(self.items)
        for cost, lb in getattr(self, '_overview_chips', {}).items():
            try:
                b, s, n = totals.get(cost, (0, 0, 0))
                lb.setText(_cost_chip_text(cost, b, s))
                lb.setToolTip(_cost_chip_tip(cost, b, s, n))
                lb.setStyleSheet(_cost_chip_qss(b, s, 'QLabel'))
            except RuntimeError:          # 이미 지운 칩
                pass

    def _make_table(self, items, finished=False):
        t = _BudgetTable()
        t.setColumnCount(len(_COLS))
        t.setHorizontalHeaderLabels(list(_COLS))
        for col in range(len(_COLS)):
            h = t.horizontalHeaderItem(col)
            h.setTextAlignment((Qt.AlignRight if 2 <= col <= 4 else Qt.AlignLeft) | Qt.AlignVCenter)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(_sizes()['row_h'])
        t.setSelectionMode(QAbstractItemView.SingleSelection)
        t.setEditTriggers(QAbstractItemView.CurrentChanged | QAbstractItemView.DoubleClicked
                          | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        t.setFocusPolicy(Qt.ClickFocus)
        t.setShowGrid(False)
        t.setWordWrap(False)
        t.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        t.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)   # 창이 아주 좁을 때만 (잘리는 대신)
        hdr = t.horizontalHeader()
        hdr.setHighlightSections(False)
        hdr.setSectionsClickable(False)
        hdr.setMinimumSectionSize(40)
        for col in range(len(_COLS)):
            hdr.setSectionResizeMode(col, QHeaderView.Fixed)
        hdr.setStretchLastSection(True)          # 반올림으로 남는 몇 px까지 채움 (오른쪽 끝 틈 없게)
        if finished:
            hdr.hide()
        t.setRowCount(2 * len(items))                               # 항목마다 줄 + 사용 내역 숨은 줄
        t.action_buttons = [None] * (2 * len(items))
        t.row_costs = [it['cost_type'] or '일반' for it in items]     # 비목 칩으로 거를 때
        t.detail_open = [False] * len(items)
        t.blockSignals(True)
        for i, item in enumerate(items):
            self._fill_row(t, 2 * i, item)
        t.blockSignals(False)
        t.sync_geometry()
        return t

    def _fill_row(self, t, r, item):
        done = bool(item['completed'])
        rem = item['budget'] - item['spent']
        pct = _pct(item['spent'], item['budget'])
        L, R = Qt.AlignLeft | Qt.AlignVCenter, Qt.AlignRight | Qt.AlignVCenter

        def cell(text, align, color=C.TEXT, bold=False, strike=False, editable=False):
            it = QTableWidgetItem(text)
            flags = Qt.ItemIsEnabled
            if editable:
                flags |= Qt.ItemIsEditable | Qt.ItemIsSelectable
            it.setFlags(flags)
            it.setTextAlignment(align)
            f = QFont(t.font())
            f.setBold(bold)
            f.setStrikeOut(strike)
            it.setFont(f)
            it.setForeground(QColor(C.FAINT if done else color))
            return it

        t.setItem(r, 0, cell(item['cost_type'], L, C.MUTED))
        t.setItem(r, 1, cell('', L))                               # 글자는 아래의 산출내역 칸(_DescCell)이 그린다
        dcell = _DescCell(t, r, item, self._expense_ui.setdefault(item['id'], {}),
                          self._expenses_by_item.get(item['id'], []), self._delete_expense, done)
        t.setCellWidget(r, 1, dcell)
        t.setSpan(r + 1, 1, 1, 3)                                  # 숨은 줄: 산출내역 ~ 잔액 칸을 하나로
        t.setCellWidget(r + 1, 1, dcell.detail)
        t.setItem(r, 2, cell(_fmt(item['budget']), R))
        rem_it = cell(_fmt(rem), R, C.NEG if rem < 0 else C.TEXT, bold=rem < 0)
        t.setItem(r, 3, rem_it)

        wrap = QWidget()
        pl = QHBoxLayout(wrap)
        pl.setContentsMargins(6, 0, 8, 0)
        pl.setSpacing(6)
        pl.addStretch()
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setTextVisible(False)
        bar.setFixedSize(40, 6)
        lbl = QLabel()
        lbl.setFixedWidth(_sizes()['pct_lbl'])
        lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        pl.addWidget(bar)
        pl.addWidget(lbl)
        t.setItem(r, 4, cell('', R))
        t.setCellWidget(r, 4, wrap)
        self._set_pct(bar, lbl, pct, done)
        t.pct_labels.append(lbl)

        sz = _sizes()
        acts = QWidget()
        al = QHBoxLayout(acts)
        al.setContentsMargins(4, 0, 10, 0)
        al.setSpacing(6)
        al.addStretch()
        b_edit = QPushButton('수정')
        b_edit.setObjectName('RowBtn')
        b_edit.setIcon(qta.icon('fa5s.pen', color=C.MUTED))
        b_edit.setToolTip('수정 · 이번에 쓴 금액 넣기 · 삭제')
        b_edit.clicked.connect(lambda _, iid=item['id']: self._open_edit_item(iid))
        b_done = QPushButton('완료 취소' if done else '완료')
        b_done.setObjectName('RowBtn' if done else 'RowDone')
        b_done.setIcon(qta.icon('fa5s.undo' if done else 'fa5s.check', color=C.MUTED if done else '#047857'))
        b_done.setToolTip('완료를 취소합니다' if done else '이 항목을 완료로 표시합니다')
        b_done.clicked.connect(lambda _, iid=item['id']: self._toggle_complete(iid))
        for b in (b_edit, b_done):                       # 수정 왼쪽 · 완료 오른쪽
            b.setIconSize(QSize(11, 11))
            b.setMinimumWidth(sz['btn_wide'] if b.text() == '완료 취소' else sz['btn_w'])
            b.setFixedHeight(sz['btn_h'])
            b.setCursor(Qt.PointingHandCursor)
            al.addWidget(b)
        t.setItem(r, 5, cell('', R))
        t.setCellWidget(r, 5, acts)
        t.action_buttons[r] = (b_edit, b_done)
        t.setRowHeight(r, sz['row_h'])
        self._row_refs[item['id']] = {'rem': rem_it, 'bar': bar, 'pct': lbl, 'desc': dcell}
        dcell._fit_row()                                   # 펼쳐 둔 줄은 다시 펼친 높이로

    @staticmethod
    def _set_pct(bar, lbl, pct, done=False):
        lvl = _level(pct, done)
        bar.setValue(int(round(min(pct, 100))))
        bar.setStyleSheet(_bar_qss(_LEVEL_BAR[lvl]))
        lbl.setText(f'{pct:.1f}%')
        lbl.setStyleSheet(f'color: {_LEVEL_TEXT[lvl]}; font-weight: {400 if done else 700};')

    def _update_row(self, item_id):
        refs = self._row_refs.get(item_id)
        item = self._items_by_id.get(item_id)
        if not refs or not item:
            return
        rem = item['budget'] - item['spent']
        refs['rem'].setText(_fmt(rem))
        refs['rem'].setForeground(QColor(C.NEG if rem < 0 else C.TEXT))
        f = refs['rem'].font()
        f.setBold(rem < 0)
        refs['rem'].setFont(f)
        self._set_pct(refs['bar'], refs['pct'], _pct(item['spent'], item['budget']))
        refs['desc'].set_expenses(self._expenses_by_item.get(item_id, []))

    # ── 합계 ───────────────────────────────────────────────────────
    def _cat_totals(self, cat_id):
        budget = spent = 0
        for it in self.items:
            if it['category_id'] == cat_id:
                budget += it['budget']
                spent += it['spent']
        return budget, spent

    def _update_section(self, cat_id):
        refs = self._sec_refs.get(cat_id)
        if not refs:
            return
        budget, spent = self._cat_totals(cat_id)
        rem = budget - spent
        pct = _pct(spent, budget)
        lvl = _level(pct)
        rem_html = f'<span style="color:{C.NEG}"><b>{_fmt(rem)}원</b></span>' if rem < 0 else f'{_fmt(rem)}원'
        refs['meta'].setText(f'예산 {_fmt(budget)}원 · 집행 {_fmt(spent)}원 · 잔액 {rem_html}')
        refs['chip'].setText(f'{pct:.1f}%')
        refs['chip'].setStyleSheet(f'background: {_LEVEL_SOFT[lvl]}; color: {_LEVEL_TEXT[lvl]}; '
                                   f'border-radius: 10px; padding: 3px 10px; font-weight: 700;')
        totals = _cost_totals([it for it in self.items if it['category_id'] == cat_id])
        for cost, b in refs.get('chips', {}).items():
            cb, cs, n = totals.get(cost, (0, 0, 0))
            b.setText(_cost_chip_text(cost, cb, cs))
            b.setToolTip(_cost_chip_tip(cost, cb, cs, n) + '\n\n누르면 이 비목의 줄만 보여요')
            b.setStyleSheet(_cost_chip_qss(cb, cs))

    def _update_summary(self):
        budget = sum(it['budget'] for it in self.items)
        spent = sum(it['spent'] for it in self.items)
        rem = budget - spent
        pct = _pct(spent, budget)
        lvl = _level(pct)
        self.stat['budget'].setText(f'{_fmt(budget)}원')
        self.stat['spent'].setText(f'{_fmt(spent)}원')
        self.stat['remaining'].setText(f'{_fmt(rem)}원')
        self.stat['remaining'].setStyleSheet(f'color: {C.NEG if rem < 0 else C.TEXT};')
        self.lbl_pct.setText(f'{pct:.1f}%')
        self.lbl_pct.setStyleSheet(f'color: {_LEVEL_TEXT[lvl]}; font-size: 12pt; font-weight: 800;')
        self.bar_total.setValue(int(round(min(pct, 100))))
        self.bar_total.setStyleSheet(_bar_qss(_LEVEL_BAR[lvl]))
        self._update_meta()

    # ── 동작: 항목/세부항목 ────────────────────────────────────────
    def _toggle_complete(self, item_id):
        toggle_budget_item_completed(item_id)
        self.refresh()

    def _open_add_item(self, cat_id=None):
        if not self.categories:
            QMessageBox.information(self, '세부항목 필요',
                                    '먼저 세부항목을 만들거나, 엑셀·캡처를 가져오세요.')
            self._open_category_manager()
            return
        default_cat = cat_id or (self.active_tab if self.active_tab != 'ALL' else self.categories[0]['id'])
        dlg = _ItemDialog(self.categories, None, default_cat, self)
        if dlg.exec_() == QDialog.Accepted and dlg.result_data:
            d = dlg.result_data
            add_budget_item(d['category_id'], d['cost_type'], d['description'], d['budget'], d['spent'])
            self.refresh()

    def _open_edit_item(self, item_id):
        item = self._items_by_id.get(item_id)
        if not item:
            return
        dlg = _ItemDialog(self.categories, item, item['category_id'], self)
        if dlg.exec_() != QDialog.Accepted:
            return
        if dlg.deleted:                                  # 수정 창 오른쪽 위 [삭제]
            delete_budget_item(item_id)
            self.refresh()
        elif dlg.result_data:
            d = dlg.result_data
            save_budget_item_edit(item_id, d['category_id'], d['cost_type'],
                                  d['description'], d['budget'], d['spent'], d.get('expense'))
            self.refresh()

    def _delete_expense(self, expense_id):
        """사용 내역 한 건 지우기 — 집행액도 그만큼 줄어든다. 펼쳐 둔 줄은 그대로 펼쳐 둔다."""
        ex = next((e for es in self._expenses_by_item.values() for e in es if e['id'] == expense_id), None)
        if not ex:
            return
        name = ex['purpose'] or '쓴 곳 미입력'
        ans = QMessageBox.question(self, '사용 내역 삭제',
                                   f"'{name}' {_fmt(ex['amount'])}원 기록을 지울까요?\n집행액도 그만큼 줄어듭니다.",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return
        delete_budget_expense(expense_id)
        self.refresh()

    def _delete_category(self, cat_id):
        """세부항목 지우기 — 안의 항목과 사용 내역도 함께 지운다. 항목·금액을 보여 주고 한 번 더 확인."""
        cat = next((c for c in self.categories if c['id'] == cat_id), None)
        if not cat:
            return
        inside = [it for it in self.items if it['category_id'] == cat_id]
        if inside:
            spent = sum(it['spent'] for it in inside)
            detail = (f"안의 항목 {len(inside)}개와 집행액 {_fmt(spent)}원, 사용 내역이 함께 지워집니다.\n"
                      "되돌릴 수 없습니다.")
        else:
            detail = '안에 항목은 없습니다.'
        ans = QMessageBox.question(self, '세부항목 지우기', f"'{cat['name']}' 세부항목을 지울까요?\n{detail}",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return
        delete_budget_category(cat_id)
        self._cost_filter.pop(cat_id, None)
        if self.active_tab == cat_id:
            self.active_tab = 'ALL'
        self.refresh()

    def _open_category_manager(self):
        _CategoryManagerDialog(self, self.refresh).exec_()

    def _reset_all(self):
        if not self.categories and not self.items:
            return
        ans = QMessageBox.question(self, '전체 지우기',
                                   '모든 세부항목과 예산 항목을 지웁니다. 계속할까요?',
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans == QMessageBox.Yes:
            clear_budget_data()
            self.active_tab = 'ALL'
            self.refresh(keep_scroll=False)

    # ── 가져오기: 엑셀 파일 · 엑셀 표 붙여넣기 · 캡처 이미지 ─────────
    def _choose_excel(self, parent=None):
        parent = parent or self
        path, _ = QFileDialog.getOpenFileName(parent, '예산 엑셀 파일 선택', '', '엑셀 파일 (*.xls *.xlsx)')
        if path:
            self._import_excel(path, parent=parent)

    def _import_excel(self, path, parent=None):
        try:
            parsed = parse_budget_excel(path)
        except Exception as e:
            QMessageBox.critical(parent or self, '가져오기 실패', f'엑셀 파일을 읽을 수 없습니다.\n\n{e}')
            return
        self._import_parsed(parsed, f'엑셀 파일 · {os.path.basename(path)}', parent=parent)

    @staticmethod
    def _clipboard_kind():
        md = QApplication.clipboard().mimeData()
        if md is None:
            return None
        if md.hasImage():
            return 'image'
        if md.hasUrls() and any(u.toLocalFile().lower().endswith(_IMAGE_EXTS + _EXCEL_EXTS) for u in md.urls()):
            return 'file'
        if md.hasText() and md.text().strip():
            return 'text'
        return None

    def _paste(self):
        cb = QApplication.clipboard()
        md = cb.mimeData()
        kind = self._clipboard_kind()
        if kind == 'image':
            self._import_image(cb.image())
            return
        if kind == 'file':
            for u in md.urls():
                path = u.toLocalFile()
                if path.lower().endswith(_EXCEL_EXTS):
                    self._import_excel(path)
                    return
                if path.lower().endswith(_IMAGE_EXTS):
                    self._import_image(QImage(path))
                    return
        if kind != 'text':
            QMessageBox.information(self, '붙여넣기',
                                    '붙여넣을 내용이 없습니다.\n'
                                    '예산 화면을 캡처하거나, 엑셀에서 셀 범위를 복사한 뒤 다시 시도해 주세요.')
            return
        try:
            parsed = parse_budget_clipboard(md.text())
        except Exception as e:
            QMessageBox.warning(self, '붙여넣기 실패', str(e))
            return
        self._import_parsed(parsed, '엑셀에서 복사한 표')

    def _import_image(self, img):
        if img is None or img.isNull():
            QMessageBox.warning(self, '캡처 정리', '이미지를 읽지 못했습니다.')
            return
        if self._img_worker is not None and self._img_worker.isRunning():
            QMessageBox.information(self, '캡처 정리', '앞의 캡처를 아직 읽는 중이에요. 잠시 뒤 다시 해 주세요.')
            return
        from ai_client import _pixmap_to_base64
        from budget_image import BudgetImageWorker
        b64 = _pixmap_to_base64(QPixmap.fromImage(img))
        ocr_img = img.scaled(img.width() * 2, img.height() * 2, Qt.KeepAspectRatio, Qt.SmoothTransformation) \
            if img.width() < 1800 else img          # 작은 글자는 키워서 읽어야 잘 읽힌다
        buf = QBuffer()
        buf.open(QIODevice.WriteOnly)
        ocr_img.save(buf, 'PNG')
        worker = BudgetImageWorker(b64, bytes(buf.data()))
        prog = QProgressDialog('캡처한 표를 읽는 중…', '취소', 0, 0, self)
        prog.setWindowTitle('캡처 정리')
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(0)
        prog.setAutoClose(False)
        prog.setAutoReset(False)
        prog.setStyleSheet(STYLE)

        def finish():
            prog.reset()
            prog.close()

        def on_done(cats, method):
            finish()
            self._import_parsed(cats, f'캡처 이미지 · {_IMAGE_METHOD.get(method, method)}')

        def on_fail(msg):
            finish()
            QMessageBox.warning(self, '캡처 정리 실패', msg)
        prog.canceled.connect(lambda: setattr(worker, 'cancelled', True))
        worker.progress.connect(prog.setLabelText)
        worker.done.connect(on_done)
        worker.failed.connect(on_fail)
        self._img_worker = worker
        worker.start()
        prog.show()

    def _import_parsed(self, parsed, source_text, parent=None):
        parent = parent or self
        if not parsed or not any(c['items'] for c in parsed):
            QMessageBox.warning(parent, '가져오기 실패',
                                '예산 데이터를 찾지 못했습니다.\n'
                                "사업관리카드(예산) 형식('예산현액', '산출내역' 열)인지 확인해주세요.")
            return
        existing = {c['name'] for c in self.categories}
        dlg = _ImportPreviewDialog(source_text, parsed, bool(self.categories or self.items), existing, parent)
        if dlg.exec_() != QDialog.Accepted:
            return
        self._apply_import(dlg.categories, dlg.mode)
        n_cat = len(dlg.categories)
        n_item = sum(len(c['items']) for c in dlg.categories)
        action = {'replace': '교체했어요', 'update': '반영했어요', 'append': '추가했어요'}[dlg.mode]
        QMessageBox.information(parent, '가져오기 완료', f'세부항목 {n_cat}개, 예산 항목 {n_item}개를 {action}.')

    def _apply_import(self, categories, mode):
        """mode: append(모두 추가) · update(같은 이름의 세부항목은 새 내용으로 바꿈) · replace(전부 지우고 교체)."""
        if mode == 'replace':
            clear_budget_data()
        existing = {c['name']: c['id'] for c in get_budget_categories()}
        for cat in categories:
            cid = existing.get(cat['name'])
            if cid is not None and mode == 'update':
                for it in get_budget_items(cid):
                    delete_budget_item(it['id'])
            if cid is None:
                cid = add_budget_category(cat['name'])
                if cid is None:
                    existing = {c['name']: c['id'] for c in get_budget_categories()}
                    cid = existing.get(cat['name'])
                if cid is None:
                    continue
                existing[cat['name']] = cid
            for it in cat['items']:
                add_budget_item(cid, it['cost_type'], it['description'], it['budget'], it['spent'])
        self.active_tab = 'ALL'
        self.refresh(keep_scroll=False)

    # ── 끌어다 놓기 · Ctrl+V ───────────────────────────────────────
    @staticmethod
    def _accepts(md):
        if md.hasImage():
            return True
        return md.hasUrls() and any(u.toLocalFile().lower().endswith(_EXCEL_EXTS + _IMAGE_EXTS) for u in md.urls())

    def dragEnterEvent(self, event):
        if self._accepts(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        self.dragEnterEvent(event)

    def dropEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            for url in md.urls():
                path = url.toLocalFile()
                low = path.lower()
                if low.endswith(_EXCEL_EXTS):
                    event.acceptProposedAction()
                    self._import_excel(path)
                    return
                if low.endswith(_IMAGE_EXTS):
                    event.acceptProposedAction()
                    self._import_image(QImage(path))
                    return
        if md.hasImage():
            event.acceptProposedAction()
            self._import_image(QImage(md.imageData()))
            return
        event.ignore()

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Paste) and self._can_paste_import():
            self._paste()
            return
        super().keyPressEvent(event)

    def _can_paste_import(self):
        w = QApplication.focusWidget()
        return not isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit, QComboBox))

    # ── 창 위치/크기 저장·복원 ────────────────────────────────────
    def _restore_window_state(self):
        try:
            with open(_STATE_FILE, 'r', encoding='utf-8') as f:
                s = json.load(f)
            from ui.common import fit_to_screen
            # 예전 창 크기가 새 화면(왼쪽 목록·버튼)에 좁으면 넓혀 준다
            self.setGeometry(*fit_to_screen(s['x'], s['y'], max(s['width'], px(1240)), s['height'],
                                            min_w=640, min_h=420))
        except Exception:
            from ui.common import default_geometry
            self.setGeometry(*default_geometry(px(1240), px(780), frac_w=0.92))

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh(keep_scroll=False)

    def closeEvent(self, event):
        try:
            os.makedirs(os.path.dirname(_STATE_FILE), exist_ok=True)
            g = self.geometry()
            with open(_STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump({'x': g.x(), 'y': g.y(), 'width': g.width(), 'height': g.height()}, f)
        except Exception:
            pass
        super().closeEvent(event)


# ── 가져오기 확인 (세부항목별로 보고 고친 뒤 가져오기) ───────────────
class _ImportPreviewDialog(QDialog):
    """가져올 내용을 세부항목별로 보여 준다. 칸을 두 번 눌러 고칠 수 있고, 확인이 필요한 줄은 노란색.

    결과: self.categories (고친 내용), self.mode ('append' | 'update' | 'replace')
    """
    _COLS = ('원가통계비목', '산출내역', '예산현액', '집행액', '잔액')

    def __init__(self, source_text, categories, has_existing, existing_names=(), parent=None):
        super().__init__(parent)
        self.setWindowTitle('가져오기 확인')
        self.setStyleSheet(STYLE)
        self.setMinimumSize(860, 620)
        self.mode = None
        self.categories = []
        self._src = categories
        self._existing = set(existing_names)

        v = QVBoxLayout(self)
        v.setContentsMargins(22, 18, 22, 18)
        v.setSpacing(8)
        title = QLabel('가져올 내용을 확인하세요')
        title.setObjectName('SectionTitle')
        src = QLabel(source_text)
        src.setObjectName('Meta')
        self.lbl_sum = QLabel()
        self.lbl_sum.setTextFormat(Qt.RichText)
        hint = QLabel('칸을 두 번 눌러 고칠 수 있어요 (세부항목 이름도). 노란 줄은 숫자를 다시 확인해 주세요.')
        hint.setObjectName('Muted')
        for w in (title, src, self.lbl_sum, hint):
            v.addWidget(w)

        t = self.table = QTableWidget()
        t.setColumnCount(len(self._COLS))
        t.setHorizontalHeaderLabels(list(self._COLS))
        t.verticalHeader().setVisible(False)
        t.setShowGrid(False)
        t.setSelectionMode(QAbstractItemView.SingleSelection)
        t.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed
                          | QAbstractItemView.SelectedClicked)
        hdr = t.horizontalHeader()
        hdr.setHighlightSections(False)
        hdr.setSectionsClickable(False)
        for col in range(len(self._COLS)):
            hdr.setSectionResizeMode(col, QHeaderView.Fixed)
            t.horizontalHeaderItem(col).setTextAlignment(
                (Qt.AlignRight if col >= 2 else Qt.AlignLeft) | Qt.AlignVCenter)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        for col, wd in ((0, 150), (2, 112), (3, 112), (4, 112)):
            t.setColumnWidth(col, wd)
        t.setStyleSheet(f'QTableWidget {{ border: 1px solid {C.BORDER}; border-radius: 10px; }}')
        self._map = []
        rows = []
        for ci, cat in enumerate(categories):
            rows.append(('cat', ci, None))
            for ii in range(len(cat['items'])):
                rows.append(('item', ci, ii))
        t.setRowCount(len(rows))
        R, L = Qt.AlignRight | Qt.AlignVCenter, Qt.AlignLeft | Qt.AlignVCenter
        for r, (kind, ci, ii) in enumerate(rows):
            self._map.append((kind, ci, ii))
            cat = categories[ci]
            if kind == 'cat':
                it = QTableWidgetItem(cat['name'])
                f = QFont(t.font())
                f.setBold(True)
                it.setFont(f)
                it.setForeground(QColor(C.ACCENT_HI))
                it.setBackground(QColor(C.ACCENT_SOFT))
                it.setToolTip('이미 있는 세부항목이에요' if cat['name'] in self._existing else '새 세부항목')
                t.setItem(r, 0, it)
                t.setSpan(r, 0, 1, len(self._COLS))
                t.setRowHeight(r, 34)
                continue
            item = cat['items'][ii]
            warn = item.get('warn') or []
            vals = (item['cost_type'], item['description'], _fmt(item['budget']), _fmt(item['spent']),
                    _fmt(item['budget'] - item['spent']))
            for col, text in enumerate(vals):
                it = QTableWidgetItem(text)
                it.setTextAlignment(R if col >= 2 else L)
                if col == 4:
                    it.setFlags(Qt.ItemIsEnabled)
                    it.setForeground(QColor(C.NEG if item['budget'] - item['spent'] < 0 else C.MUTED))
                if warn:
                    it.setBackground(QColor(C.WARN_BG))
                    it.setToolTip('\n'.join(warn))
                t.setItem(r, col, it)
            t.setRowHeight(r, 30)
        t.itemChanged.connect(self._on_changed)
        v.addWidget(t, 1)

        self._radios = {}
        box = QVBoxLayout()
        box.setSpacing(4)
        if has_existing:
            group = QButtonGroup(self)
            for key, text in (('update', '같은 이름의 세부항목은 새 내용으로 바꾸고, 나머지는 추가 (추천)'),
                              ('append', '지금 있는 내용은 그대로 두고 모두 추가'),
                              ('replace', '지금 있는 예산을 모두 지우고 이것으로 교체')):
                rb = QRadioButton(text)
                group.addButton(rb)
                box.addWidget(rb)
                self._radios[key] = rb
            self._radios['update'].setChecked(True)
        v.addLayout(box)

        btns = QHBoxLayout()
        btns.addStretch()
        b_ok = QPushButton('가져오기')
        b_ok.setObjectName('Primary')
        b_ok.setDefault(True)
        b_ok.clicked.connect(self._accept)
        b_cancel = QPushButton('취소')
        b_cancel.setAutoDefault(False)
        b_cancel.clicked.connect(self.reject)
        btns.addWidget(b_ok)                              # 가져오기 왼쪽 · 취소 오른쪽
        btns.addWidget(b_cancel)
        v.addLayout(btns)
        self.btn_import, self.btn_cancel = b_ok, b_cancel
        self._update_sum()

    def _read(self):
        """표 → 세부항목 목록 (고친 내용 반영)."""
        cats = []
        for r, (kind, ci, ii) in enumerate(self._map):
            if kind == 'cat':
                cats.append({'name': (self.table.item(r, 0).text() or '').strip() or '미분류', 'items': []})
                continue
            txt = [(self.table.item(r, c).text() if self.table.item(r, c) else '').strip() for c in range(4)]
            if not txt[1] and not _parse_amount(txt[2]):
                continue                     # 내용도 금액도 지운 줄은 뺀다
            cats[-1]['items'].append({'cost_type': txt[0] or '일반', 'description': txt[1] or '(내용 없음)',
                                      'budget': _parse_amount(txt[2]), 'spent': _parse_amount(txt[3])})
        return [c for c in cats if c['items']]

    def _on_changed(self, it):
        r = it.row()
        kind = self._map[r][0] if r < len(self._map) else None
        if kind == 'item' and it.column() in (2, 3):
            t = self.table
            t.blockSignals(True)
            b, s = _parse_amount(t.item(r, 2).text()), _parse_amount(t.item(r, 3).text())
            t.item(r, 2).setText(_fmt(b))
            t.item(r, 3).setText(_fmt(s))
            t.item(r, 4).setText(_fmt(b - s))
            t.item(r, 4).setForeground(QColor(C.NEG if b - s < 0 else C.MUTED))
            t.blockSignals(False)
        self._update_sum()

    def _update_sum(self):
        cats = self._read()
        b = sum(i['budget'] for c in cats for i in c['items'])
        s = sum(i['spent'] for c in cats for i in c['items'])
        n = sum(len(c['items']) for c in cats)
        warn = sum(1 for c in self._src for i in c['items'] if i.get('warn'))
        same = sum(1 for c in cats if c['name'] in self._existing)
        parts = [f'세부항목 <b>{len(cats)}</b>개', f'항목 <b>{n}</b>개', f'예산현액 <b>{_fmt(b)}원</b>',
                 f'집행액 <b>{_fmt(s)}원</b>', f'잔액 <b>{_fmt(b - s)}원</b>']
        if same:
            parts.append(f'이미 있는 세부항목 <b>{same}</b>개')
        if warn:
            parts.append(f'<span style="color:#B45309">확인 필요 <b>{warn}</b>줄</span>')
        self.lbl_sum.setText(' · '.join(parts))

    def _accept(self):
        cats = self._read()
        if not cats:
            QMessageBox.warning(self, '가져오기', '가져올 항목이 없습니다.')
            return
        self.categories = cats
        self.mode = next((k for k, rb in self._radios.items() if rb.isChecked()), 'append')
        self.accept()


# ── 예산 항목 추가/수정 다이얼로그 ────────────────────────────────
class _DescCell(QWidget):
    """산출내역 칸: 이름 한 줄 + 사용 내역 요약. 누르면 명목별 사용 내역이 펼쳐진다.
    펼침·더 보기 상태는 state(창을 여는 동안만 유지)에 둔다."""
    SHOW = 5                                       # 펼쳤을 때 먼저 보여 줄 건수

    def __init__(self, table, row, item, state, expenses, on_delete, done=False):
        super().__init__()
        self._table, self._row, self._item = table, row, item
        self._state, self._on_delete, self._done = state, on_delete, done
        self.setStyleSheet('background: transparent;')
        v = QVBoxLayout(self)
        v.setContentsMargins(6, 0, 6, 0)
        v.setSpacing(2)

        head = QVBoxLayout()
        head.setSpacing(0)
        self.btn_toggle = QPushButton()
        self.btn_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_toggle.setFlat(True)
        self.btn_toggle.setStyleSheet(
            f'QPushButton {{ text-align: left; border: none; background: transparent; padding: 0; '
            f'color: {C.FAINT if done else C.TEXT}; font-weight: 700; }} '
            f'QPushButton:hover {{ color: {C.ACCENT}; }}')
        self.btn_toggle.clicked.connect(self._toggle)
        self.btn_toggle.setToolTip(item['description'])
        name_font = self.btn_toggle.font()
        name_font.setStrikeOut(done)
        self.btn_toggle.setFont(name_font)
        self.lbl_sum = QLabel()
        self.lbl_sum.setStyleSheet(f'color: {C.MUTED}; font-size: 10pt; background: transparent;')
        head.addWidget(self.btn_toggle)
        head.addWidget(self.lbl_sum)
        v.addLayout(head)

        self.detail = QWidget()                    # 항목 줄 아래의 숨은 줄에 들어간다 (_fill_row)
        self.detail.setStyleSheet('background: transparent;')
        self._dv = QVBoxLayout(self.detail)
        self._dv.setContentsMargins(px(14), px(2), px(6), px(6))
        self._dv.setSpacing(px(3))
        self.set_expenses(expenses)

    def _open(self):
        return bool(self._state.get('open'))

    def set_expenses(self, expenses):
        """사용 내역이 바뀌면 요약과 펼친 내용을 새로 그린다."""
        self._expenses = list(expenses)
        used = sum(e['amount'] for e in self._expenses)
        unassigned = int(self._item['spent']) - used
        if unassigned < 0:
            self.lbl_sum.setText('⚠ 기록 초과')
            self.lbl_sum.setStyleSheet(f'color: {C.NEG}; font-size: 10pt; background: transparent;')
        else:
            self.lbl_sum.setText('')
        self.btn_toggle.setText(self._item['description'])
        self.btn_toggle.setIcon(qta.icon('fa5s.chevron-down' if self._open() else 'fa5s.chevron-right', color=C.MUTED))
        self.btn_toggle.setIconSize(QSize(9, 9))
        self._rebuild_detail(unassigned)
        self.detail.setVisible(self._open())
        self._fit_row()

    @staticmethod
    def _clear_layout(layout):
        while layout.count():
            it = layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
            elif it.layout() is not None:
                _DescCell._clear_layout(it.layout())

    def _rebuild_detail(self, unassigned):
        self._clear_layout(self._dv)
        if unassigned > 0:
            self._dv.addLayout(self._entry_row('', '명목 없음', unassigned))
        more = bool(self._state.get('more'))
        shown = self._expenses if more else self._expenses[:self.SHOW]
        for ex in shown:
            self._dv.addLayout(self._entry_row(ex['spent_on'] or '-', ex['purpose'], ex['amount'], ex['id']))
        if len(self._expenses) > self.SHOW:
            left = len(self._expenses) - self.SHOW
            btn = QPushButton('접기' if more else f'더 보기 ({left}건 더)')
            btn.setFlat(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f'QPushButton {{ text-align: left; border: none; background: transparent; '
                              f'padding: 0; color: {C.ACCENT_HI}; font-size: 10pt; }}')
            btn.clicked.connect(self._toggle_more)
            self._dv.addWidget(btn)
        if unassigned < 0:
            warn = QLabel('⚠ 기록 합계가 집행액보다 큽니다. 집행액을 직접 낮췄다면 기록을 확인하세요.')
            warn.setWordWrap(True)
            warn.setStyleSheet(f'color: {C.NEG}; font-size: 10pt; background: transparent;')
            self._dv.addWidget(warn)

    def _entry_row(self, date, purpose, amount, eid=None):
        row = QHBoxLayout()
        row.setSpacing(8)
        d = QLabel(date)
        d.setFixedWidth(88)
        d.setStyleSheet(f'color: {C.MUTED}; font-size: 10pt; background: transparent;')
        p = QLabel(purpose or '(쓴 곳 미입력)')
        p.setStyleSheet(f'color: {C.TEXT if purpose or eid is None else C.FAINT}; font-size: 10pt; background: transparent;')
        a = QLabel(f"{_fmt(amount)}원")
        a.setStyleSheet(f'color: {C.TEXT}; font-size: 10pt; font-weight: 700; background: transparent;')
        a.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        row.addWidget(d)
        row.addWidget(p, 1)
        row.addWidget(a)
        if eid is None:
            spacer = QWidget()
            spacer.setFixedSize(20, 20)
            row.addWidget(spacer)
            return row
        x = QPushButton('✕')
        x.setFlat(True)
        x.setFixedSize(20, 20)
        x.setCursor(Qt.PointingHandCursor)
        x.setToolTip('이 사용 내역을 지웁니다 (집행액도 그만큼 줄어요)')
        x.setStyleSheet(f'QPushButton {{ border: none; background: transparent; color: {C.FAINT}; }} '
                        f'QPushButton:hover {{ color: {C.NEG}; }}')
        x.clicked.connect(lambda _, eid=eid: self._on_delete(eid))
        row.addWidget(x)
        return row

    def _toggle(self):
        self._state['open'] = not self._open()
        self.set_expenses(self._expenses)

    def _toggle_more(self):
        self._state['more'] = not self._state.get('more')
        self._rebuild_detail(int(self._item['spent']) - sum(e['amount'] for e in self._expenses))
        self._fit_row()

    def _fit_row(self):
        """펼치면 항목 줄 아래의 숨은 줄이 나타나 사용 내역을 보여 주고, 접으면 숨긴다.
        숨은 줄은 항목을 비목 칩으로 거르면 함께 숨는다."""
        t, r = self._table, self._row
        open_ = self._open()
        t.detail_open[r // 2] = open_
        hidden = (not open_) or t.isRowHidden(r)
        t.setRowHidden(r + 1, hidden)
        if not hidden:
            span = sum(t.columnWidth(c) for c in range(1, 5))          # 산출내역 ~ 집행률 칸 폭
            h = self.detail.heightForWidth(span)
            if h < 0:
                h = self.detail.sizeHint().height()
            t.setRowHeight(r + 1, h + px(8))
        t.sync_geometry()


class _ItemDialog(QDialog):
    def __init__(self, categories, item, default_cat_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle('예산 항목 수정' if item else '새 예산 항목 추가')
        self.setStyleSheet(STYLE)
        self.setMinimumWidth(420)
        self.result_data = None
        self.deleted = False
        self._item = item

        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        head = QHBoxLayout()
        title = QLabel('예산 항목 수정' if item else '새 예산 항목')
        title.setObjectName('SectionTitle')
        head.addWidget(title)
        head.addStretch()
        if item:                                          # 삭제는 표가 아니라 여기서 (실수로 누르지 않게)
            self.btn_delete = QPushButton('삭제')
            self.btn_delete.setObjectName('Danger')
            self.btn_delete.setIcon(qta.icon('fa5s.trash-alt', color='#E11D48'))
            self.btn_delete.setIconSize(QSize(11, 11))
            self.btn_delete.setCursor(Qt.PointingHandCursor)
            self.btn_delete.setAutoDefault(False)         # Enter로 삭제되는 일이 없게
            self.btn_delete.setToolTip('이 예산 항목을 지웁니다')
            self.btn_delete.clicked.connect(self._delete)
            head.addWidget(self.btn_delete)
        v.addLayout(head)
        v.addSpacing(4)

        v.addWidget(self._label('세부항목'))
        self.combo_cat = QComboBox()
        for cat in categories:
            self.combo_cat.addItem(cat['name'], cat['id'])
        idx = self.combo_cat.findData(default_cat_id)
        if idx >= 0:
            self.combo_cat.setCurrentIndex(idx)
        v.addWidget(self.combo_cat)

        v.addWidget(self._label('원가통계비목'))
        self.edit_cost = QLineEdit()
        self.edit_cost.setPlaceholderText('목록에서 고르거나 직접 입력 (예: 교육운영비)')
        from budget_types import known as known_cost_types
        self._cost_completer = QCompleter(known_cost_types(), self)
        self._cost_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._cost_completer.setFilterMode(Qt.MatchContains)
        self._cost_completer.popup().setStyleSheet(STYLE)
        self.edit_cost.setCompleter(self._cost_completer)
        show_all = self.edit_cost.addAction(qta.icon('fa5s.chevron-down', color=C.FAINT), QLineEdit.TrailingPosition)
        show_all.setToolTip('원가통계비목 목록 보기')
        show_all.triggered.connect(self._show_cost_list)
        v.addWidget(self.edit_cost)

        v.addWidget(self._label('산출내역 (내용)'))
        self.edit_desc = QLineEdit()
        self.edit_desc.setPlaceholderText('예: 교과융합캠프 간식비')
        v.addWidget(self.edit_desc)

        # 이번에 쓴 금액 (수정할 때만): 적으면 집행액에 더해지고 잔액이 준다 — 따로 계산할 필요 없게
        self.edit_spend_now = None
        self.edit_purpose = None
        self.edit_spend_date = None
        if item:
            lbl_now = QLabel('이번에 쓴 금액')
            lbl_now.setStyleSheet(f'font-size: 10pt; font-weight: 700; color: {C.ACCENT};')
            v.addWidget(lbl_now)
            self.edit_spend_now = QLineEdit()
            self.edit_spend_now.setAlignment(Qt.AlignRight)
            self.edit_spend_now.setPlaceholderText('0')
            self.edit_spend_now.setStyleSheet(
                f'QLineEdit {{ color: {C.ACCENT}; font-weight: 700; border: 1px solid {C.ACCENT_LINE}; }}'
                f'QLineEdit:focus {{ border: 1px solid {C.ACCENT}; }}')
            _attach_amount_format(self.edit_spend_now)
            v.addWidget(self.edit_spend_now)
            from datetime import date
            where = QHBoxLayout()
            where.setSpacing(8)
            self.edit_purpose = QLineEdit()
            self.edit_purpose.setPlaceholderText('쓴 곳 (예: 체육대회 협의회비)')
            self.edit_spend_date = QLineEdit(date.today().isoformat())
            self.edit_spend_date.setFixedWidth(px(120))
            self.edit_spend_date.setToolTip('사용한 날짜 (예: 2026-11-10)')
            where.addWidget(self.edit_purpose, 1)
            where.addWidget(self.edit_spend_date)
            v.addLayout(where)
            self.lbl_spend_preview = QLabel()
            self.lbl_spend_preview.setTextFormat(Qt.RichText)
            self.lbl_spend_preview.setStyleSheet(f'color: {C.MUTED}; font-size: 10pt;')
            v.addWidget(self.lbl_spend_preview)

        amount_row = QHBoxLayout()
        amount_row.setSpacing(12)
        budget_col = QVBoxLayout()
        budget_col.setSpacing(4)
        budget_col.addWidget(self._label('예산현액 (원)'))
        self.edit_budget = QLineEdit('0')
        self.edit_budget.setAlignment(Qt.AlignRight)
        _attach_amount_format(self.edit_budget)
        budget_col.addWidget(self.edit_budget)
        spent_col = QVBoxLayout()
        spent_col.setSpacing(4)
        spent_col.addWidget(self._label('집행액 (지출액)'))
        self.edit_spent = QLineEdit('0')
        self.edit_spent.setAlignment(Qt.AlignRight)
        _attach_amount_format(self.edit_spent)
        spent_col.addWidget(self.edit_spent)
        amount_row.addLayout(budget_col)
        amount_row.addLayout(spent_col)
        v.addLayout(amount_row)

        btns = QHBoxLayout()
        btns.addStretch()
        self.btn_save = QPushButton('저장')
        self.btn_save.setObjectName('Primary')
        self.btn_save.setDefault(True)                    # 입력 칸에서 Enter → 저장
        self.btn_save.clicked.connect(self._save)
        self.btn_cancel = QPushButton('취소')
        self.btn_cancel.setAutoDefault(False)             # Enter가 취소를 누르는 일이 없게 (취소는 Esc나 직접 클릭)
        self.btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self.btn_save)                     # 저장 왼쪽 · 취소 오른쪽
        btns.addWidget(self.btn_cancel)
        v.addLayout(btns)

        if item:
            self.edit_cost.setText(item['cost_type'])
            self.edit_desc.setText(item['description'])
            self.edit_budget.setText(f'{item["budget"]:,}')
            self.edit_spent.setText(f'{item["spent"]:,}')
            for e in (self.edit_spend_now, self.edit_budget, self.edit_spent):
                e.textChanged.connect(self._update_spend_preview)
            self._update_spend_preview()
            self.edit_spend_now.setFocus()           # 수정 창을 여는 가장 흔한 이유
        else:
            self.edit_desc.setFocus()

    def _final_spent(self):
        """저장할 집행액 = 집행액 칸 + 이번에 쓴 금액."""
        add = _parse_amount(self.edit_spend_now.text()) if self.edit_spend_now else 0
        return _parse_amount(self.edit_spent.text()) + add

    def _update_spend_preview(self):
        budget = _parse_amount(self.edit_budget.text())
        base = _parse_amount(self.edit_spent.text())
        new = self._final_spent()
        if new == base:
            self.lbl_spend_preview.setText('적으면 집행액에 더해지고, 잔액은 그만큼 줄어요.')
            return
        rem_before, rem_after = budget - base, budget - new
        rem_color = C.NEG if rem_after < 0 else C.TEXT
        self.lbl_spend_preview.setText(
            f'집행액 {_fmt(base)} → <b style="color:{C.ACCENT}">{_fmt(new)}원</b> · '
            f'잔액 {_fmt(rem_before)} → <b style="color:{rem_color}">{_fmt(rem_after)}원</b>')

    def _delete(self):
        ans = QMessageBox.question(self, '항목 삭제', f"'{self._item['description']}' 항목을 삭제할까요?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans == QMessageBox.Yes:
            self.deleted = True
            self.accept()

    def _show_cost_list(self):
        self._cost_completer.setCompletionPrefix('')
        self._cost_completer.complete()

    @staticmethod
    def _label(text):
        lbl = QLabel(text)
        lbl.setStyleSheet(f'font-size: 10pt; font-weight: 700; color: {C.MUTED};')
        return lbl

    def _save(self):
        desc = self.edit_desc.text().strip()
        if not desc:
            QMessageBox.warning(self, '저장 실패', '산출내역(내용)을 입력해주세요.')
            return
        cost = self.edit_cost.text().strip() or '일반'
        expense = None
        add = _parse_amount(self.edit_spend_now.text()) if self.edit_spend_now else 0
        if add > 0:
            from datetime import datetime
            spent_on = self.edit_spend_date.text().strip()
            try:
                datetime.strptime(spent_on, '%Y-%m-%d')
            except ValueError:
                QMessageBox.warning(self, '저장 실패', '날짜를 2026-11-10처럼 적어 주세요.')
                return
            expense = {'amount': add, 'purpose': self.edit_purpose.text().strip(), 'spent_on': spent_on}
        self.result_data = {
            'category_id': self.combo_cat.currentData(),
            'cost_type': cost,
            'description': desc,
            'budget': _parse_amount(self.edit_budget.text()),
            'spent': self._final_spent(),
            'expense': expense,
        }
        self.accept()


# ── 세부항목(추경) 관리 다이얼로그 ────────────────────────────────
class _CategoryManagerDialog(QDialog):
    def __init__(self, parent, on_changed):
        super().__init__(parent)
        self.setWindowTitle('세부항목(추경) 관리')
        self.setStyleSheet(STYLE)
        self.setMinimumWidth(460)
        self.setMinimumHeight(380)
        self._on_changed = on_changed

        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        lbl = QLabel('추경 발생 시 세부항목을 새로 추가하거나 명칭 변경, 삭제가 가능합니다.')
        lbl.setObjectName('Muted')
        v.addWidget(lbl)

        add_row = QHBoxLayout()
        add_row.setSpacing(8)
        self.edit_new = QLineEdit()
        self.edit_new.setPlaceholderText('새 세부항목 이름 (예: [목적]신규사업)')
        self.edit_new.returnPressed.connect(self._add_category)
        btn_add = QPushButton('추가')
        btn_add.setObjectName('Primary')
        btn_add.clicked.connect(self._add_category)
        add_row.addWidget(self.edit_new, stretch=1)
        add_row.addWidget(btn_add)
        v.addLayout(add_row)

        self.list_scroll = QScrollArea()
        self.list_scroll.setWidgetResizable(True)
        self.list_scroll.setStyleSheet(
            f'QScrollArea {{ border: 1px solid {C.BORDER}; border-radius: 10px; background: {C.SURFACE}; }}')
        self.list_container = QWidget()
        self.list_layout = QVBoxLayout(self.list_container)
        self.list_layout.setContentsMargins(6, 6, 6, 6)
        self.list_layout.setSpacing(4)
        self.list_scroll.setWidget(self.list_container)
        v.addWidget(self.list_scroll, stretch=1)

        btns = QHBoxLayout()
        self.btn_import = QPushButton('엑셀 가져오기')
        self.btn_import.setObjectName('Ghost')
        self.btn_import.setIcon(qta.icon('fa5s.file-import', color=C.MUTED))
        self.btn_import.setIconSize(QSize(14, 14))
        self.btn_import.setToolTip('사업관리카드(예산) 엑셀 파일(.xls/.xlsx)을 가져와 '
                                   '세부항목을 자동 생성합니다.')
        self.btn_import.clicked.connect(lambda: self._import_excel())
        btns.addWidget(self.btn_import)
        btns.addStretch()
        btn_close = QPushButton('닫기')
        btn_close.clicked.connect(self.accept)
        btns.addWidget(btn_close)
        v.addLayout(btns)

        self._rebuild_list()

    def _import_excel(self):
        win = self.parent()
        if win is None or not hasattr(win, '_choose_excel'):
            return
        win._choose_excel(parent=self)
        self._rebuild_list()

    def _rebuild_list(self):
        while self.list_layout.count():
            w = self.list_layout.takeAt(0).widget()
            if w:
                w.setParent(None)
                w.deleteLater()

        cats = get_budget_categories()
        if not cats:
            empty = QLabel('등록된 세부항목이 없습니다.\n'
                           '아래 [엑셀 가져오기]로 사업관리카드 파일을 불러오거나,\n'
                           '위에서 직접 추가하세요.')
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(f'color: {C.FAINT}; padding: 20px;')
            self.list_layout.addWidget(empty)
        for cat in cats:
            row = QWidget()
            rl = QHBoxLayout(row)
            rl.setContentsMargins(2, 2, 2, 2)
            rl.setSpacing(6)
            edit = QLineEdit(cat['name'])
            edit.editingFinished.connect(
                lambda cid=cat['id'], e=edit, old=cat['name']: self._rename_category(cid, e, old))
            btn_del = _icon_btn('fa5s.trash-alt', '세부항목 삭제', color='#E11D48', size=13)
            btn_del.clicked.connect(lambda _, cid=cat['id']: self._delete_category(cid))
            rl.addWidget(edit, stretch=1)
            rl.addWidget(btn_del)
            self.list_layout.addWidget(row)
        self.list_layout.addStretch()

    def _add_category(self):
        name = self.edit_new.text().strip()
        if not name:
            return
        if add_budget_category(name) is None:
            QMessageBox.warning(self, '추가 실패', '이미 존재하는 세부항목 이름입니다.')
            return
        self.edit_new.clear()
        self._rebuild_list()
        self._on_changed()

    def _rename_category(self, cat_id, edit, old_name):
        new_name = edit.text().strip()
        if not new_name or new_name == old_name:
            edit.setText(old_name)
            return
        if not rename_budget_category(cat_id, new_name):
            QMessageBox.warning(self, '변경 실패', '이미 존재하는 세부항목 이름입니다.')
            edit.setText(old_name)
            return
        self._on_changed()

    def _delete_category(self, cat_id):
        cats = {c['id']: c for c in get_budget_categories()}
        cat = cats.get(cat_id)
        if not cat:
            return
        count = len([it for it in get_budget_items(cat_id)])
        if count > 0:
            ans = QMessageBox.question(
                self, '세부항목 삭제',
                f"'{cat['name']}' 세부항목에 속한 예산 항목이 {count}개 있습니다.\n"
                '정말 함께 삭제하시겠습니까?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans != QMessageBox.Yes:
                return
        delete_budget_category(cat_id)
        self._rebuild_list()
        self._on_changed()


# ── 원가통계비목 목록 ──────────────────────────────────────────────
class _CostTypeDialog(QDialog):
    """캡처를 읽을 때 흐린 글자를 바로잡고, 항목을 입력할 때 고를 수 있는 원가통계비목 목록."""

    def __init__(self, parent=None):
        super().__init__(parent)
        import budget_types
        self._bt = budget_types
        self.setWindowTitle('원가통계비목 목록')
        self.setStyleSheet(STYLE)
        self.setMinimumSize(420, 460)
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)
        hint = QLabel('캡처를 읽을 때 한두 글자 틀린 비목 이름을 이 목록에 맞춰 고치고,\n'
                      '항목을 입력할 때 목록에서 고를 수 있어요.')
        hint.setObjectName('Muted')
        v.addWidget(hint)
        row = QHBoxLayout()
        self.edit_new = QLineEdit()
        self.edit_new.setPlaceholderText('새 원가통계비목 (예: 업무추진비)')
        self.edit_new.returnPressed.connect(self._add)
        b_add = QPushButton('추가')
        b_add.setObjectName('Primary')
        b_add.clicked.connect(self._add)
        row.addWidget(self.edit_new, 1)
        row.addWidget(b_add)
        v.addLayout(row)
        self.list_box = QWidget()
        self.list_layout = QVBoxLayout(self.list_box)
        self.list_layout.setContentsMargins(6, 6, 6, 6)
        self.list_layout.setSpacing(2)
        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setWidget(self.list_box)
        sc.setStyleSheet(f'QScrollArea {{ border: 1px solid {C.BORDER}; border-radius: 10px; background: {C.SURFACE}; }}')
        v.addWidget(sc, 1)
        self.lbl_auto = QLabel()
        self.lbl_auto.setObjectName('Muted')
        self.lbl_auto.setWordWrap(True)
        v.addWidget(self.lbl_auto)
        b_close = QPushButton('닫기')
        b_close.clicked.connect(self.accept)
        btns = QHBoxLayout()
        btns.addStretch()
        btns.addWidget(b_close)
        v.addLayout(btns)
        self._rebuild()

    def _rebuild(self):
        while self.list_layout.count():
            w = self.list_layout.takeAt(0).widget()
            if w:
                w.setParent(None)
                w.deleteLater()
        saved = self._bt.load_saved()
        for name in saved:
            r = QWidget()
            rl = QHBoxLayout(r)
            rl.setContentsMargins(8, 2, 2, 2)
            lb = QLabel(name)
            b = _icon_btn('fa5s.trash-alt', '목록에서 빼기', color='#E11D48', size=12)
            b.clicked.connect(lambda _, n=name: self._remove(n))
            rl.addWidget(lb, 1)
            rl.addWidget(b)
            self.list_layout.addWidget(r)
        self.list_layout.addStretch()
        auto = [n for n in self._bt.in_use() if n.replace(' ', '') not in {s.replace(' ', '') for s in saved}]
        self.lbl_auto.setText('지금 예산에 쓰이고 있어 자동으로 함께 쓰는 비목: ' + ', '.join(auto) if auto else '')
        self.lbl_auto.setVisible(bool(auto))

    def _add(self):
        name = self.edit_new.text().strip()
        if not name:
            return
        self._bt.save(self._bt.load_saved() + [name])
        self.edit_new.clear()
        self._rebuild()

    def _remove(self, name):
        self._bt.save([n for n in self._bt.load_saved() if n != name])
        self._rebuild()


# ── 싱글턴 창 열기 ────────────────────────────────────────────────
_instance = None


def _on_destroyed(*_args):
    global _instance
    _instance = None


def open_budget_window():
    global _instance
    if _instance is None:
        _instance = BudgetWindow()
        _instance.destroyed.connect(_on_destroyed)
    _instance.show()
    _instance.raise_()
    _instance.activateWindow()
