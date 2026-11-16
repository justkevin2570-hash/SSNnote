# -*- coding: utf-8 -*-
"""budget.py — 예산 관리 창.

화면: 위 요약 한 줄(예산현액·집행액·잔액·집행률) / 왼쪽 세부항목 목록 / 오른쪽 세부항목별 표.
  · 집행액 칸은 눌러서 바로 고치고, 완료·수정·삭제는 행에 마우스를 올리면 나온다.
  · 완료한 항목은 흐리게 아래에 늘 보인다. 집행률 색: 높을수록 진한 파랑 → 100% 넘으면 빨강.
가져오기: 엑셀 파일(끌어다 놓기) · 엑셀에서 복사한 표 · 예산 화면 캡처([가져오기] → 붙여넣기, Ctrl+V)
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
    QMenu, QRadioButton, QProgressDialog, QLayout, QCompleter, QTextBrowser,
    QInputDialog, QCheckBox
)
from PyQt5.QtCore import Qt, QEvent, QSize, QBuffer, QIODevice, QPoint, QRect, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QFontDatabase, QKeySequence, QImage, QPixmap

from db import (get_budget_categories, add_budget_category, rename_budget_category,
                delete_budget_category, get_budget_items, add_budget_item,
                update_budget_item, delete_budget_item, set_budget_item_spent,
                toggle_budget_item_completed, clear_budget_data,
                set_budget_category_excluded, set_budget_category_group, get_budget_groups,
                add_budget_group, rename_budget_group, delete_budget_group, set_budget_group_collapsed,
                get_budget_expenses, add_budget_expense, save_budget_item_edit, delete_budget_expense,
                update_budget_expense, record_unassigned_expense)
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
    """assets/font 안의 Pretendard 정적 폰트(.ttf/.otf)를 모두 등록한다. 1회만."""
    global _pretendard_loaded
    if _pretendard_loaded:
        return
    _pretendard_loaded = True
    for dirpath, _dirs, files in os.walk(os.path.join(_base_path(), 'assets', 'font')):
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
    NAV_ON      = '#EEF3FF'    # 왼쪽 목록에서 고른 줄 바탕
    # 집행 상태 색 (의미 하나에 색 하나): 진행 파랑 · 90% 이상 주황 · 초과 빨강 · 완료 회색
    OK_BAR, OK_TEXT, OK_SOFT       = '#3B6EF0', '#2F5BD3', '#EEF3FF'
    WARN_BAR, WARN_TEXT, WARN_SOFT = '#F0A040', '#A85A08', '#FFF4E5'
    OVER_BAR, OVER_TEXT, OVER_SOFT = '#E5484D', '#BE123C', '#FDECEC'
    DONE_BAR                       = '#C9CED6'
    TAG_BG      = '#F2F4F7'    # 산출내역 옆 비목 꼬리표
    TOTAL_BG    = '#FAFBFC'    # 합계 줄 · 표 머리
    SPEND_BG, SPEND_LINE = '#F5F8FF', '#C9D6F7'   # [쓴 돈 추가] 단추
    HEAD_BG, HEAD_TEXT = '#FAFBFC', '#64748B'     # 표 머리줄
    # 모서리 둥글기 (px)
    R_CARD, R_BTN, R_ROW, R_ICON, R_PILL, R_TAG, R_BAR = 12, 8, 7, 6, 11, 4, 3
    # 사무형 디자인에서만 켜는 모양
    OFFICE = False             # 왼쪽 목록 진한 바탕 · 표 격자선 · 90% 이상 줄 바탕 · 요약 칸 세로선
    NAV_BG, NAV_TEXT, NAV_MUTED = '#1F2925', '#D9E0DB', '#9FB0A7'
    NAV_ON_BG, NAV_HOVER, NAV_MARK, NAV_TRACK = '#2E3F37', '#2A3832', '#6FCF97', '#3B4B44'
    NAV_OK, NAV_WARN, NAV_OVER, NAV_DONE = '#6FCF97', '#F2C46D', '#FF9C8F', '#5C6B64'
    GRID        = '#E1E3DE'    # 사무형 표 격자선
    NAV_ALL_BG, NAV_ALL_LINE = '#F3F5F9', '#E3E7EE'   # 왼쪽 목록 「전체」 카드
    NAV_GUIDE   = '#DDE2EA'    # 그룹 안 세부항목 왼쪽 안내선
    NAV_ALL_BG_D, NAV_GUIDE_D = '#2A3832', '#3B4B44'  # 사무형(진한 바탕)용


# ── 디자인 (⋯ → 디자인): 기본 · 사무형. 틀(칸 · 단추 · 기능)은 같고 색과 모양만 다르다 ──────────
DESIGNS = {'base': '기본', 'office': '사무형'}
DESIGN_KEY = 'budget_design'
_BASE_COLORS = {k: v for k, v in vars(C).items() if k.isupper()}
_OFFICE_COLORS = {
    'BG': '#EDEDEA', 'SURFACE': '#FFFFFF', 'BORDER': '#D5D8D2', 'BORDER_HI': '#C4C7C1', 'LINE': '#E1E3DE',
    'TEXT': '#1B2420', 'BODY': '#22302A', 'MUTED': '#55605A', 'FAINT': '#6B746F', 'HOVER': '#F3F4F1',
    'ACCENT': '#1E6B47', 'ACCENT_HI': '#17573A', 'ACCENT_SOFT': '#E8F2EC', 'ACCENT_LINE': '#B9CDBF',
    'WARN_BG': '#FFF7E6', 'NAV_ON': '#2E3F37',
    'OK_BAR': '#2E8B5A', 'OK_TEXT': '#1E6B47', 'OK_SOFT': '#E8F2EC',
    'WARN_BAR': '#E0A040', 'WARN_TEXT': '#8A4B00', 'WARN_SOFT': '#FFF7E6',
    'DONE_BAR': '#C4C7C1', 'TAG_BG': '#F3F4F1', 'TOTAL_BG': '#F3F4F1',
    'SPEND_BG': '#FFFFFF', 'SPEND_LINE': '#B9CDBF', 'HEAD_BG': '#F3F4F1', 'HEAD_TEXT': '#3C4742',
    'R_CARD': 3, 'R_BTN': 3, 'R_ROW': 3, 'R_ICON': 3, 'R_PILL': 3, 'R_TAG': 2, 'R_BAR': 0,
    'OFFICE': True,
}


def current_design():
    try:
        from db import get_setting
        d = get_setting(DESIGN_KEY, 'base')
    except Exception:
        d = 'base'
    return d if d in DESIGNS else 'base'


# ── 글자 크기 (⋯ → 글자 크기): 기본 · 크게 · 더 크게 ─────────────────────
TEXT_SIZES = {'base': ('기본', 1.0), 'large': ('크게', 1.15), 'xlarge': ('더 크게', 1.3)}
TEXT_SIZE_KEY = 'budget_text_size'
_TEXT_SCALE = 1.0


def current_text_size():
    try:
        from db import get_setting
        t = get_setting(TEXT_SIZE_KEY, 'base')
    except Exception:
        t = 'base'
    return t if t in TEXT_SIZES else 'base'


def apply_text_size(name):
    """글자 크기 배율을 정하고 QSS · 칸 크기 계산을 다시 한다."""
    global _TEXT_SCALE, STYLE, _SIZES
    name = name if name in TEXT_SIZES else 'base'
    _TEXT_SCALE = TEXT_SIZES[name][1]
    STYLE = _build_style()
    _SIZES = None
    return name


def _pt(n):
    """글자 크기(pt)에 배율을 곱한다 (0.5pt 단위)."""
    return round(n * _TEXT_SCALE * 2) / 2


def _sz(n):
    """글자를 따라 커져야 하는 고정 크기(px)."""
    return int(round(n * _TEXT_SCALE))


def apply_design(name):
    """색표를 바꾸고 전체 QSS를 다시 만든다. 이미 열린 창은 BudgetWindow.set_design으로 다시 그린다."""
    global STYLE
    name = name if name in DESIGNS else 'base'
    for k, v in _BASE_COLORS.items():
        setattr(C, k, v)
    if name == 'office':
        for k, v in _OFFICE_COLORS.items():
            setattr(C, k, v)
    STYLE = _build_style()
    globals()['_SIZES'] = None
    return name


# ── 전체 QSS ─────────────────────────────────────────────────────
def _build_style():
    return _BASE_QSS() + (_OFFICE_QSS() if C.OFFICE else '')


def _BASE_QSS():
    return f"""
    QMainWindow, QDialog {{ background: {C.BG}; }}
    QWidget {{ font-family: {_UI_FONT_QSS}; font-size: {_pt(10)}pt; color: {C.TEXT}; }}
    QLabel {{ background: transparent; }}
    QScrollArea {{ border: none; background: transparent; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QFrame#Card {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_CARD}px; }}
    QFrame#VSep {{ background: {C.BORDER}; border: none; }}
    QLabel#Title {{ font-size: {_pt(16)}pt; font-weight: 800; }}
    QLabel#Meta {{ color: {C.MUTED}; font-size: {_pt(10)}pt; }}
    QLabel#StatLabel {{ color: {C.MUTED}; font-size: {_pt(10)}pt; font-weight: 600; }}
    QLabel#StatValue {{ font-size: {_pt(15)}pt; font-weight: 800; }}
    QLabel#SectionTitle {{ font-size: {_pt(12)}pt; font-weight: 800; }}
    QLabel#SectionMeta {{ color: {C.MUTED}; font-size: {_pt(10)}pt; }}
    QLabel#Muted {{ color: {C.MUTED}; }}
    QLabel#EmptyTitle {{ font-size: {_pt(13)}pt; font-weight: 800; }}

    QPushButton {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_BTN}px;
        padding: 6px 14px; color: {C.TEXT}; font-weight: 600;
    }}
    QPushButton:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton:pressed {{ background: {C.BORDER}; }}
    QPushButton:disabled {{ color: {C.FAINT}; background: {C.SURFACE}; border-color: {C.BORDER}; }}
    QPushButton::menu-indicator {{ subcontrol-position: right center; right: 8px; width: 8px; }}
    QPushButton#Primary {{ background: {C.ACCENT}; border: 1px solid {C.ACCENT}; color: white; font-weight: 700; }}
    QPushButton#Primary:hover {{ background: {C.ACCENT_HI}; border-color: {C.ACCENT_HI}; }}
    QPushButton#Primary:disabled {{ background: {C.ACCENT_LINE}; border-color: {C.ACCENT_LINE}; color: white; }}
    QPushButton#IconBtn {{ background: transparent; border: none; border-radius: {C.R_ICON}px; padding: 0; }}
    QPushButton#IconBtn:hover {{ background: {C.HOVER}; }}
    QPushButton#IconBtn:disabled {{ background: transparent; }}
    QPushButton#RowBtn {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_ROW}px;
        padding: 0 10px; font-size: {_pt(10)}pt; font-weight: 600; color: {C.BODY};
    }}
    QPushButton#RowBtn:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton#RowDone {{
        background: #ECFDF5; border: 1px solid #A7F3D0; border-radius: {C.R_ROW}px;
        padding: 0 10px; font-size: {_pt(10)}pt; font-weight: 700; color: #047857;
    }}
    QPushButton#RowDone:hover {{ background: #D1FAE5; border-color: #6EE7B7; }}
    QPushButton#Danger {{ background: {C.SURFACE}; border: 1px solid #FBCFD8; color: #E11D48; font-weight: 700; }}
    QPushButton#Danger:hover {{ background: #FFF1F3; border-color: #F9A8B8; }}
    QPushButton#Link {{ background: transparent; border: none; color: {C.MUTED}; padding: 4px 2px; text-align: left; }}
    QPushButton#Link:hover {{ color: {C.TEXT}; }}
    QPushButton#QuickBtn {{
        background: {C.BG}; border: 1px solid {C.BORDER}; border-radius: {C.R_PILL}px; padding: 3px 11px;
        font-size: {_pt(10)}pt; font-weight: 600; color: {C.BODY};
    }}
    QPushButton#QuickBtn:hover {{ border-color: {C.ACCENT_LINE}; color: {C.ACCENT_HI}; background: {C.ACCENT_SOFT}; }}
    QTextBrowser#Answer {{
        background: {C.BG}; border: none; border-radius: {C.R_BTN}px; padding: 8px 10px; color: {C.TEXT};
    }}
    QFrame#Proposal {{ background: {C.ACCENT_SOFT}; border: 1px solid {C.ACCENT_LINE}; border-radius: {C.R_BTN}px; }}

    QFrame#Nav {{ background: transparent; border: none; border-radius: {C.R_BTN}px; }}
    QFrame#StatSep {{ background: transparent; border: none; }}
    QFrame#Nav:hover {{ background: {C.HOVER}; }}
    QFrame#Nav[active="true"] {{ background: {C.NAV_ON}; }}
    QLabel#NavName {{ font-weight: 600; }}
    QLabel#NavMeta {{ color: {C.FAINT}; font-size: {_pt(10)}pt; }}
    QFrame#Nav[indent="true"] {{ margin-left: 14px; border-left: 2px solid {C.NAV_GUIDE};
        border-top-left-radius: 0; border-bottom-left-radius: 0; }}
    QFrame#Nav[kind="all"] {{ background: {C.NAV_ALL_BG}; border: 1px solid {C.NAV_ALL_LINE}; }}
    QFrame#Nav[kind="all"][active="true"] {{ background: {C.NAV_ON}; border: 1px solid {C.ACCENT_LINE}; }}
    QFrame#Nav[kind="all"] QLabel#NavName {{ font-size: {_pt(12)}pt; font-weight: 800; }}
    QFrame#NavDivider {{ background: {C.BORDER}; border: none; }}
    QLabel#NavCount {{ background: {C.HOVER}; color: {C.FAINT}; border-radius: {C.R_TAG}px; padding: 0 5px;
        font-size: {_pt(10)}pt; font-weight: 600; }}
    QFrame#Nav[excluded="true"] QLabel#NavName {{ color: {C.FAINT}; }}
    QFrame#Nav[group="true"] QLabel#NavName {{ color: {C.TEXT}; font-weight: 800; }}
    QLabel#ExTag {{
        background: {C.HOVER}; color: {C.FAINT}; border-radius: {C.R_ICON}px; padding: 1px 6px;
        font-size: {_pt(10)}pt; font-weight: 600;
    }}
    QPushButton#NavFold {{ background: transparent; border: none; padding: 0; color: {C.MUTED}; font-weight: 800; }}

    QTableWidget {{
        background: {C.SURFACE}; border: none; gridline-color: transparent;
        selection-background-color: transparent; selection-color: {C.TEXT}; outline: 0;
    }}
    QTableWidget::item {{ border-bottom: 1px solid {C.LINE}; padding: 0 10px; }}
    QTableWidget::item:selected {{ background: transparent; color: {C.TEXT}; }}
    QTableWidget QLineEdit {{ border: 1px solid {C.ACCENT}; border-radius: 4px; padding: 2px 6px; }}
    QHeaderView::section {{
        background: {C.HEAD_BG}; color: {C.HEAD_TEXT}; padding: 6px 10px; border: none;
        border-top: 1px solid {C.LINE}; border-bottom: 1px solid {C.LINE}; font-weight: 500; font-size: {_pt(10)}pt;
    }}
    QPushButton#RowIcon {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_ICON}px; padding: 0; }}
    QPushButton#RowIcon:hover {{ background: {C.HOVER}; border-color: {C.BORDER_HI}; }}
    QPushButton#RowSpend {{ background: {C.SPEND_BG}; border: 1px solid {C.SPEND_LINE}; border-radius: {C.R_ICON}px; padding: 0 9px;
        color: {C.OK_TEXT}; font-weight: 600; font-size: {_pt(10)}pt; }}
    QPushButton#RowSpend:hover {{ background: {C.OK_SOFT}; border-color: {C.OK_BAR}; }}
    QLabel#CostTag {{ background: {C.TAG_BG}; color: {C.MUTED}; border-radius: {C.R_TAG}px; padding: 1px 6px; font-size: {_pt(10)}pt; }}
    QLabel#Legend {{ color: {C.FAINT}; font-size: {_pt(10)}pt; }}
    QLabel#Pill {{ border-radius: {C.R_PILL}px; padding: 3px 10px; font-weight: 700; font-size: {_pt(10)}pt; }}
    QLineEdit, QComboBox {{
        background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_BTN}px; padding: 6px 9px;
    }}
    QLineEdit:focus, QComboBox:focus {{ border: 1px solid {C.ACCENT}; }}
    QRadioButton {{ spacing: 8px; }}
    QProgressBar {{ background: {C.LINE}; border: none; border-radius: {C.R_BAR}px; }}
    QProgressBar::chunk {{ background: {C.ACCENT}; border-radius: {C.R_BAR}px; }}
    QMenu {{ background: {C.SURFACE}; border: 1px solid {C.BORDER}; border-radius: {C.R_BTN}px; padding: 6px; }}
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


def _OFFICE_QSS():
    """사무형에만 더하는 모양: 진한 왼쪽 목록 · 격자표 머리 · 요약 칸 세로선."""
    return f"""
    QFrame#Card[nav="true"] {{ background: {C.NAV_BG}; border: 1px solid {C.NAV_BG}; }}
    QFrame#Card[nav="true"] QLabel#StatLabel {{ color: {C.NAV_MUTED}; }}
    QFrame#Card[nav="true"] QFrame#Nav {{ border-radius: 0; border-left: 3px solid transparent; }}
    QFrame#Card[nav="true"] QFrame#Nav:hover {{ background: {C.NAV_HOVER}; }}
    QFrame#Card[nav="true"] QFrame#Nav[active="true"] {{ background: {C.NAV_ON_BG}; border-left: 3px solid {C.NAV_MARK}; }}
    QFrame#Card[nav="true"] QLabel#NavName {{ color: {C.NAV_TEXT}; }}
    QFrame#Card[nav="true"] QFrame#Nav[active="true"] QLabel#NavName {{ color: white; }}
    QFrame#Card[nav="true"] QFrame#Nav[group="true"] QLabel#NavName {{ color: {C.NAV_MUTED}; }}
    QFrame#Card[nav="true"] QFrame#Nav[excluded="true"] QLabel#NavName {{ color: {C.NAV_MUTED}; }}
    QFrame#Card[nav="true"] QLabel#NavMeta {{ color: {C.NAV_MUTED}; }}
    QFrame#Card[nav="true"] QLabel#ExTag {{ background: {C.NAV_TRACK}; color: {C.NAV_TEXT}; }}
    QFrame#Card[nav="true"] QPushButton#NavFold {{ color: {C.NAV_MUTED}; }}
    QFrame#Card[nav="true"] QFrame#Nav[indent="true"] {{ border-left: 3px solid {C.NAV_GUIDE_D}; }}
    QFrame#Card[nav="true"] QFrame#Nav[indent="true"][active="true"] {{ border-left: 3px solid {C.NAV_MARK}; }}
    QFrame#Card[nav="true"] QFrame#Nav[kind="all"] {{ background: {C.NAV_ALL_BG_D}; border: none; border-left: 3px solid transparent; }}
    QFrame#Card[nav="true"] QFrame#Nav[kind="all"][active="true"] {{ background: {C.NAV_ON_BG}; border-left: 3px solid {C.NAV_MARK}; }}
    QFrame#Card[nav="true"] QFrame#Nav[kind="all"] QLabel#NavName {{ color: white; }}
    QFrame#Card[nav="true"] QFrame#Nav[group="true"] QLabel#NavName {{ color: white; }}
    QFrame#Card[nav="true"] QFrame#NavDivider {{ background: {C.NAV_GUIDE_D}; }}
    QFrame#Card[nav="true"] QLabel#NavCount {{ background: {C.NAV_TRACK}; color: {C.NAV_TEXT}; }}
    QFrame#Card[nav="true"] QScrollBar::handle:vertical {{ background: {C.NAV_TRACK}; }}
    QFrame#StatSep {{ background: {C.LINE}; }}
    QHeaderView::section {{ border: none; border-right: 1px solid {C.BORDER}; border-bottom: 1px solid {C.BORDER}; }}
    QTableWidget {{ gridline-color: {C.GRID}; }}
    QTableWidget::item {{ border-bottom: none; }}
"""


STYLE = _build_style()

_STATE_FILE = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'budget_state.json')

_ROW_H = 38

# 열: 산출내역(+비목 꼬리표) · 예산현액 · 잔액 · 집행률 · [✎][✓] (마우스를 올린 줄에만). 집행액 칸은 두지 않는다.
# 넉넉할 때의 폭과, 좁아져도 이 아래로는 줄이지 않는 최소 폭(100% 배율 기준, _sizes()가 배율을 곱함).
# 집행률·버튼 칸은 글꼴 폭으로 계산한다. 남는 폭은 산출내역이 가져간다.
_COL_W = (300, 118, 118, 0)
_COL_MIN = (150, 96, 96, 0)
_COLS = ('산출내역', '예산현액', '잔액', '집행률')
_COL_BUDGET, _COL_REM, _COL_PCT = 1, 2, 3     # 줄 단추 칸은 없앰 (쓴 돈 추가 · 수정 · 완료 · 삭제는 우클릭)

# 집행률 색: 0%는 연한 파랑, 100%에 가까울수록 진한 파랑, 100% 초과는 빨강, 완료는 회색.
# 글자용 파랑은 흰 바탕에서 읽히도록 바보다 한 단계 진하게 시작한다.
# 잔액 숫자는 고정: 양수 진한 파랑, 음수 빨강.
WARN_PCT = 90                     # 이만큼 쓰면 주황 (거의 다 씀)
def _level_soft(level):
    return {'ok': C.OK_SOFT, 'warn': C.WARN_SOFT, 'over': C.OVER_SOFT, 'done': C.HOVER}[level]


def _level(pct, done=False):
    if done:
        return 'done'
    if pct > 100:
        return 'over'
    if pct >= WARN_PCT:
        return 'warn'
    return 'ok'


def _mix(a, b, t):
    return '#%02X%02X%02X' % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def _bar_color(pct, done=False):
    return {'done': C.DONE_BAR, 'over': C.OVER_BAR, 'warn': C.WARN_BAR}.get(_level(pct, done), C.OK_BAR)


def _text_color(pct, done=False):
    return {'done': C.FAINT, 'over': C.OVER_TEXT, 'warn': C.WARN_TEXT}.get(_level(pct, done), C.OK_TEXT)


def _bal_color(rem, done=False, pct=None):
    """잔액 글자색: 완료 회색 · 초과(음수) 빨강 · 90% 이상 쓴 항목 주황 · 나머지 파랑."""
    if done:
        return C.FAINT
    if rem < 0:
        return C.OVER_TEXT
    if pct is not None and pct >= WARN_PCT:
        return C.WARN_TEXT
    return C.OK_TEXT


def _pill_qss(pct, done=False):
    return f'background: {_level_soft(_level(pct, done))}; color: {_text_color(pct, done)};'


def _bar_qss(color, track=None):
    return (f'QProgressBar {{ background: {track or C.LINE}; border: none; border-radius: {C.R_BAR}px; }}'
            f'QProgressBar::chunk {{ background: {color}; border-radius: {C.R_BAR}px; }}')


def _won(n):
    """요약 칸 금액: 사무형은 「원」을 작고 흐리게."""
    if C.OFFICE:
        return (f'{_fmt(n)}<span style="font-size: {_pt(10)}pt; font-weight:400; color:{C.MUTED}"> 원</span>')
    return f'{_fmt(n)}원'


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
    """금액 칸에 쉼표를 자동으로 붙인다. 커서는 고치던 자리 그대로 (쉼표가 생기거나 없어져도 끝으로 튀지 않게)."""
    def on_change(text):
        digits = ''.join(ch for ch in text if ch.isdigit())
        formatted = f'{int(digits):,}' if digits else ''
        if formatted != text:
            left = sum(ch.isdigit() for ch in text[:edit.cursorPosition()])   # 커서 왼쪽 숫자 개수
            if digits and int(digits) and digits[0] == '0':                    # 앞의 0이 사라진 만큼 보정
                left = max(0, left - (len(digits) - len(digits.lstrip('0'))))
            pos, seen = 0, 0
            while pos < len(formatted) and seen < left:
                if formatted[pos].isdigit():
                    seen += 1
                pos += 1
            edit.blockSignals(True)
            edit.setText(formatted)
            edit.blockSignals(False)
            edit.setCursorPosition(pos)
    edit.textChanged.connect(on_change)


def _make_strike_font(font):
    f = QFont(font)
    f.setStrikeOut(True)
    return f


def _icon_btn(icon_name, tooltip, color=None, size=16):
    color = color or C.MUTED
    btn = QPushButton()
    btn.setObjectName('IconBtn')
    btn.setIcon(qta.icon(icon_name, color=color))
    btn.setIconSize(QSize(size, size))
    btn.setFixedSize(_sz(28), _sz(28))
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
        fs = fm
        pct_lbl = fm.horizontalAdvance('100.0%') + px(6)
        pad = px(20)                       # 표 칸 안쪽 여백(QSS 'padding: 0 10px')만큼 칸 속 위젯이 좁아진다
        pct_col = px(6) + px(60) + px(8) + pct_lbl + px(10) + pad
        btn_h = max(px(26), fs.height() + px(8))
        btn_w = btn_wide = btn_h            # [✎] [✓] 네모 아이콘 단추
        # 금액 칸: 글꼴이 바뀌어도(맑은 고딕 ↔ Pretendard) 9자리 금액까지 안 잘리게 실제 글꼴로 잰다
        num_w = fm.horizontalAdvance('999,999,999') + pad + px(6)       # 넉넉할 때: 9자리 금액
        num_min = fm.horizontalAdvance('99,999,999') + pad              # 줄여도 이 아래로는 안 줄임: 8자리 금액
        w = [px(x) for x in _COL_W[:3]]
        m = [px(x) for x in _COL_MIN[:3]]
        for i in (_COL_BUDGET, _COL_REM):
            w[i], m[i] = max(w[i], num_w), max(m[i], num_min)
        w, m = tuple(w) + (pct_col,), tuple(m) + (pct_col,)
        # 왼쪽 세부항목 목록: 가장 긴 보조 글과 퍼센트가 안 잘리게
        meta = QLabel('잔액 99,999,999원 · 초과 1건')
        meta.setObjectName('NavMeta')
        meta.setStyleSheet(STYLE)
        meta.ensurePolished()
        nav_pct = fm.horizontalAdvance('100%') + px(10)
        nav_w = max(px(230), meta.fontMetrics().horizontalAdvance(meta.text()) + px(44))   # 안쪽 여백
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
    for col in (_COL_PCT,):                # 집행률 칸: 화면에 실제로 그려질 크기(글꼴·배율 반영)에 맞춤
        need = _cell_widgets_width(table, col)
        col_min[col] = max(col_min[col], need)
        col_w[col] = max(col_w[col], need)
    n = len(col_w)
    if avail <= 0 or hdr.count() < n:
        return
    total = sum(col_w)
    if avail >= total:
        widths = list(col_w)
        widths[0] += avail - total
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
    return f'{cost}  {_pct(spent, budget):.0f}%'


def _cost_chip_tip(cost, budget, spent, n):
    return (f'{cost} ({n}개 항목)\n예산 {_fmt(budget)}원\n집행 {_fmt(spent)}원\n'
            f'잔액 {_fmt(budget - spent)}원 · 집행률 {_pct(spent, budget):.1f}%')


def _cost_chip_qss(budget, spent, selector='QPushButton'):
    """비목 칩: 흰 바탕 · 옅은 테두리, 집행률 글자만 상태 색 (초과는 빨강)."""
    pct = _pct(spent, budget)
    color = C.BODY if _level(pct) == 'ok' else _text_color(pct)
    return (f'{selector} {{ background: {C.SURFACE}; color: {color};'
            f' border: 1px solid {C.BORDER}; border-radius: {C.R_PILL}px; padding: 3px 10px;'
            f' font-weight: 600; font-size: {_pt(10)}pt; }}'
            f'{selector}:hover {{ border-color: {C.BORDER_HI}; }}'
            f'{selector}:checked {{ background: {C.OK_SOFT}; border: 1px solid {C.OK_BAR}; color: {C.OK_TEXT}; }}')


class _BudgetTable(QTableWidget):
    """세부항목 하나의 표. 폭에 맞춰 열을 나누고 높이를 내용에 맞춘다(안쪽 스크롤 없음).

    - 수정·완료 단추는 마우스를 올린 줄에만 보인다 (숫자가 먼저 눈에 들어오게).
    - 오른쪽 클릭하면 「쓴 돈 추가 · 수정 · 완료 · 지우기」 메뉴.
    - follow: 같은 세부항목의 다른 표(완료 항목 · 합계 줄)는 첫 표와 열 폭을 똑같이 맞춘다.
    """
    rowMenu = pyqtSignal(int, QPoint)           # 항목 id, 화면 위치 (오른쪽 클릭)

    def __init__(self):
        super().__init__()
        self.action_buttons = []           # 행마다 (수정, 완료)
        self.row_costs = []
        self.pct_labels = []
        self.row_items = {}                # 표 줄 번호 → 항목 id
        self.follow = None                 # 열 폭을 따라 할 표
        self.followers = []
        self.tint_rows = {}                # 줄 번호 → 바탕색 (사무형: 거의 다 쓴 줄)
        self._hover_row = -1
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.viewport().installEventFilter(self)

    def _sync_pct_labels(self):
        """집행률 글자 칸 폭을 창에 입혀진 실제 글꼴로 다시 잰다 (모든 줄을 같은 폭으로 → 막대가 나란히)."""
        if not self.pct_labels or not self.isVisible():
            return
        need = self.pct_labels[0].fontMetrics().horizontalAdvance('100.0%') + px(6)
        for lb in self.pct_labels:
            if lb.width() != need:
                lb.setFixedWidth(need)

    def _copy_columns(self, src):
        hdr, sh = self.horizontalHeader(), src.horizontalHeader()
        for c in range(min(hdr.count(), sh.count())):
            if self.isColumnHidden(c) != src.isColumnHidden(c):
                self.setColumnHidden(c, src.isColumnHidden(c))
            if hdr.sectionSize(c) != sh.sectionSize(c):
                hdr.resizeSection(c, sh.sectionSize(c))

    def sync_geometry(self):
        self._sync_pct_labels()
        if self.follow is not None:
            try:
                self._copy_columns(self.follow)
            except RuntimeError:
                self.follow = None
                _fit_columns(self)
        else:
            _fit_columns(self)
        hdr = self.horizontalHeader()
        rows_h = sum(self.rowHeight(r) for r in range(self.rowCount()) if not self.isRowHidden(r))
        head_h = 0 if hdr.isHidden() else max(hdr.height(), hdr.sizeHint().height())
        cols = sum(hdr.sectionSize(c) for c in range(hdr.count()) if not self.isColumnHidden(c))
        hbar = self.horizontalScrollBar().sizeHint().height() if cols > self.viewport().width() + 1 > 1 else 0
        target = head_h + rows_h + hbar + 2
        if self.height() != target:
            self.setFixedHeight(target)
        for f in list(self.followers):
            try:
                f.sync_geometry()
            except RuntimeError:
                self.followers.remove(f)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.sync_geometry()

    def paintEvent(self, event):
        # 줄 바탕색은 칸 그림(QSS ::item)이 덮어 버려서, 표 바탕에 먼저 칠한다 (칸 속 위젯은 투명)
        if self.tint_rows:
            from PyQt5.QtGui import QPainter
            p = QPainter(self.viewport())
            w = self.viewport().width()
            for r, color in self.tint_rows.items():
                if 0 <= r < self.rowCount() and not self.isRowHidden(r):
                    p.fillRect(QRect(0, self.rowViewportPosition(r), w, self.rowHeight(r)), QColor(color))
            p.end()
        if self._hover_row in self.row_items:
            from PyQt5.QtGui import QPainter
            p = QPainter(self.viewport())
            w = self.viewport().width()
            for r in (self._hover_row, self._hover_row + 1):     # 항목 줄 + 펼친 사용 내역 줄
                if 0 <= r < self.rowCount() and not self.isRowHidden(r):
                    p.fillRect(QRect(0, self.rowViewportPosition(r), w, self.rowHeight(r)), QColor(C.HOVER))
            p.end()
        super().paintEvent(event)

    # ── 마우스를 올린 줄에만 단추 ───────────────────────────────────
    def _item_row(self, r):
        """사용 내역 줄(홀수)에 마우스가 있어도 그 위 항목 줄로."""
        if r < 0:
            return -1
        return r - (r % 2) if (r - (r % 2)) in self.row_items else -1

    def set_hover_row(self, r):
        """마우스를 올린 줄: 바탕을 옅게 칠해 어느 줄을 우클릭할지 보이게."""
        if r == self._hover_row:
            return
        self.viewport().update()
        for row in (self._hover_row, r):
            pair = self.action_buttons[row] if 0 <= row < len(self.action_buttons) else None
            if pair:
                for b in pair:
                    b.setVisible(row == r)
        self._hover_row = r

    def watch_row(self, widget):
        """칸 속 위젯(산출내역 · 막대 · 단추)에 마우스가 들어가도 그 줄을 '마우스 올린 줄'로.
        칸 속 위젯 위에서는 표 바탕(viewport)에 마우스 움직임이 오지 않고 오히려 '떠남'이 오기 때문."""
        widget.installEventFilter(self)
        for child in widget.findChildren(QWidget):
            child.installEventFilter(self)

    def _sync_hover_to_cursor(self):
        from PyQt5.QtGui import QCursor
        pos = self.viewport().mapFromGlobal(QCursor.pos())
        if not self.viewport().rect().contains(pos):
            self.set_hover_row(-1)
        else:
            self.set_hover_row(self._item_row(self.rowAt(pos.y())))

    def eventFilter(self, obj, e):
        t = e.type()
        if obj is self.viewport():
            if t == QEvent.MouseMove:
                self.set_hover_row(self._item_row(self.rowAt(e.pos().y())))
            elif t == QEvent.Leave:
                self._sync_hover_to_cursor()          # 칸 속 단추로 들어간 것이면 그대로 둠
            elif t == QEvent.ContextMenu:
                r = self._item_row(self.rowAt(e.pos().y()))
                if r >= 0:
                    self.rowMenu.emit(self.row_items[r], e.globalPos())
                    return True
        elif t in (QEvent.Enter, QEvent.MouseMove, QEvent.Leave) and isinstance(obj, QWidget):
            self._sync_hover_to_cursor()
        return super().eventFilter(obj, e)


class _NavItem(QFrame):
    """왼쪽 세부항목 목록의 한 줄: 이름 · 집행률 막대 · 잔액.
    group=True면 그룹 줄(접기 ▾/▸ 단추), indent면 그룹 안 세부항목(왼쪽으로 들여 씀)."""
    clicked = pyqtSignal()
    menuRequested = pyqtSignal(QPoint)
    foldClicked = pyqtSignal()

    BAR_H = {'all': 6, 'group': 4, 'cat': 3}       # 막대 두께로도 단계를 구분 (전체 > 그룹 > 세부항목)

    def __init__(self, name, group=False, indent=False, excluded=False, collapsed=False, total=False, count=0):
        super().__init__()
        self.kind = 'all' if total else ('group' if group else 'cat')
        self.setObjectName('Nav')
        self.setProperty('active', 'false')
        self.setProperty('kind', self.kind)
        self.setProperty('group', 'true' if group else 'false')
        self.setProperty('excluded', 'true' if excluded else 'false')
        self.setProperty('indent', 'true' if indent else 'false')
        self.setCursor(Qt.PointingHandCursor)
        v = QVBoxLayout(self)
        v.setContentsMargins(10, 9 if total else 8, 10, 10 if total else 9)
        v.setSpacing(5 if total else 4)
        top = QHBoxLayout()
        top.setSpacing(4)
        self.btn_fold = None
        if group:
            # 접기 단추: 글자(▾▸) 대신 벡터 아이콘 — 화면 배율에 맞는 크기로 바로 그려 선명하게
            icon_c = C.NAV_TEXT if C.OFFICE else C.BODY
            self.btn_fold = QPushButton()
            self.btn_fold.setObjectName('NavFold')
            self.btn_fold.setIcon(qta.icon('ph.caret-right-bold' if collapsed else 'ph.caret-down-bold', color=icon_c))
            ip = px(_sz(14))
            self.btn_fold.setIconSize(QSize(ip, ip))
            self.btn_fold.setFixedSize(px(_sz(20)), px(_sz(20)))
            self.btn_fold.setCursor(Qt.PointingHandCursor)
            self.btn_fold.setToolTip('펼치기' if collapsed else '접기')
            self.btn_fold.clicked.connect(self.foldClicked.emit)
            top.addWidget(self.btn_fold, 0, Qt.AlignVCenter)
            folder = QLabel()                          # 그룹 표시: 폴더 아이콘
            fp = px(_sz(17))
            folder.setPixmap(qta.icon('ph.folder-simple-bold' if collapsed else 'ph.folder-open-bold',
                                      color=icon_c).pixmap(fp, fp))
            folder.setFixedSize(fp + 2, fp)
            top.addWidget(folder, 0, Qt.AlignVCenter)
        self.lbl_name = QLabel(name)
        self.lbl_name.setObjectName('NavName')
        self.lbl_name.setWordWrap(True)
        top.addWidget(self.lbl_name, 1)
        if group and count:
            cnt = QLabel(f'{count}개')
            cnt.setObjectName('NavCount')
            cnt.setToolTip(f'세부항목 {count}개를 묶은 그룹')
            top.addWidget(cnt, 0, Qt.AlignVCenter)
        if excluded:
            tag = QLabel('합계 제외')
            tag.setObjectName('ExTag')
            tag.setToolTip('전체 합계에 넣지 않는 세부항목이에요 (우클릭 → 전체 합계에 넣기)')
            top.addWidget(tag, 0, Qt.AlignTop)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(self.BAR_H[self.kind])
        self.lbl_pct = QLabel()
        self.lbl_pct.setObjectName('NavMeta')
        self.lbl_pct.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lbl_pct.setFixedWidth(_sizes()['nav_pct'] + (px(_sz(12)) if total else 0))
        row.addWidget(self.bar, 1)
        row.addWidget(self.lbl_pct)
        self.lbl_meta = QLabel()
        self.lbl_meta.setObjectName('NavMeta')
        self.lbl_meta.setWordWrap(True)
        v.addLayout(top)
        v.addLayout(row)
        v.addWidget(self.lbl_meta)
        if self.kind != 'cat':                    # 전체 · 그룹: 막대 · 집행률 없이 (집행률은 가운데 요약 · 머리에)
            self.bar.hide()
            self.lbl_pct.hide()

    def set_active(self, on):
        self.setProperty('active', 'true' if on else 'false')
        for w in (self, self.lbl_name):
            w.style().unpolish(w)
            w.style().polish(w)

    def set_values(self, budget, spent, extra=''):
        pct = _pct(spent, budget)
        self.bar.setValue(int(round(min(pct, 100))))
        rem = budget - spent
        if C.OFFICE:                                    # 진한 바탕용 밝은 색
            lv = _level(pct)
            bar = {'over': C.NAV_OVER, 'warn': C.NAV_WARN}.get(lv, C.NAV_OK)
            txt = {'over': C.NAV_OVER, 'warn': C.NAV_WARN}.get(lv, C.NAV_TEXT)
            self.bar.setStyleSheet(_bar_qss(bar, C.NAV_TRACK))
            meta = C.NAV_OVER if rem < 0 else C.NAV_MUTED
        else:
            self.bar.setStyleSheet(_bar_qss(_bar_color(pct)))
            txt = _text_color(pct)
            meta = C.OVER_TEXT if rem < 0 else C.FAINT
        self.lbl_pct.setText(f'{pct:.0f}%')
        big = f' font-size: {_pt(12)}pt; font-weight: 800;' if self.kind == 'all' else ' font-weight: 600;'
        self.lbl_pct.setStyleSheet(f'color: {txt};{big}')
        self.lbl_meta.setText(f'<span style="color:{meta}">잔액 {_fmt(rem)}원</span>{extra}')

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()

    def contextMenuEvent(self, e):
        self.menuRequested.emit(e.globalPos())
        e.accept()


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
        self.groups = []
        self._overview_scope = None       # 원가통계비목별 합계 범위 (None이면 전체 합계)
        self.active_tab = 'ALL'          # 'ALL' · 세부항목 id · ('G', 그룹 id)
        self._items_by_id = {}
        self._row_refs = {}
        self._sec_refs = {}
        self._nav_refs = {}
        self._show_done = set()           # 완료 항목을 펼쳐 둔 세부항목 (기본은 접음)
        self._cost_filter = {}            # 세부항목별로 골라 본 원가통계비목 (없으면 전부)
        self._overview_chips = {}
        self._img_worker = None
        self._ask_worker = None
        self._ask_history = []            # [(질문, 답)] — 이어지는 질문에 쓰려고 최근 것만
        self._ask_q = ''
        self._expenses_by_item = {}       # 항목별 사용 내역 (최근 날짜 먼저)
        self._expense_ui = {}             # 이번 창을 여는 동안만 기억: 펼침 · 더 보기 (창을 닫으면 사라짐)
        self._restore_window_state()
        self.text_size = apply_text_size(current_text_size())   # 기본 · 크게 · 더 크게 (⋯ → 글자 크기)
        self.design = apply_design(current_design())   # 기본 · 사무형 (⋯ → 디자인)
        self._init_ui()
        self._apply_ai_btn()
        from ui import ai_connect
        ai_connect.on_change(self._on_ai_changed)   # 공문 작성 창에서 바꿔도 바로 맞춤
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
        root.setContentsMargins(24, 18, 24, 14)
        root.setSpacing(12)

        # 머리: 제목 · 요약 글 | 가져오기 ▾ · 항목 추가 · ⋯
        header = QHBoxLayout()
        header.setSpacing(8)
        tb = QHBoxLayout()
        tb.setSpacing(10)
        title = QLabel('예산 관리')
        title.setObjectName('Title')
        self.lbl_meta = QLabel()
        self.lbl_meta.setObjectName('Meta')
        tb.addWidget(title)
        tb.addWidget(self.lbl_meta, 0, Qt.AlignBottom)
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
        self.btn_more.setFixedSize(_sz(34), _sz(34))
        self.btn_more.setToolTip('원가통계비목 목록 · 전체 지우기')
        self.btn_more.clicked.connect(lambda: self._exec_menu(self._more_menu(), self.btn_more))
        # AI 연결 (공문 작성 창과 같은 메뉴 · 같은 설정)
        self.btn_ai = QPushButton()
        self.btn_ai.setIconSize(QSize(9, 9))
        self.btn_ai.setToolTip('AI 연결 바꾸기 (OpenCode · Claude · ChatGPT · Claude 구독)')
        self.btn_ai.clicked.connect(lambda: self._ai_menu())
        for b in (self.btn_ai, self.btn_import, self.btn_manage):
            b.setFixedHeight(_sz(34))
            b.setCursor(Qt.PointingHandCursor)
        header.addWidget(self.btn_ai)
        header.addSpacing(6)
        header.addWidget(self.btn_import)
        header.addWidget(self.btn_manage)
        header.addWidget(self.btn_more)
        root.addLayout(header)

        # 요약: 예산현액 · 집행액 · 잔액 · 집행률을 같은 간격으로, 그 아래 전체 폭 집행률 막대
        self.summary = QFrame()
        self.summary.setObjectName('Card')
        sv = QVBoxLayout(self.summary)
        sv.setContentsMargins(20, 12, 20, 12)
        sv.setSpacing(10)
        sl = QHBoxLayout()
        sl.setSpacing(16)
        self.stat = {}
        for i, (key, label) in enumerate((('budget', '예산현액'), ('spent', '집행액'),
                                          ('remaining', '잔액'), ('pct', '집행률'))):
            if i:                                           # 칸 사이 세로선 (사무형에서만 보임)
                sep = QFrame()
                sep.setObjectName('StatSep')
                sep.setFixedWidth(1)
                sl.addWidget(sep)
            box = QVBoxLayout()
            box.setSpacing(1)
            lb = QLabel(label)
            lb.setObjectName('StatLabel')
            val = QLabel('0원')
            val.setObjectName('StatValue')
            box.addWidget(lb)
            box.addWidget(val)
            sl.addLayout(box, 1)
            self.stat[key] = val
        self.lbl_pct = self.stat['pct']
        self.bar_total = QProgressBar()
        self.bar_total.setRange(0, 100)
        self.bar_total.setTextVisible(False)
        self.bar_total.setFixedHeight(6)
        sv.addLayout(sl)
        sv.addWidget(self.bar_total)
        root.addWidget(self.summary)

        # 본문: 왼쪽 세부항목 목록 | 오른쪽 표
        self.body = QWidget()
        bl = QHBoxLayout(self.body)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(14)
        self.nav_card = QFrame()
        self.nav_card.setObjectName('Card')
        self.nav_card.setProperty('nav', 'true')          # 사무형에서는 진한 바탕
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
        self.scroll.verticalScrollBar().valueChanged.connect(self._schedule_save)   # 보던 자리 기억
        self.sections_container = QWidget()
        self.sections_layout = QVBoxLayout(self.sections_container)
        self.sections_layout.setContentsMargins(0, 0, 4, 0)
        self.sections_layout.setSpacing(12)
        self.scroll.setWidget(self.sections_container)
        bl.addWidget(self.scroll, 1)
        root.addWidget(self.body, 1)

        self.empty = self._make_empty()
        root.addWidget(self.empty, 1)

        self.ask_card = self._make_ask_panel()
        root.addWidget(self.ask_card)

    # ── AI에게 묻기 ────────────────────────────────────────────────
    def _make_ask_panel(self):
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(14, 6, 8, 6)
        v.setSpacing(6)

        self.ask_answer = QTextBrowser()
        self.ask_answer.setObjectName('Answer')
        self.ask_answer.setOpenLinks(False)
        self.ask_answer.hide()
        v.addWidget(self.ask_answer)

        # AI 제안 확인 칸 (AI 대화 창 · 공문 작성 창과 같은 기능 목록 · 같은 모양)
        from ui.ai_confirm import AiConfirmPanel
        self.ask_confirm = AiConfirmPanel(self._ai_app, self._ai_say, on_done=lambda: self.refresh())
        v.addWidget(self.ask_confirm)

        # 한 줄: 💬 AI에게 묻기 [입력칸 ……] 상태 · 되돌리기 · 보내기
        row = QHBoxLayout()
        row.setSpacing(8)
        icon = QLabel()
        icon.setPixmap(qta.icon('fa5s.comment-dots', color=C.ACCENT).pixmap(14, 14))
        title = QLabel('AI에게 묻기')
        title.setStyleSheet(f'color: {C.ACCENT}; font-weight: 700;')
        self.quick_buttons = []           # 빠른 질문 버튼은 없앰 (AI 대화 창과 같게)
        self.ask_input = QLineEdit()
        self.ask_input.setFixedHeight(_sz(32))
        self.ask_input.setStyleSheet('QLineEdit { border: none; background: transparent; padding: 0 4px; }')
        self.ask_input.returnPressed.connect(lambda: self._ask(self.ask_input.text()))
        self.ask_status = QLabel()
        self.ask_status.setObjectName('Muted')
        self.btn_undo_apply = QPushButton('되돌리기')
        self.btn_undo_apply.setObjectName('QuickBtn')
        self.btn_undo_apply.setCursor(Qt.PointingHandCursor)
        self.btn_undo_apply.setToolTip('방금 [넣기]·[지우기]로 한 일을 전으로 되돌립니다')
        self.btn_undo_apply.clicked.connect(self._undo_apply)
        self.btn_undo_apply.hide()
        self.ask_send = QPushButton()
        self.ask_send.setObjectName('Primary')
        self.ask_send.setIcon(qta.icon('fa5s.paper-plane', color='white'))
        self.ask_send.setFixedSize(_sz(36), _sz(30))
        self.ask_send.setToolTip('보내기 (Enter)')
        self.ask_send.setCursor(Qt.PointingHandCursor)
        self.ask_send.clicked.connect(lambda: self._ask(self.ask_input.text()))
        row.addWidget(icon)
        row.addWidget(title)
        row.addWidget(self.ask_input, 1)
        row.addWidget(self.ask_status)
        row.addWidget(self.btn_undo_apply)
        row.addWidget(self.ask_send)
        v.addLayout(row)
        return card

    def _ai_ready(self):
        try:
            from ai_client import load_ai_mode
            return load_ai_mode() != 'none'
        except Exception:
            return False

    # ── AI 연결 버튼 ───────────────────────────────────────────────
    def _apply_ai_btn(self):
        from ui import ai_connect
        try:
            text, state = ai_connect.status()
        except Exception:
            text, state = 'AI 꺼짐 · 연결하기', 'off'
        icon, color = {'on': ('fa5s.circle', '#16A34A'), 'legacy': ('fa5s.exclamation-circle', '#D97706')}.get(
            state, ('fa5s.circle', C.FAINT))
        if state == 'off':                     # 꺼져 있으면 누를 수 있는 단추로 보이게
            self.btn_ai.setText('AI 연결하기')
            self.btn_ai.setIcon(qta.icon('fa5s.magic', color=C.ACCENT))
            self.btn_ai.setIconSize(QSize(12, 12))
            self.btn_ai.setStyleSheet(f'QPushButton {{ color: {C.ACCENT}; }}')
            return
        self.btn_ai.setStyleSheet('')
        self.btn_ai.setText(text)
        self.btn_ai.setIcon(qta.icon(icon, color=color))
        self.btn_ai.setIconSize(QSize(12, 12) if state == 'legacy' else QSize(9, 9))

    def _ai_menu(self):
        from ui import ai_connect
        choice = ai_connect.show_menu(self, self.btn_ai, STYLE)
        if choice == 'none':
            ai_connect.turn_off()
        elif choice:
            ai_connect.connect(self, choice)

    def _on_ai_changed(self):
        try:
            self._apply_ai_btn()
            self._update_ask_enabled()
        except RuntimeError:          # 이미 닫힌 창
            pass

    def _update_ask_enabled(self):
        busy = self._ask_worker is not None and self._ask_worker.isRunning()
        ready = self._ai_ready()
        for w in [self.ask_send] + self.quick_buttons:
            w.setEnabled(ready and not busy)
        self.ask_input.setEnabled(not busy)            # AI가 꺼져 있어도 /clear는 되게
        self.ask_input.setPlaceholderText(
            '예산에 대해 묻거나 지시하세요 — 예: 협의회비 잔액 30만원으로 맞춰줘 · /clear 대화 지우기'
            if ready else 'AI가 꺼져 있어요 — 위의 [AI 연결하기]를 누르면 쓸 수 있어요')

    def _ai_app(self):
        from ui.ai_app import AppAdapter
        return AppAdapter(on_clear=self._clear_ask)

    def _ai_say(self, text, undo=False):
        self.ask_status.setText(text)
        self.btn_undo_apply.setVisible(bool(undo) and self.ask_confirm.can_undo())

    def _ask(self, question):
        q = (question or '').strip()
        if not q or (self._ask_worker is not None and self._ask_worker.isRunning()):
            return
        if q.startswith('/'):                          # 입력칸 명령 (AI에게 보내지 않음)
            self.ask_input.clear()
            if q.lower() == '/clear':
                self._clear_ask()
            else:
                self.ask_status.setText(f'모르는 명령이에요: {q} — /clear(대화 지우기)만 쓸 수 있어요.')
            return
        if not self._ai_ready():
            self.ask_status.setText('AI가 꺼져 있어요 — 위의 [AI 꺼짐 · 연결하기]를 눌러 연결하세요.')
            return
        import ai_actions
        self._ask_q = q
        self.ask_input.clear()
        self._hide_proposal()
        self._set_answer(f'Q. {q}\n\n생각하는 중…')
        self.ask_status.setText('AI가 보는 중…')
        ctx = ai_actions.build_state(self._ai_app())
        w = ai_actions.AskWorker(q, ctx, self._ask_history)
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
        self.ask_status.setText('')
        done = self.ask_confirm.show_reply(reply)
        if done:
            text += '\n\n' + ' · '.join(done) + ' 했어요.'
        self._set_answer(text)

    def _on_ask_failed(self, msg):
        self._set_answer(f'Q. {self._ask_q}\n\n⚠ {msg}')
        self.ask_status.setText('')

    def _clear_ask(self):
        """/clear: 답 칸 · 이전 질문 기억 · 넣기 확인 칸 지우기 (넣은 예산은 그대로)."""
        self._ask_history = []
        self._ask_q = ''
        self._hide_proposal()
        self.ask_answer.clear()
        self.ask_answer.hide()
        self.ask_status.setText('대화를 지웠어요.')

    def _set_answer(self, text):
        """답 칸: 글 길이에 맞춰 높이를 정하되 170px까지 (넘으면 스크롤)."""
        a = self.ask_answer
        a.setPlainText(text)
        a.show()
        doc = a.document()
        doc.setTextWidth(max(200, a.viewport().width() or self.width() - 90))
        a.setFixedHeight(int(min(170, max(52, doc.size().height() + 20))))

    # ── 말로 지출 입력: AI가 제안 → [넣기]를 눌러야 반영 ─────────────
    def _hide_proposal(self):
        self.ask_confirm.hide_changes()
        self.ask_confirm.hide_deletes()

    def _undo_apply(self):
        self.ask_confirm.undo()
        self.btn_undo_apply.hide()

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
                   '예산 화면을 캡처해 붙여넣으면([가져오기] → 붙여넣기 또는 Ctrl+V) 세부항목별로 정리해 드려요.')
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
        m.addAction('화면 배치 기본값으로', lambda: self.reset_layout())
        dm = m.addMenu('디자인')
        dm.setStyleSheet(STYLE)
        for key, label in DESIGNS.items():
            act = dm.addAction(label + ('  (초록 · 격자표)' if key == 'office' else '  (지금 모습)'),
                               lambda k=key: self.set_design(k))
            act.setCheckable(True)
            act.setChecked(key == self.design)
        tm = m.addMenu('글자 크기')
        tm.setStyleSheet(STYLE)
        for key, (label, _scale) in TEXT_SIZES.items():
            act = tm.addAction(label, lambda k=key: self.set_text_size(k))
            act.setCheckable(True)
            act.setChecked(key == self.text_size)
        m.addSeparator()
        a = m.addAction('전체 지우기…', lambda: self._reset_all())
        a.setEnabled(bool(self.categories or self.items))
        return m

    def set_design(self, name):
        """디자인 바꾸기: 색표를 바꾸고 창 안을 다시 그린다 (보던 세부항목 · 펼침 · 스크롤은 그대로)."""
        if name not in DESIGNS:
            return
        from db import set_setting
        set_setting(DESIGN_KEY, name)
        self.design = apply_design(name)
        self._redraw()

    def set_text_size(self, name):
        """글자 크기 바꾸기 (기본 · 크게 · 더 크게): 글자와 칸 크기를 다시 재고 창 안을 다시 그린다."""
        if name not in TEXT_SIZES:
            return
        from db import set_setting
        set_setting(TEXT_SIZE_KEY, name)
        self.text_size = apply_text_size(name)
        self._redraw()

    def _redraw(self):
        """창 안을 새로 만든다 (보던 세부항목 · 펼침 · 스크롤 · 적던 질문은 그대로)."""
        pos = self.scroll.verticalScrollBar().value()
        q = self.ask_input.text()
        self.setStyleSheet(STYLE)
        self._init_ui()
        self._apply_ai_btn()
        self.refresh(keep_scroll=False)
        self.ask_input.setText(q)
        keep_light(self)
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(0, lambda: self.scroll.verticalScrollBar().setValue(pos))

    def contextMenuEvent(self, e):
        e.accept()                        # 창 빈 곳 우클릭 메뉴는 없음 (가져오기 · 항목 추가는 위 단추와 [+]로)

    # ── 데이터 새로고침 ────────────────────────────────────────────
    def refresh(self, keep_scroll=True):
        pos = self.scroll.verticalScrollBar().value() if keep_scroll else 0
        self.categories = get_budget_categories()
        self.items = get_budget_items()
        self.groups = get_budget_groups()
        self._items_by_id = {it['id']: it for it in self.items}
        self._expenses_by_item = {}
        for ex in get_budget_expenses():
            self._expenses_by_item.setdefault(ex['item_id'], []).append(ex)
        if not self._tab_exists(self.active_tab):
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
        n_ex = len(self._excluded_ids())
        if n_ex:
            text += f' · 합계 제외 {n_ex}개'
        self.lbl_meta.setText(text + (f' · 완료 {done}개' if done else ''))

    # ── 왼쪽 목록 ──────────────────────────────────────────────────
    def _rebuild_nav(self):
        while self.nav_layout.count():
            w = self.nav_layout.takeAt(0).widget()
            if w:
                w.hide()                  # 숨긴 뒤 떼어야 지우기 전 잠깐 따로 창으로 뜨지 않음
                w.setParent(None)
                w.deleteLater()
        self._nav_refs = {}

        def add(key, nav):
            nav.clicked.connect(lambda k=key: self._set_tab(k))
            nav.menuRequested.connect(lambda pos, k=key: self._nav_menu(k, pos))
            self.nav_layout.addWidget(nav)
            self._nav_refs[key] = nav

        add('ALL', _NavItem('전체', total=True))
        div = QFrame()                        # 「전체」와 목록 사이 구분선
        div.setObjectName('NavDivider')
        div.setFixedHeight(1)
        self.nav_layout.addSpacing(4)
        self.nav_layout.addWidget(div)
        self.nav_layout.addSpacing(4)
        for g in self.groups:
            members = self._group_members(g['id'])
            if not members:
                continue                      # 빈 그룹은 [세부항목 관리]에서만 보임
            gnav = _NavItem(g['name'], group=True, collapsed=bool(g['collapsed']), count=len(members))
            gnav.foldClicked.connect(lambda gid=g['id'], c=bool(g['collapsed']): self._fold_group(gid, not c))
            add(('G', g['id']), gnav)
            if not g['collapsed']:
                for c in members:
                    add(c['id'], _NavItem(c['name'], indent=True, excluded=bool(c.get('excluded'))))
        group_ids = {g['id'] for g in self.groups}
        for c in self.categories:
            if c.get('group_id') not in group_ids:
                add(c['id'], _NavItem(c['name'], excluded=bool(c.get('excluded'))))
        self.nav_layout.addStretch()
        self._update_nav()

    def _update_nav(self):
        for key, nav in self._nav_refs.items():
            if key == 'ALL':
                b, s = self._sum(self._total_items())
                n_ex = len(self._excluded_ids())
                extra = ''
                if n_ex:
                    extra += f'<br><span style="color:{C.FAINT}">합계 제외 {n_ex}개 뺌</span>'
                nav.set_values(b, s, extra)
            elif isinstance(key, tuple):
                members = self._group_members(key[1])
                b, s = self._sum(self._items_of({c['id'] for c in members}))
                nav.set_values(b, s)
            else:
                b, s = self._cat_totals(key)
                n_over = self._over_count({key})
                nav.set_values(b, s, f' · <span style="color:{C.NAV_OVER if C.OFFICE else C.OVER_TEXT}">초과 {n_over}건</span>' if n_over else '')
            nav.set_active(key == self.active_tab)

    def _set_tab(self, tab):
        if tab == self.active_tab:
            return
        self.active_tab = tab
        self._update_nav()
        self._rebuild_sections()
        self.scroll.verticalScrollBar().setValue(0)
        self._schedule_save()

    # ── 합계 제외 · 그룹 ───────────────────────────────────────────
    def _tab_exists(self, tab):
        if tab == 'ALL':
            return True
        if isinstance(tab, tuple):
            return bool(self._group_members(tab[1]))
        return tab in {c['id'] for c in self.categories}

    def _excluded_ids(self):
        return {c['id'] for c in self.categories if c.get('excluded')}

    def _total_items(self):
        """전체 합계에 들어가는 항목 (합계에서 뺀 세부항목의 항목은 빼고)."""
        ex = self._excluded_ids()
        return [it for it in self.items if it['category_id'] not in ex]

    def _items_of(self, cat_ids):
        return [it for it in self.items if it['category_id'] in cat_ids]

    @staticmethod
    def _sum(items):
        return sum(it['budget'] for it in items), sum(it['spent'] for it in items)

    def _group_members(self, group_id):
        return [c for c in self.categories if c.get('group_id') == group_id]

    def _fold_group(self, group_id, collapsed):
        set_budget_group_collapsed(group_id, collapsed)
        self.groups = get_budget_groups()
        self._rebuild_nav()

    def _nav_menu(self, key, pos):
        if key == 'ALL':
            return
        m = self._menu()
        if isinstance(key, tuple):
            gid = key[1]
            g = next((g for g in self.groups if g['id'] == gid), None)
            members = self._group_members(gid)
            if not g:
                return
            all_ex = bool(members) and all(c.get('excluded') for c in members)
            m.addAction('그룹 전체를 전체 합계에 넣기' if all_ex else '그룹 전체를 전체 합계에서 빼기',
                        lambda: self._set_excluded([c['id'] for c in members], not all_ex))
            m.addSeparator()
            m.addAction('그룹 이름 바꾸기…', lambda: self._rename_group(gid))
            m.addAction('그룹 풀기 (세부항목은 그대로 남아요)', lambda: self._delete_group(gid))
        else:
            cat = next((c for c in self.categories if c['id'] == key), None)
            if not cat:
                return
            ex = bool(cat.get('excluded'))
            m.addAction('전체 합계에 넣기' if ex else '전체 합계에서 빼기',
                        lambda: self._set_excluded([key], not ex))
            sub = m.addMenu('그룹으로 옮기기')
            sub.setStyleSheet(STYLE)
            for g in self.groups:
                a = sub.addAction(g['name'], lambda gid=g['id']: self._move_to_group(key, gid))
                a.setCheckable(True)
                a.setChecked(cat.get('group_id') == g['id'])
            if self.groups:
                sub.addSeparator()
            sub.addAction('새 그룹 만들기…', lambda: self._move_to_new_group(key))
            if cat.get('group_id') is not None:
                m.addAction('그룹에서 빼기', lambda: self._move_to_group(key, None))
        m.exec_(pos)

    def _set_excluded(self, cat_ids, excluded):
        for cid in cat_ids:
            set_budget_category_excluded(cid, excluded)
        self.refresh()

    def _move_to_group(self, cat_id, group_id):
        set_budget_category_group(cat_id, group_id)
        self.refresh()

    def _ask_group_name(self, title, text=''):
        name, ok = QInputDialog.getText(self, title, '그룹 이름', QLineEdit.Normal, text)
        name = (name or '').strip()
        return name if ok and name else None

    def _move_to_new_group(self, cat_id):
        name = self._ask_group_name('새 그룹')
        if not name:
            return
        gid = add_budget_group(name)
        if gid is None:
            gid = next((g['id'] for g in get_budget_groups() if g['name'] == name), None)
        if gid is not None:
            self._move_to_group(cat_id, gid)

    def _rename_group(self, group_id):
        g = next((g for g in self.groups if g['id'] == group_id), None)
        if not g:
            return
        name = self._ask_group_name('그룹 이름 바꾸기', g['name'])
        if not name or name == g['name']:
            return
        if not rename_budget_group(group_id, name):
            QMessageBox.warning(self, '바꾸기 실패', '같은 이름의 그룹이 이미 있어요.')
            return
        self.refresh()

    def _delete_group(self, group_id):
        delete_budget_group(group_id)
        if self.active_tab == ('G', group_id):
            self.active_tab = 'ALL'
        self.refresh()

    # ── 오른쪽 표 ──────────────────────────────────────────────────
    def _rebuild_sections(self):
        self._row_refs = {}
        self._sec_refs = {}
        while self.sections_layout.count():
            w = self.sections_layout.takeAt(0).widget()
            if w:
                w.hide()                  # 숨긴 뒤 떼어야 지우기 전 잠깐 따로 창으로 뜨지 않음
                w.setParent(None)
                w.deleteLater()
        self._overview_chips = {}
        tab = self.active_tab
        if tab == 'ALL':
            target = self.categories
            self._overview_scope = None
            if self.items:
                self.sections_layout.addWidget(self._make_cost_overview())
        elif isinstance(tab, tuple):
            target = self._group_members(tab[1])
            self._overview_scope = {c['id'] for c in target}
            self.sections_layout.addWidget(self._make_group_head(tab[1]))
            if self._items_of(self._overview_scope):
                self.sections_layout.addWidget(self._make_cost_overview())
        else:
            target = [c for c in self.categories if c['id'] == tab]
        for cat in target:
            cat_items = [it for it in self.items if it['category_id'] == cat['id']]
            self.sections_layout.addWidget(self._make_section(cat, cat_items))
        self.sections_layout.addStretch()

    def _make_section(self, cat, cat_items):
        card = QFrame()
        card.setObjectName('Card')
        v = QVBoxLayout(card)
        v.setContentsMargins(1, 0, 1, 1)
        v.setSpacing(0)

        # 머리 한 줄: 이름 [합계 제외] ……… [초과 N건] [79.6%] + 🗑
        head = QWidget()
        hl = QHBoxLayout(head)
        hl.setContentsMargins(16, 10, 10, 10)
        hl.setSpacing(8)
        name = QLabel(cat['name'])
        name.setObjectName('SectionTitle')
        name.setWordWrap(True)
        meta = QLabel()                                   # 예산·집행·잔액은 아래 합계 줄에 (이 글은 숨김)
        meta.hide()
        over = QLabel()
        over.setObjectName('Pill')
        chip = QLabel()
        chip.setObjectName('Pill')
        add = _icon_btn('fa5s.plus', '이 세부항목에 항목 추가', color=C.MUTED, size=12)
        add.clicked.connect(lambda _, cid=cat['id']: self._open_add_item(cid))
        drop = _icon_btn('fa5s.trash-alt', '이 세부항목 지우기 (안의 항목과 사용 내역도 함께 지워요)',
                         color=C.FAINT, size=11)
        drop.clicked.connect(lambda _, cid=cat['id']: self._delete_category(cid))
        hl.addWidget(name)
        if cat.get('excluded'):
            tag = QLabel('합계 제외')
            tag.setObjectName('ExTag')
            tag.setToolTip('전체 합계에 넣지 않는 세부항목이에요 (왼쪽 목록에서 우클릭 → 전체 합계에 넣기)')
            hl.addWidget(tag, 0, Qt.AlignVCenter)
        hl.addStretch(1)
        hl.addWidget(over)
        hl.addWidget(chip)
        hl.addWidget(add)
        hl.addWidget(drop)
        v.addWidget(head)

        # 원가통계비목이 둘 이상이면 작은 칩 (누르면 그 비목의 줄만 보기)
        chips = {}
        costs = _cost_totals(cat_items)
        if len(costs) > 1:
            chip_box = QWidget()
            flow = _FlowLayout(chip_box, spacing=5)
            flow.setContentsMargins(16, 0, 16, 8)
            for cost in costs:
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
        main = None
        if not cat_items:
            lbl = QLabel('아직 항목이 없어요. 오른쪽 위 + 로 추가하세요.')
            lbl.setObjectName('Muted')
            lbl.setContentsMargins(16, 2, 16, 12)
            v.addWidget(lbl)
        elif active:
            main = self._make_table(active)
            tables.append(main)
            v.addWidget(main)
        else:
            lbl = QLabel('모든 항목을 완료했어요.')
            lbl.setObjectName('Muted')
            lbl.setContentsMargins(16, 2, 16, 6)
            v.addWidget(lbl)
        if done:
            # 완료한 항목은 진행 중인 항목 아래에 늘 보이게 (접기 단추 없음)
            dtable = self._make_table(done, finished=True)
            if main is not None:
                dtable.follow = main
                main.followers.append(dtable)
            else:
                main = dtable
            tables.append(dtable)
            v.addWidget(dtable)
        total = None
        if cat_items:
            total = self._make_total_table(main)
            v.addWidget(total)
        self._sec_refs[cat['id']] = {'meta': meta, 'chip': chip, 'over': over, 'chips': chips,
                                     'tables': tables, 'total': total}
        self._update_section(cat['id'])
        self._apply_cost_filter(cat['id'])
        return card

    def _make_total_table(self, follow):
        """세부항목 맨 아래 합계 줄: 예산현액 · 집행액 · 잔액을 각 칸 바로 밑에 (열 폭은 위 표를 따름)."""
        t = _BudgetTable()
        t.setColumnCount(len(_COLS))
        t.horizontalHeader().hide()
        t.verticalHeader().setVisible(False)
        t.setSelectionMode(QAbstractItemView.NoSelection)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setFocusPolicy(Qt.NoFocus)
        t.setShowGrid(False)
        t.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        t.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        hdr = t.horizontalHeader()
        for col in range(len(_COLS)):
            hdr.setSectionResizeMode(col, QHeaderView.Fixed)
        hdr.setStretchLastSection(True)
        t.setShowGrid(C.OFFICE)                          # 사무형: 격자선
        t.setStyleSheet(f'QTableWidget {{ background: {C.TOTAL_BG}; border-top: 1px solid {C.LINE}; '
                        f'border-bottom-left-radius: {C.R_CARD}px; border-bottom-right-radius: {C.R_CARD}px; }}'
                        f'QTableWidget::item {{ border: none; padding: 0 10px; }}')
        t.setRowCount(1)
        t.setRowHeight(0, px(36))
        f = QFont(t.font())
        f.setBold(True)
        for col in range(len(_COLS)):
            it = QTableWidgetItem('합계' if col == 0 else '')
            it.setFlags(Qt.ItemIsEnabled)
            it.setFont(f)
            it.setForeground(QColor(C.BODY))
            it.setTextAlignment((Qt.AlignLeft if col == 0 else Qt.AlignRight) | Qt.AlignVCenter)
            t.setItem(0, col, it)
        if follow is not None:
            t.follow = follow
            follow.followers.append(t)
        t.sync_geometry()
        return t

    # ── 원가통계비목 ───────────────────────────────────────────────
    def _toggle_cost_filter(self, cat_id, cost):
        cur = self._cost_filter.get(cat_id)
        self._cost_filter[cat_id] = None if cur == cost else cost
        self._apply_cost_filter(cat_id)
        self._schedule_save()

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
        """'전체'(또는 그룹)를 볼 때 맨 위 한 줄: 원가통계비목별 집행률 칩 ……… 색 범례."""
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(4, 0, 4, 0)
        hl.setSpacing(6)
        cap = QLabel('원가통계비목별')
        cap.setObjectName('Legend')
        if self._overview_scope is None:
            cap.setToolTip('세부항목을 가리지 않고 같은 비목끼리 합친 집행률이에요.'
                           + (' 합계 제외 세부항목은 뺐어요.' if self._excluded_ids() else ''))
        else:
            cap.setToolTip('이 그룹의 세부항목끼리 같은 비목을 합친 집행률이에요.')
        box = QWidget()
        flow = _FlowLayout(box, spacing=5)
        flow.setContentsMargins(0, 0, 0, 0)
        self._overview_chips = {}
        for cost in _cost_totals(self._overview_items()):
            lb = QLabel()
            lb.setObjectName('CostChip')
            flow.addWidget(lb)
            self._overview_chips[cost] = lb
        hl.addWidget(cap)
        hl.addWidget(box, 1)
        for color, text in ((C.OK_BAR, '진행'), (C.WARN_BAR, f'{WARN_PCT}% 이상'), (C.OVER_BAR, '초과'),
                            (C.DONE_BAR, '완료')):
            lg = QLabel(f'<span style="color:{color}">■</span> {text}')
            lg.setObjectName('Legend')
            hl.addWidget(lg)
        self._update_cost_overview()
        return row

    def _overview_items(self):
        scope = getattr(self, '_overview_scope', None)
        return self._total_items() if scope is None else self._items_of(scope)

    def _make_group_head(self, group_id):
        """그룹을 볼 때 맨 위: 그룹 이름과 묶인 세부항목을 합친 금액."""
        g = next((g for g in self.groups if g['id'] == group_id), {'name': ''})
        members = self._group_members(group_id)
        b, s = self._sum(self._items_of({c['id'] for c in members}))
        rem, pct = b - s, _pct(s, b)
        card = QFrame()
        card.setObjectName('Card')
        hl = QHBoxLayout(card)
        hl.setContentsMargins(18, 14, 18, 14)
        hl.setSpacing(10)
        tv = QVBoxLayout()
        tv.setSpacing(2)
        name = QLabel(f'{g["name"]} (그룹)')
        name.setObjectName('SectionTitle')
        names = ', '.join(c['name'] + (' (합계 제외)' if c.get('excluded') else '') for c in members)
        sub = QLabel(f'묶인 세부항목 {len(members)}개: {names}')
        sub.setObjectName('SectionMeta')
        sub.setWordWrap(True)
        tv.addWidget(name)
        tv.addWidget(sub)
        hl.addLayout(tv, 1)
        rem_html = f'<span style="color:{C.NEG}"><b>{_fmt(rem)}원</b></span>' if rem < 0 else f'{_fmt(rem)}원'
        meta = QLabel(f'예산 {_fmt(b)}원 · 집행 {_fmt(s)}원 · 잔액 {rem_html}')
        meta.setObjectName('SectionMeta')
        meta.setTextFormat(Qt.RichText)
        chip = QLabel(f'{pct:.1f}%')
        chip.setObjectName('Pill')
        chip.setStyleSheet(_pill_qss(pct))
        hl.addWidget(meta)
        hl.addWidget(chip)
        return card

    def _update_cost_overview(self):
        totals = _cost_totals(self._overview_items())
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
            h.setTextAlignment((Qt.AlignRight if 1 <= col <= _COL_PCT else Qt.AlignLeft) | Qt.AlignVCenter)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(_sizes()['row_h'])
        t.setSelectionMode(QAbstractItemView.NoSelection)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)        # 더블클릭 = 항목 수정 창
        t.setFocusPolicy(Qt.ClickFocus)
        t.setShowGrid(C.OFFICE)                          # 사무형: 칸마다 격자선
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
        t.rowMenu.connect(self._row_menu)
        t.sync_geometry()
        return t

    def _row_menu(self, item_id, pos):
        """항목 줄 오른쪽 클릭: 쓴 돈 추가 · 수정 · 완료(완료 취소) · (고른 글자 복사) · 삭제.
        산출내역 전체 복사는 비목 꼬리표 옆 복사 단추로."""
        item = self._items_by_id.get(item_id)
        if not item:
            return
        m = QMenu(self)
        m.setStyleSheet(STYLE)
        if not item['completed']:
            m.addAction(qta.icon('fa5s.plus', color=C.MUTED), '쓴 돈 추가', lambda: self._open_spend(item_id))
        m.addAction(qta.icon('fa5s.pen', color=C.MUTED), '수정', lambda: self._open_edit_item(item_id))
        if item['completed']:
            m.addAction(qta.icon('fa5s.undo', color=C.MUTED), '완료 취소', lambda: self._toggle_complete(item_id))
        else:
            m.addAction(qta.icon('fa5s.check', color=C.MUTED), '완료', lambda: self._toggle_complete(item_id))
        m.addSeparator()
        cell = (self._row_refs.get(item_id) or {}).get('desc')
        if cell is not None and cell.lbl_name.hasSelectedText():
            m.addAction(qta.icon('fa5s.copy', color=C.MUTED), '고른 글자 복사',
                        lambda: self._copy_desc(item_id, selected_only=True))
        m.addSeparator()
        m.addAction(qta.icon('fa5s.trash-alt', color=C.NEG), '삭제', lambda: self._delete_item(item_id))
        m.exec_(pos)

    def _delete_item(self, item_id):
        item = self._items_by_id.get(item_id)
        if not item:
            return
        n = len(self._expenses_by_item.get(item_id, []))
        extra = f'\n사용 내역 {n}건도 함께 지워져요.' if n else ''
        if QMessageBox.question(self, '항목 지우기', f'「{item["description"]}」 항목을 지울까요?{extra}',
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        delete_budget_item(item_id)
        self.refresh()

    def _fill_row(self, t, r, item):
        done = bool(item['completed'])
        rem = item['budget'] - item['spent']
        pct = _pct(item['spent'], item['budget'])
        R = Qt.AlignRight | Qt.AlignVCenter

        def cell(text, align, color=C.TEXT, bold=False):
            it = QTableWidgetItem(text)
            it.setFlags(Qt.ItemIsEnabled)
            it.setTextAlignment(align)
            f = QFont(t.font())
            f.setBold(bold)
            it.setFont(f)
            it.setForeground(QColor(C.FAINT if done else color))
            return it

        # 산출내역(이름 + 비목 꼬리표) — 글자는 칸 위젯(_DescCell)이 그린다
        t.setItem(r, 0, cell('', Qt.AlignLeft | Qt.AlignVCenter))
        dcell = _DescCell(t, r, item, self._expense_ui.setdefault(item['id'], {}),
                          self._expenses_by_item.get(item['id'], []), self._delete_expense, done,
                          on_edit=lambda ex, iid=item['id']: self._edit_expense(iid, ex),
                          actions={'spend': lambda iid=item['id']: self._open_spend(iid),
                                   'done': lambda iid=item['id']: self._toggle_complete(iid)})
        t.setCellWidget(r, 0, dcell)
        t.setSpan(r + 1, 0, 1, 3)                                  # 숨은 줄: 산출내역 ~ 잔액 칸을 하나로
        t.setCellWidget(r + 1, 0, dcell.detail)
        t.setItem(r, _COL_BUDGET, cell(_fmt(item['budget']), R))
        rem_it = cell(_fmt(rem), R, _bal_color(rem, pct=pct), bold=rem < 0)
        rem_it.setToolTip(f"집행액 {_fmt(item['spent'])}원")
        t.setItem(r, _COL_REM, rem_it)

        wrap = QWidget()
        pl = QHBoxLayout(wrap)
        pl.setContentsMargins(6, 0, 8, 0)
        pl.setSpacing(8)
        pl.addStretch()
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setTextVisible(False)
        bar.setFixedSize(px(60), 5)
        lbl = QLabel()
        lbl.setFixedWidth(_sizes()['pct_lbl'])
        lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        pl.addWidget(bar)
        pl.addWidget(lbl)
        t.setItem(r, _COL_PCT, cell('', R))
        t.setCellWidget(r, _COL_PCT, wrap)
        self._set_pct(bar, lbl, pct, done)
        t.pct_labels.append(lbl)

        sz = _sizes()
        for wdg in (dcell, wrap):                        # 칸 속 위젯 위에 있어도 그 줄을 '마우스 올린 줄'로
            t.watch_row(wdg)
        t.row_items[r] = item['id']
        if C.OFFICE and _level(pct, done) == 'warn':        # 사무형: 거의 다 쓴 줄은 연한 주황 바탕
            t.tint_rows[r] = C.WARN_BG
        t.setRowHeight(r, sz['row_h'])
        self._row_refs[item['id']] = {'rem': rem_it, 'bar': bar, 'pct': lbl, 'desc': dcell}
        dcell._fit_row()                                   # 펼쳐 둔 줄은 다시 펼친 높이로

    @staticmethod
    def _set_pct(bar, lbl, pct, done=False):
        bar.setValue(100 if done else int(round(min(pct, 100))))
        bar.setStyleSheet(_bar_qss(_bar_color(pct, done)))
        lbl.setText('완료' if done else f'{pct:.1f}%')
        lbl.setToolTip(f'집행률 {pct:.1f}%' if done else '')
        lbl.setStyleSheet(f'color: {_text_color(pct, done)}; font-weight: {500 if done else 700};')

    def _update_row(self, item_id):
        refs = self._row_refs.get(item_id)
        item = self._items_by_id.get(item_id)
        if not refs or not item:
            return
        rem = item['budget'] - item['spent']
        pct = _pct(item['spent'], item['budget'])
        refs['rem'].setText(_fmt(rem))
        refs['rem'].setForeground(QColor(_bal_color(rem, item['completed'], pct)))
        f = refs['rem'].font()
        f.setBold(rem < 0)
        refs['rem'].setFont(f)
        refs['rem'].setToolTip(f"집행액 {_fmt(item['spent'])}원")
        self._set_pct(refs['bar'], refs['pct'], pct, bool(item['completed']))
        refs['desc'].set_expenses(self._expenses_by_item.get(item_id, []))

    # ── 합계 ───────────────────────────────────────────────────────
    def _cat_totals(self, cat_id):
        budget = spent = 0
        for it in self.items:
            if it['category_id'] == cat_id:
                budget += it['budget']
                spent += it['spent']
        return budget, spent

    def _over_count(self, cat_ids):
        return sum(1 for it in self.items
                   if it['category_id'] in cat_ids and not it['completed'] and it['spent'] > it['budget'])

    def _update_section(self, cat_id):
        refs = self._sec_refs.get(cat_id)
        if not refs:
            return
        budget, spent = self._cat_totals(cat_id)
        rem = budget - spent
        pct = _pct(spent, budget)
        rem_html = f'<span style="color:{C.NEG}"><b>{_fmt(rem)}원</b></span>' if rem < 0 else f'{_fmt(rem)}원'
        refs['meta'].setText(f'예산 {_fmt(budget)}원 · 집행 {_fmt(spent)}원 · 잔액 {rem_html}')
        refs['chip'].setText(f'{pct:.1f}%')
        refs['chip'].setStyleSheet(_pill_qss(pct))
        n_over = self._over_count({cat_id})
        refs['over'].setText(f'초과 {n_over}건')
        refs['over'].setStyleSheet(f'background: {C.OVER_SOFT}; color: {C.OVER_TEXT};')
        refs['over'].setToolTip('예산보다 많이 쓴 항목이에요')
        refs['over'].setVisible(n_over > 0)
        total = refs.get('total')
        if total is not None:
            try:
                total.item(0, _COL_BUDGET).setText(_fmt(budget))
                total.item(0, _COL_REM).setText(_fmt(rem))
                total.item(0, _COL_REM).setForeground(QColor(_bal_color(rem)))
                total.item(0, _COL_REM).setToolTip(f'집행액 {_fmt(spent)}원')
            except RuntimeError:
                pass
        totals = _cost_totals([it for it in self.items if it['category_id'] == cat_id])
        for cost, b in refs.get('chips', {}).items():
            cb, cs, n = totals.get(cost, (0, 0, 0))
            b.setText(_cost_chip_text(cost, cb, cs))
            b.setToolTip(_cost_chip_tip(cost, cb, cs, n) + '\n\n누르면 이 비목의 줄만 보여요')
            b.setStyleSheet(_cost_chip_qss(cb, cs))

    def _update_summary(self):
        budget, spent = self._sum(self._total_items())
        rem = budget - spent
        pct = _pct(spent, budget)
        self.stat['budget'].setText(_won(budget))
        self.stat['spent'].setText(_won(spent))
        self.stat['remaining'].setText(_won(rem))
        self.stat['remaining'].setStyleSheet(f'color: {_bal_color(rem)};')
        self.lbl_pct.setText(f'{pct:.1f}%')
        self.lbl_pct.setStyleSheet(f'color: {_text_color(pct)};')
        self.bar_total.setValue(int(round(min(pct, 100))))
        self.bar_total.setStyleSheet(_bar_qss(_bar_color(pct)))
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
        tab = self.active_tab
        if isinstance(tab, tuple):                       # 그룹을 보는 중이면 그 그룹의 첫 세부항목
            tab = (self._group_members(tab[1]) or self.categories)[0]['id']
        default_cat = cat_id or (tab if tab != 'ALL' else self.categories[0]['id'])
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

    def _open_spend(self, item_id):
        """[쓴 돈 추가]: 쓴 금액을 집행액에 더하고 사용 내역에 남긴다."""
        item = self._items_by_id.get(item_id)
        if not item:
            return
        dlg = _SpendDialog(item, self._expenses_by_item.get(item_id, []), self)
        if dlg.exec_() != QDialog.Accepted or not dlg.result_data:
            return
        d = dlg.result_data
        add_budget_expense(item_id, d['amount'], d['purpose'], d['spent_on'])
        self.refresh()

    def _edit_expense(self, item_id, ex):
        """사용 내역 고치기 (✎ · 줄 더블클릭). 금액이 바뀌면 집행액도 그 차이만큼. 펼쳐 둔 줄은 그대로."""
        item = self._items_by_id.get(item_id)
        if not item:
            return
        dlg = _SpendDialog(item, (), self, expense=ex)
        if dlg.exec_() != QDialog.Accepted or not dlg.result_data:
            return
        d = dlg.result_data
        if ex.get('id') is None:                         # 「명목 없음」 → 기록으로 남기기
            record_unassigned_expense(item_id, int(ex['amount']), d['amount'], d['purpose'], d['spent_on'])
        else:
            update_budget_expense(ex['id'], d['amount'], d['purpose'], d['spent_on'])
        self.refresh()

    def _copy_desc(self, item_id, selected_only=False):
        item = self._items_by_id.get(item_id)
        if not item:
            return
        text = item['description']
        if selected_only:
            cell = (self._row_refs.get(item_id) or {}).get('desc')
            text = cell.lbl_name.selectedText() if cell is not None else ''
        if text:
            QApplication.clipboard().setText(text)

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
        # 합계 제외 · 그룹 설정은 세부항목 이름으로 기억해 두었다가 다시 붙인다 (전체 교체여도 유지)
        keep = {c['name']: (c.get('excluded') or 0, c.get('group_id')) for c in get_budget_categories()}
        if mode == 'replace':
            clear_budget_data(keep_groups=True)
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
        groups = {g['id'] for g in get_budget_groups()}
        for c in get_budget_categories():
            ex, gid = keep.get(c['name'], (0, None))
            if ex and not c.get('excluded'):
                set_budget_category_excluded(c['id'], True)
            if gid in groups and c.get('group_id') is None:
                set_budget_category_group(c['id'], gid)
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
    # ── 화면 배치 기억 (창 위치·크기 + 보던 칸·펼침·비목·스크롤) ──────────
    # 바꿀 때마다 잠깐 뒤 budget_state.json에 저장한다 (창을 닫지 않고 앱·PC를 꺼도 남게).
    def _restore_window_state(self):
        self._saved_scroll = 0
        try:
            with open(_STATE_FILE, 'r', encoding='utf-8') as f:
                s = json.load(f)
            from ui.common import fit_to_screen
            # 예전 창 크기가 새 화면(왼쪽 목록·버튼)에 좁으면 넓혀 준다
            self.setGeometry(*fit_to_screen(s['x'], s['y'], max(s['width'], px(1240)), s['height'],
                                            min_w=640, min_h=420))
            self._restore_maximized = bool(s.get('maximized'))
        except Exception:
            from ui.common import default_geometry
            self.setGeometry(*default_geometry(px(1240), px(780), frac_w=0.92))
            return
        try:
            v = s.get('view') or {}
            tab = v.get('tab', 'ALL')
            if isinstance(tab, list) and len(tab) == 2 and tab[0] == 'G':
                tab = ('G', int(tab[1]))
            elif tab != 'ALL':
                tab = int(tab)
            self.active_tab = tab
            self._show_done = {int(c) for c in v.get('show_done') or []}
            self._cost_filter = {int(c): cost for c, cost in (v.get('cost_filter') or {}).items() if cost}
            self._saved_scroll = int(v.get('scroll') or 0)
        except (TypeError, ValueError):
            self.active_tab = 'ALL'

    def _state_dict(self):
        g = self.normalGeometry() if self.isMaximized() else self.geometry()
        tab = self.active_tab
        return {'x': g.x(), 'y': g.y(), 'width': g.width(), 'height': g.height(),
                'maximized': self.isMaximized(),
                'view': {'tab': list(tab) if isinstance(tab, tuple) else tab,
                         'show_done': sorted(self._show_done),
                         'cost_filter': {str(c): cost for c, cost in self._cost_filter.items() if cost},
                         'scroll': self.scroll.verticalScrollBar().value() if hasattr(self, 'scroll') else 0}}

    def _save_state(self):
        if not getattr(self, '_state_ready', False):
            return                       # 처음 그리는 중에는 저장하지 않음 (기억한 값을 덮어쓰지 않게)
        try:
            os.makedirs(os.path.dirname(_STATE_FILE), exist_ok=True)
            tmp = _STATE_FILE + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(self._state_dict(), f, ensure_ascii=False)
            os.replace(tmp, _STATE_FILE)
        except Exception:
            pass

    def _schedule_save(self, *_):
        t = getattr(self, '_state_timer', None)
        if t is None:
            from PyQt5.QtCore import QTimer
            t = self._state_timer = QTimer(self)
            t.setSingleShot(True)
            t.setInterval(800)
            t.timeout.connect(self._save_state)
        t.start()

    def reset_layout(self):
        """「화면 배치 기본값으로」: 창 크기·위치와 보던 칸·펼침을 처음 상태로 (예산 내용은 그대로)."""
        from ui.common import default_geometry
        self.showNormal()
        self.setGeometry(*default_geometry(px(1240), px(780), frac_w=0.92))
        self.active_tab = 'ALL'
        self._show_done = set()
        self._cost_filter = {}
        self._saved_scroll = 0
        self.refresh(keep_scroll=False)
        self._save_state()

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh(keep_scroll=False)
        if not getattr(self, '_state_ready', False):
            if getattr(self, '_restore_maximized', False):
                self.showMaximized()
            pos = getattr(self, '_saved_scroll', 0)
            from PyQt5.QtCore import QTimer

            def ready():
                try:
                    self.scroll.verticalScrollBar().setValue(pos)
                    self._state_ready = True
                except RuntimeError:
                    pass
            QTimer.singleShot(0, lambda: QTimer.singleShot(0, ready))

    def moveEvent(self, event):
        super().moveEvent(event)
        self._schedule_save()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._schedule_save()

    def closeEvent(self, event):
        t = getattr(self, '_state_timer', None)
        if t is not None:
            t.stop()
        self._save_state()
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
class _RowIconButton(QPushButton):
    """사용 내역 줄의 작은 선 아이콘 단추 (연필 · 휴지통). 벡터 아이콘이라 화면 배율이 커도 선명.
    평소엔 옅은 회색, 마우스를 올리면 옅은 바탕의 둥근 네모와 진한 색 아이콘."""
    ICON = 15
    SIDE = 26

    @classmethod
    def side(cls):
        return px(_sz(cls.SIDE))

    def __init__(self, icon, hover_color, hover_bg, tip, icon_px=None, side=None):
        super().__init__()
        # 굵은 선(-bold): 가는 선은 작은 크기에서 1px 회색 선으로 흐릿하게 보였다
        bold = icon + '-bold'
        try:
            qta.icon(bold)
            icon = bold
        except Exception:
            pass
        self._icons = (qta.icon(icon, color=C.MUTED), qta.icon(icon, color=hover_color))
        self.setIcon(self._icons[0])
        # 크기는 화면 배율(px)까지 곱한다: 125%·150%에서도 글자와 같은 비율로 커져서
        # 작은 그림을 늘려 그리지 않고 그 크기로 바로 그린다
        ip = px(_sz(icon_px or self.ICON))
        self.setIconSize(QSize(ip, ip))
        sd = px(_sz(side)) if side else self.side()
        self.setFixedSize(sd, sd)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip(tip)
        self.setFlat(True)
        base = f'border: none; padding: 0; border-radius: {min(C.R_ICON, 6)}px;'
        self._qss = (f'QPushButton {{ {base} background: transparent; }}',
                     f'QPushButton {{ {base} background: {hover_bg}; }}')
        self.setStyleSheet(self._qss[0])

    def set_hot(self, on):
        """마우스를 올린 모양 (진한 아이콘 + 옅은 바탕)."""
        self.setIcon(self._icons[1 if on else 0])
        self.setStyleSheet(self._qss[1 if on else 0])

    def enterEvent(self, e):
        self.set_hot(True)
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.set_hot(False)
        super().leaveEvent(e)


def _make_selectable(label):
    """글자를 마우스로 드래그해서 고르고 Ctrl+C로 복사할 수 있게. 우클릭은 표의 줄 메뉴로 넘긴다."""
    label.setTextInteractionFlags(Qt.TextSelectableByMouse)
    label.setContextMenuPolicy(Qt.NoContextMenu)
    return label


class _SelectableName(QLabel):
    """산출내역 이름: 드래그하면 글자 고르기, 그냥 한 번 누르면 clicked (펼치기)."""
    clicked = pyqtSignal()

    def __init__(self, text=''):
        super().__init__(text)
        _make_selectable(self)
        self._press = None

    def mousePressEvent(self, e):
        self._press = e.pos() if e.button() == Qt.LeftButton else None
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e):
        # 더블클릭: 이름 전체를 고른다 (Ctrl+C로 바로 복사)
        if e.button() == Qt.LeftButton:
            self._press = None
            self._dbl = True
            self.setSelection(0, len(self.text()))
            self.setFocus(Qt.MouseFocusReason)
            return
        super().mouseDoubleClickEvent(e)

    def mouseReleaseEvent(self, e):
        if getattr(self, '_dbl', False):          # 더블클릭 직후 손을 뗄 때는 펼치기 · 선택 해제 없음
            self._dbl = False
            return
        super().mouseReleaseEvent(e)
        start, self._press = self._press, None
        if (start is not None and e.button() == Qt.LeftButton and not self.hasSelectedText()
                and (e.pos() - start).manhattanLength() < QApplication.startDragDistance()):
            self.clicked.emit()


class _DescCell(QWidget):
    """산출내역 칸: 이름 한 줄 + 사용 내역 요약. 누르면 명목별 사용 내역이 펼쳐진다.
    펼침·더 보기 상태는 state(창을 여는 동안만 유지)에 둔다."""
    SHOW = 5                                       # 펼쳤을 때 먼저 보여 줄 건수

    def __init__(self, table, row, item, state, expenses, on_delete, done=False, on_edit=None, actions=None):
        super().__init__()
        self._table, self._row, self._item = table, row, item
        self._state, self._on_delete, self._done = state, on_delete, done
        self._on_edit = on_edit                    # 사용 내역 고치기 (기록 dict, 명목 없음이면 id None)
        self.setStyleSheet('background: transparent;')
        v = QVBoxLayout(self)
        v.setContentsMargins(6, 0, 6, 0)
        v.setSpacing(2)

        head = QVBoxLayout()
        head.setSpacing(0)
        name_row = QHBoxLayout()
        name_row.setSpacing(6)
        # ▸ 화살표: 펼치기/접기 · 이름: 드래그해서 고를 수 있는 글자 (그냥 한 번 누르면 펼치기)
        self.btn_toggle = QPushButton()
        self.btn_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_toggle.setFlat(True)
        self.btn_toggle.setFixedSize(px(_sz(16)), px(_sz(20)))
        self.btn_toggle.setToolTip('사용 내역 펼치기/접기')
        self.btn_toggle.setStyleSheet('QPushButton { border: none; background: transparent; padding: 0; }')
        self.btn_toggle.clicked.connect(self._toggle)
        self.lbl_name = _SelectableName(item['description'])
        self.lbl_name.clicked.connect(self._toggle)
        self.lbl_name.setToolTip(f"{item['description']}\n드래그 · 더블클릭으로 고른 뒤 Ctrl+C로 복사\n우클릭: 쓴 돈 추가 · 수정 · 완료 · 삭제")
        self.lbl_name.setStyleSheet(
            f'QLabel {{ background: transparent; color: {C.FAINT if done else C.TEXT}; font-weight: 600; }} '
            f'QLabel:hover {{ color: {C.ACCENT}; }}')
        name_font = self.lbl_name.font()
        name_font.setStrikeOut(done)
        self.lbl_name.setFont(name_font)
        self.lbl_sum = QLabel()
        self.lbl_sum.setStyleSheet(f'color: {C.MUTED}; font-size: {_pt(10)}pt; background: transparent;')
        name_row.setSpacing(2)
        name_row.addWidget(self.btn_toggle)
        name_row.addWidget(self.lbl_name)
        name_row.addSpacing(4)
        self.lbl_cost = QLabel(item.get('cost_type') or '')       # 원가통계비목 꼬리표 (운영비 · 재료비 …)
        self.lbl_cost.setObjectName('CostTag')
        self.lbl_cost.setStyleSheet(f'background: {C.TAG_BG}; color: {C.MUTED}; border-radius: 4px; '
                                    f'padding: 1px 6px; font-size: {_pt(10)}pt;')   # 칸의 투명 바탕보다 먼저
        name_row.addWidget(self.lbl_cost)                          # 칸에 넣은 뒤에 숨겨야 따로 창으로 뜨지 않음
        if not item.get('cost_type'):
            self.lbl_cost.hide()
        # 비목 꼬리표 오른쪽: [+ 쓴 돈 추가] [✓ 완료] (완료한 항목은 [↺ 완료 취소]만)
        # 수정 · 삭제는 우클릭 메뉴, 복사는 이름을 드래그(또는 더블클릭)한 뒤 Ctrl+C
        actions = actions or {}
        self.row_buttons = {}
        name_row.addSpacing(2)
        if done:
            specs = (('done', 'ph.arrow-counter-clockwise', C.ACCENT, C.ACCENT_SOFT, '완료 취소'),)
        else:
            specs = (('spend', 'ph.plus', C.ACCENT, C.ACCENT_SOFT, '쓴 돈 추가'),
                     ('done', 'ph.check', '#047857', '#E7F6EE', '완료'))
        for key, icon, color, bg, tip in specs:
            if key not in actions:
                continue
            b = _RowIconButton(icon, color, bg, tip, icon_px=15, side=24)
            b.clicked.connect(lambda _, f=actions[key]: f())
            name_row.addWidget(b)
            self.row_buttons[key] = b
        name_row.addStretch(1)
        head.addLayout(name_row)
        head.addWidget(self.lbl_sum)
        v.addStretch(1)                            # 이름을 칸 세로 가운데에 (다른 칸 글자와 줄 맞춤)
        v.addLayout(head)
        v.addStretch(1)

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
            self.lbl_sum.setStyleSheet(f'color: {C.NEG}; font-size: {_pt(10)}pt; background: transparent;')
        else:
            self.lbl_sum.setText('')
        self.lbl_sum.setVisible(unassigned < 0)    # 빈 줄이 자리를 차지하면 이름이 위로 밀린다
        self.lbl_name.setText(self._item['description'])
        self.btn_toggle.setIcon(qta.icon('ph.caret-down-bold' if self._open() else 'ph.caret-right-bold', color=C.MUTED))
        ip = px(_sz(12))                          # 화면 배율에 맞는 크기로 바로 그려 선명하게
        self.btn_toggle.setIconSize(QSize(ip, ip))
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
            self._dv.addLayout(self._entry_row('', '명목 없음', unassigned,
                                               ex={'id': None, 'amount': unassigned, 'purpose': '', 'spent_on': ''}))
        more = bool(self._state.get('more'))
        shown = self._expenses if more else self._expenses[:self.SHOW]
        for ex in shown:
            self._dv.addLayout(self._entry_row(ex['spent_on'] or '-', ex['purpose'], ex['amount'], ex['id'], ex))
        if len(self._expenses) > self.SHOW:
            left = len(self._expenses) - self.SHOW
            btn = QPushButton('접기' if more else f'더 보기 ({left}건 더)')
            btn.setFlat(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f'QPushButton {{ text-align: left; border: none; background: transparent; '
                              f'padding: 0; color: {C.ACCENT_HI}; font-size: {_pt(10)}pt; }}')
            btn.clicked.connect(self._toggle_more)
            self._dv.addWidget(btn)
        if unassigned < 0:
            warn = QLabel('⚠ 기록 합계가 집행액보다 큽니다. 집행액을 직접 낮췄다면 기록을 확인하세요.')
            warn.setWordWrap(True)
            warn.setStyleSheet(f'color: {C.NEG}; font-size: {_pt(10)}pt; background: transparent;')
            self._dv.addWidget(warn)

    def _entry_row(self, date, purpose, amount, eid=None, ex=None):
        """사용 내역 한 줄: 날짜 · 쓴 곳 · 금액 · [연필] [휴지통]. 글자는 드래그해서 복사, 줄을 더블클릭하면 고치기."""
        row = QHBoxLayout()
        row.setSpacing(8)
        d = QLabel(date)
        d.setFixedWidth(_sz(88))
        d.setStyleSheet(f'color: {C.MUTED}; font-size: {_pt(10)}pt; background: transparent;')
        p = QLabel(purpose or ('(쓴 곳 미입력)' if eid is not None else purpose))
        p.setStyleSheet(f'color: {C.TEXT if purpose or eid is None else C.FAINT}; font-size: {_pt(10)}pt; background: transparent;')
        a = QLabel(f"{_fmt(amount)}원")
        a.setStyleSheet(f'color: {C.TEXT}; font-size: {_pt(10)}pt; font-weight: 700; background: transparent;')
        a.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        for lb in (d, p, a):
            _make_selectable(lb)
        row.addWidget(d)
        row.addWidget(p, 1)
        row.addWidget(a)
        btns = QHBoxLayout()                          # 아이콘 단추 둘은 조금 붙여서
        btns.setSpacing(px(4))
        row.addLayout(btns)
        editable = ex is not None and self._on_edit is not None
        if editable:
            for lb in (d, p, a):
                lb.installEventFilter(self)
                lb.setProperty('expense', ex)
            e = _RowIconButton('ph.pencil-simple', C.ACCENT, C.ACCENT_SOFT,
                               '쓴 곳 · 날짜 적기 (기록으로 남겨요)' if eid is None else '이 사용 내역 고치기')
            e.clicked.connect(lambda _, ex=ex: self._on_edit(ex))
            btns.addWidget(e)
        else:
            spacer = QWidget()
            spacer.setFixedSize(_RowIconButton.side(), _RowIconButton.side())
            btns.addWidget(spacer)
        if eid is None:                              # 명목 없음: 지우기 자리는 비워 아이콘 줄을 맞춤
            spacer = QWidget()
            spacer.setFixedSize(_RowIconButton.side(), _RowIconButton.side())
            btns.addWidget(spacer)
            return row
        x = _RowIconButton('ph.trash', C.NEG, C.OVER_SOFT, '이 사용 내역 지우기 (집행액도 그만큼 줄어요)')
        x.clicked.connect(lambda _, eid=eid: self._on_delete(eid))
        btns.addWidget(x)
        return row

    def eventFilter(self, obj, e):
        # 사용 내역 줄 글자를 더블클릭하면 고치기 (낱말 고르기 대신)
        if e.type() == QEvent.MouseButtonDblClick and e.button() == Qt.LeftButton:
            ex = obj.property('expense')
            if ex is not None and self._on_edit is not None:
                self._on_edit(ex)
                return True
        return super().eventFilter(obj, e)

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
            span = sum(t.columnWidth(c) for c in range(0, _COL_PCT))   # 산출내역 ~ 잔액 칸 폭
            h = self.detail.heightForWidth(span)
            if h < 0:
                h = self.detail.sizeHint().height()
            t.setRowHeight(r + 1, h + px(8))
        t.sync_geometry()


class _SpendDialog(QDialog):
    """[쓴 돈 추가]: 이번에 쓴 금액 · 쓴 곳 · 날짜만 적는 작은 창. 넣으면 집행액에 더하고 사용 내역에 남긴다.
    expense를 주면 「사용 내역 수정」: 그 기록 값을 채워 열고, 금액 차이만큼 집행액이 바뀐다.
    (expense['id']가 None이면 「명목 없음」 금액을 기록으로 남기기)"""
    RECENT = 3

    def __init__(self, item, expenses=(), parent=None, expense=None):
        super().__init__(parent)
        from datetime import date
        self._expense = expense
        self._old = int(expense['amount']) if expense else 0      # 고치기 전 금액 (이미 집행액에 들어 있음)
        title_text = '쓴 돈 추가' if expense is None else '사용 내역 수정'
        self.setWindowTitle(title_text)
        self.setStyleSheet(STYLE)
        self.setMinimumWidth(px(_sz(400)))
        self._item = item
        self.result_data = None

        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(6)
        title = QLabel(title_text)
        title.setObjectName('SectionTitle')
        sub = QLabel(f"{item['description']} · {item['cost_type'] or '일반'}")
        sub.setObjectName('SectionMeta')
        sub.setWordWrap(True)
        v.addWidget(title)
        v.addWidget(sub)
        v.addSpacing(6)

        v.addWidget(_ItemDialog._label('이번에 쓴 금액' if expense is None else '쓴 금액'))
        self.edit_spend_now = QLineEdit()                 # 왼쪽 정렬 (다른 칸과 같은 방향, 커서가 왼쪽에서 시작)
        self.edit_spend_now.setPlaceholderText('0')
        self.edit_spend_now.setStyleSheet(
            f'QLineEdit {{ font-size: {_pt(16)}pt; font-weight: 700; padding: 8px 12px; border: 1px solid {C.ACCENT_LINE}; }}'
            f'QLineEdit:focus {{ border: 1px solid {C.ACCENT}; }}')
        _attach_amount_format(self.edit_spend_now)
        v.addWidget(self.edit_spend_now)

        v.addSpacing(4)
        v.addWidget(_ItemDialog._label('쓴 곳 · 날짜'))
        where = QHBoxLayout()
        where.setSpacing(8)
        self.edit_purpose = QLineEdit()
        self.edit_purpose.setPlaceholderText('쓴 곳 (예: 11월 방과후 강사료)')
        self.edit_spend_date = QLineEdit(date.today().isoformat())
        self.edit_spend_date.setFixedWidth(px(_sz(120)))
        self.edit_spend_date.setToolTip('사용한 날짜 (예: 2026-11-10)')
        where.addWidget(self.edit_purpose, 1)
        where.addWidget(self.edit_spend_date)
        v.addLayout(where)

        self.lbl_spend_preview = QLabel()
        self.lbl_spend_preview.setTextFormat(Qt.RichText)
        self.lbl_spend_preview.setWordWrap(False)            # 한 줄 (잔액만)
        self.lbl_spend_preview.setStyleSheet(f'background: {C.SURFACE}; border: 1px solid {C.BORDER}; '
                                             f'border-radius: 8px; padding: 9px 12px; color: {C.MUTED}; font-size: {_pt(10)}pt;')
        self.lbl_spend_preview.ensurePolished()
        _fm = self.lbl_spend_preview.fontMetrics()
        self.lbl_spend_preview.setMinimumHeight(_fm.lineSpacing() + px(24))
        v.addSpacing(6)
        v.addWidget(self.lbl_spend_preview)

        if expense is not None:                         # 고치기: 지금 값을 채워 둔다
            self.edit_spend_now.setText(_fmt(self._old))
            self.edit_purpose.setText(expense.get('purpose') or '')
            if expense.get('spent_on'):
                self.edit_spend_date.setText(expense['spent_on'])
        recent = [] if expense is not None else list(expenses)[:self.RECENT]   # 최근 쓴 돈 (참고용)
        if recent:
            v.addSpacing(4)
            cap = QLabel('최근 쓴 돈')
            cap.setStyleSheet(f'color: {C.FAINT}; font-size: {_pt(10)}pt; font-weight: 600;')
            v.addWidget(cap)
            for ex in recent:
                row = QHBoxLayout()
                d = QLabel(ex.get('spent_on') or '-')
                d.setFixedWidth(px(_sz(88)))
                d.setStyleSheet(f'color: {C.FAINT}; font-size: {_pt(10)}pt;')
                pu = QLabel(ex.get('purpose') or '(쓴 곳 미입력)')
                pu.setStyleSheet(f'color: {C.BODY}; font-size: {_pt(10)}pt;')
                am = QLabel(f"{_fmt(ex['amount'])}원")
                am.setStyleSheet(f'color: {C.BODY}; font-size: {_pt(10)}pt; font-weight: 600;')
                row.addWidget(d)
                row.addWidget(pu, 1)
                row.addWidget(am)
                v.addLayout(row)

        btns = QHBoxLayout()
        btns.addStretch()
        self.btn_save = QPushButton('넣기' if expense is None else '저장')
        self.btn_save.setObjectName('Primary')
        self.btn_save.setDefault(True)                    # Enter → 넣기
        self.btn_save.clicked.connect(self._save)
        self.btn_cancel = QPushButton('취소')
        self.btn_cancel.setAutoDefault(False)
        self.btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_cancel)
        v.addSpacing(8)
        v.addLayout(btns)

        self.edit_spend_now.textChanged.connect(self._update_spend_preview)
        self._update_spend_preview()
        self.edit_spend_now.setFocus()

    def _update_spend_preview(self):
        """잔액만: 금액을 적기 전에는 지금 잔액, 적으면 '지금 → 바뀔 잔액'."""
        budget = self._item['budget']
        base = self._item['spent'] - self._old          # 고치기면 원래 금액을 뺀 집행액에서 시작
        add = _parse_amount(self.edit_spend_now.text())
        rem_before = budget - self._item['spent']
        if self._expense is not None:                   # 고치기: 지금 잔액 → 바뀔 잔액 (같으면 지금 잔액만)
            rem_after = budget - base - add
            color = _bal_color(rem_after, pct=_pct(base + add, budget))
            if add <= 0 or rem_after == rem_before:
                self.lbl_spend_preview.setText(f'잔액 <b style="color:{color}">{_fmt(rem_before)}원</b>')
            else:
                self.lbl_spend_preview.setText(f'잔액 {_fmt(rem_before)} → <b style="color:{color}">{_fmt(rem_after)}원</b>')
            return
        if add <= 0:
            self.lbl_spend_preview.setText(
                f'잔액 <b style="color:{_bal_color(rem_before, pct=_pct(base, budget))}">{_fmt(rem_before)}원</b>')
            return
        rem_after = budget - base - add
        color = _bal_color(rem_after, pct=_pct(base + add, budget))
        self.lbl_spend_preview.setText(f'잔액 {_fmt(rem_before)} → <b style="color:{color}">{_fmt(rem_after)}원</b>')

    def _save(self):
        from datetime import datetime
        amount = _parse_amount(self.edit_spend_now.text())
        if amount <= 0:
            QMessageBox.warning(self, '넣기 실패', '쓴 금액을 적어 주세요.'
                                + ('\n기록을 없애려면 줄 끝 ✕를 누르세요.' if self._expense is not None else ''))
            self.edit_spend_now.setFocus()
            return
        spent_on = self.edit_spend_date.text().strip()
        try:
            datetime.strptime(spent_on, '%Y-%m-%d')
        except ValueError:
            QMessageBox.warning(self, '넣기 실패', '날짜를 2026-11-10처럼 적어 주세요.')
            return
        self.result_data = {'amount': amount, 'purpose': self.edit_purpose.text().strip(), 'spent_on': spent_on}
        self.accept()


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
        # 목록에서 고르기만 (직접 입력 없음): (비목 없음) · 자주 쓰는 비목 ─ 나머지 · 지금 쓰는 비목 ─ 목록 관리…
        self.edit_cost = QComboBox()
        self.edit_cost.setMaxVisibleItems(18)
        self._fill_cost_combo((item or {}).get('cost_type', ''))
        self.edit_cost.activated.connect(self._on_cost_picked)
        v.addWidget(self.edit_cost)

        v.addWidget(self._label('산출내역 (내용)'))
        self.edit_desc = QLineEdit()
        self.edit_desc.setPlaceholderText('예: 교과융합캠프 간식비')
        v.addWidget(self.edit_desc)

        # 쓴 돈은 줄의 [쓴 돈 추가] 창에서 따로 넣는다 (이 창은 항목 정보 고치기만)
        self.edit_spend_now = None
        self.edit_purpose = None
        self.edit_spend_date = None

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
        spent_col.addWidget(self._label('집행액 (바로잡기)' if item else '집행액 (지출액)'))
        self.edit_spent = QLineEdit('0')
        self.edit_spent.setAlignment(Qt.AlignRight)
        _attach_amount_format(self.edit_spent)
        spent_col.addWidget(self.edit_spent)
        amount_row.addLayout(budget_col)
        amount_row.addLayout(spent_col)
        v.addLayout(amount_row)
        if item:
            note = QLabel('집행액은 잘못 넣은 값을 바로잡을 때만 고쳐요. 새로 쓴 돈은 줄의 [쓴 돈 추가]로 넣으면 '
                          '사용 내역에 남아요.')
            note.setWordWrap(True)
            note.setStyleSheet(f'color: {C.FAINT}; font-size: {_pt(10)}pt;')
            note.ensurePolished()
            note.setMinimumHeight(note.heightForWidth(px(380)))   # 줄바꿈 높이를 미리 잡아 금액 칸이 눌리지 않게
            v.addWidget(note)

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
            self.edit_desc.setText(item['description'])
            self.edit_budget.setText(f'{item["budget"]:,}')
            self.edit_spent.setText(f'{item["spent"]:,}')
        self.edit_desc.setFocus()

    def _final_spent(self):
        return _parse_amount(self.edit_spent.text())

    def _delete(self):
        ans = QMessageBox.question(self, '항목 삭제', f"'{self._item['description']}' 항목을 삭제할까요?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans == QMessageBox.Yes:
            self.deleted = True
            self.accept()

    _MANAGE = '__manage__'

    def _fill_cost_combo(self, current=''):
        """원가통계비목 고르는 칸 채우기. current가 목록에 없어도(예전에 가져온 비목) 고를 수 있게 넣는다."""
        import budget_types as bt
        cb = self.edit_cost
        cb.blockSignals(True)
        cb.clear()
        cb.addItem('(비목 없음)', '')
        saved = bt.load_saved()
        keys = {n.replace(' ', '') for n in saved}
        pri = [n for n in saved if n in bt.PRIORITY]
        rest = [n for n in saved if n not in bt.PRIORITY]
        extra = [n for n in bt.in_use() if n.replace(' ', '') not in keys]
        cur = ' '.join(str(current or '').split())
        if cur and cur != '일반' and cur.replace(' ', '') not in keys and cur not in extra:
            extra.append(cur)
        for n in pri:
            cb.addItem(n, n)
        if pri and (rest or extra):
            cb.insertSeparator(cb.count())
        for n in rest + extra:
            cb.addItem(n, n)
        cb.insertSeparator(cb.count())
        cb.addItem('목록 관리…', self._MANAGE)
        idx = cb.findData(cur) if cur and cur != '일반' else 0
        if idx < 0:
            idx = next((i for i in range(cb.count())
                        if str(cb.itemData(i) or '').replace(' ', '') == cur.replace(' ', '')), 0)
        cb.setCurrentIndex(idx)
        self._cost_last = cb.currentIndex()
        cb.blockSignals(False)

    def _on_cost_picked(self, idx):
        if self.edit_cost.itemData(idx) != self._MANAGE:
            self._cost_last = idx
            return
        keep = self.edit_cost.itemData(self._cost_last) or ''
        _CostTypeDialog(self).exec_()                # 목록에 더하거나 뺀 뒤 다시 채움
        self._fill_cost_combo(keep)

    def cost_value(self):
        return self.edit_cost.currentData() or ''

    @staticmethod
    def _label(text):
        lbl = QLabel(text)
        lbl.setStyleSheet(f'font-size: {_pt(10)}pt; font-weight: 700; color: {C.MUTED};')
        return lbl

    def _save(self):
        desc = self.edit_desc.text().strip()
        if not desc:
            QMessageBox.warning(self, '저장 실패', '산출내역(내용)을 입력해주세요.')
            return
        cost = self.cost_value() or '일반'
        expense = None
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
        self.setMinimumWidth(620)
        self.setMinimumHeight(560)
        self._on_changed = on_changed

        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        lbl = QLabel('추경 발생 시 세부항목을 새로 추가하거나 명칭 변경, 삭제가 가능합니다.\n'
                     '「합계 포함」을 끄면 전체 합계에서 빠지고, 「그룹」으로 비슷한 세부항목을 묶어 합친 금액을 볼 수 있어요.')
        lbl.setWordWrap(True)
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

        # 그룹: 만들기 · 이름 바꾸기 · 풀기
        gcap = QLabel('그룹')
        gcap.setObjectName('StatLabel')
        v.addWidget(gcap)
        grow = QHBoxLayout()
        grow.setSpacing(8)
        self.edit_group = QLineEdit()
        self.edit_group.setPlaceholderText('새 그룹 이름 (예: 방과후)')
        self.edit_group.returnPressed.connect(self._add_group)
        btn_gadd = QPushButton('그룹 추가')
        btn_gadd.clicked.connect(self._add_group)
        grow.addWidget(self.edit_group, 1)
        grow.addWidget(btn_gadd)
        v.addLayout(grow)
        self.group_box = QWidget()
        self.group_layout = QVBoxLayout(self.group_box)
        self.group_layout.setContentsMargins(0, 0, 0, 0)
        self.group_layout.setSpacing(4)
        v.addWidget(self.group_box)

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
                w.hide()                  # 숨긴 뒤 떼어야 지우기 전 잠깐 따로 창으로 뜨지 않음
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
        groups = get_budget_groups()
        for cat in cats:
            row = QWidget()
            rl = QHBoxLayout(row)
            rl.setContentsMargins(2, 2, 2, 2)
            rl.setSpacing(6)
            edit = QLineEdit(cat['name'])
            edit.editingFinished.connect(
                lambda cid=cat['id'], e=edit, old=cat['name']: self._rename_category(cid, e, old))
            combo = QComboBox()
            combo.setToolTip('이 세부항목을 넣을 그룹')
            combo.addItem('그룹 없음', None)
            for g in groups:
                combo.addItem(g['name'], g['id'])
            idx = combo.findData(cat.get('group_id'))
            combo.setCurrentIndex(idx if idx >= 0 else 0)
            combo.setMinimumWidth(110)
            combo.currentIndexChanged.connect(
                lambda _, cid=cat['id'], cb=combo: self._set_group(cid, cb.currentData()))
            chk = QCheckBox('합계 포함')
            chk.setToolTip('끄면 전체 합계(예산현액·집행액·잔액·집행률)에서 빠져요')
            chk.setChecked(not cat.get('excluded'))
            chk.toggled.connect(lambda on, cid=cat['id']: self._set_included(cid, on))
            btn_del = _icon_btn('fa5s.trash-alt', '세부항목 삭제', color='#E11D48', size=13)
            btn_del.clicked.connect(lambda _, cid=cat['id']: self._delete_category(cid))
            rl.addWidget(edit, stretch=1)
            rl.addWidget(combo)
            rl.addWidget(chk)
            rl.addWidget(btn_del)
            self.list_layout.addWidget(row)
        self.list_layout.addStretch()
        self._rebuild_groups(groups, cats)

    def _rebuild_groups(self, groups, cats):
        while self.group_layout.count():
            w = self.group_layout.takeAt(0).widget()
            if w:
                w.hide()                  # 숨긴 뒤 떼어야 지우기 전 잠깐 따로 창으로 뜨지 않음
                w.setParent(None)
                w.deleteLater()
        if not groups:
            hint = QLabel('아직 그룹이 없어요. 위에서 만들거나, 왼쪽 목록에서 세부항목을 우클릭 → 그룹으로 옮기기.')
            hint.setObjectName('Muted')
            hint.setWordWrap(True)
            self.group_layout.addWidget(hint)
        for g in groups:
            row = QWidget()
            rl = QHBoxLayout(row)
            rl.setContentsMargins(2, 0, 2, 0)
            rl.setSpacing(6)
            edit = QLineEdit(g['name'])
            edit.editingFinished.connect(
                lambda gid=g['id'], e=edit, old=g['name']: self._rename_group(gid, e, old))
            n = sum(1 for c in cats if c.get('group_id') == g['id'])
            cnt = QLabel(f'세부항목 {n}개')
            cnt.setObjectName('Muted')
            btn_del = _icon_btn('fa5s.times', '그룹 풀기 (세부항목은 그대로 남아요)', color=C.MUTED, size=13)
            btn_del.clicked.connect(lambda _, gid=g['id']: self._delete_group(gid))
            rl.addWidget(edit, 1)
            rl.addWidget(cnt)
            rl.addWidget(btn_del)
            self.group_layout.addWidget(row)

    def _set_group(self, cat_id, group_id):
        set_budget_category_group(cat_id, group_id)
        self._on_changed()

    def _set_included(self, cat_id, on):
        set_budget_category_excluded(cat_id, not on)
        self._on_changed()

    def _add_group(self):
        name = self.edit_group.text().strip()
        if not name:
            return
        if add_budget_group(name) is None:
            QMessageBox.warning(self, '추가 실패', '같은 이름의 그룹이 이미 있어요.')
            return
        self.edit_group.clear()
        self._rebuild_list()
        self._on_changed()

    def _rename_group(self, group_id, edit, old_name):
        new_name = edit.text().strip()
        if not new_name or new_name == old_name:
            edit.setText(old_name)
            return
        if not rename_budget_group(group_id, new_name):
            QMessageBox.warning(self, '변경 실패', '같은 이름의 그룹이 이미 있어요.')
            edit.setText(old_name)
            return
        self._rebuild_list()
        self._on_changed()

    def _delete_group(self, group_id):
        delete_budget_group(group_id)
        self._rebuild_list()
        self._on_changed()

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
                w.hide()                  # 숨긴 뒤 떼어야 지우기 전 잠깐 따로 창으로 뜨지 않음
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
