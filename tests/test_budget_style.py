"""예산 창 집행률 색 기준 (수준: 정상 · 80% 이상 · 100% 이상 · 완료)."""
import pytest

pytest.importorskip('PyQt5')


@pytest.mark.parametrize('pct, done, expected', [
    (0, False, 'ok'), (79.9, False, 'ok'), (80, False, 'warn'),
    (99.9, False, 'warn'), (100, False, 'over'), (130, False, 'over'),
    (120, True, 'done'),
])
def test_level(pct, done, expected):
    from budget import _level
    assert _level(pct, done) == expected


# ── 선명하게 보이는 기준: 작은 글자·연한 글자 없애기 ──

def _contrast(a, b='#FFFFFF'):
    def lum(h):
        h = h.lstrip('#')
        r, g, bl = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(bl)
    hi, lo = sorted((lum(a), lum(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def test_secondary_text_colors_are_readable():
    from budget import C
    assert _contrast(C.MUTED) >= 6.0          # 보조 글 (예전 4.8)
    assert _contrast(C.FAINT) >= 4.5          # 가장 연한 글도 기준 이상 (예전 #94A3B8은 2.6 → 흐려 보였음)
    assert _contrast(C.MUTED, C.BG) >= 5.5


def test_no_text_smaller_than_10pt():
    import re
    src = open(__import__('budget').__file__, encoding='utf-8').read()
    sizes = [float(x) for x in re.findall(r'font-size:\s*([\d.]+)pt', src)]
    assert sizes and min(sizes) >= 10.0       # 힌팅 없는 Pretendard는 12px 안팎에서 가장 흐려 보인다


def test_budget_window_uses_malgun_gothic_first():
    import budget
    assert budget._UI_FONTS[0] == 'Malgun Gothic' and 'Pretendard' in budget._UI_FONTS     # 없는 PC는 Pretendard로
    assert "font-family: 'Malgun Gothic', 'Pretendard'" in budget.STYLE


def test_budget_widgets_get_the_font(tmp_path, monkeypatch):
    import os
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    import db
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])
    import budget
    monkeypatch.setattr(budget, '_STATE_FILE', str(tmp_path / 'state.json'))
    w = budget.BudgetWindow()
    assert 'Malgun Gothic' in w.styleSheet()
    w.close()
