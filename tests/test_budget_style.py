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
