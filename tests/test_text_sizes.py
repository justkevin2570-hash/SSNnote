"""글자 크기 5단계 (기본 · 크게 · 더 크게 · 더더 크게 · 아주 크게)."""
import ai_actions
from ui import style_menu


def test_memo_sizes_five_steps():
    assert [l for l, _ in style_menu.SIZES] == ['기본', '크게', '더 크게', '더더 크게', '아주 크게']
    vals = [v for _, v in style_menu.SIZES]
    assert vals == sorted(vals) and vals[0] == 1.0 and vals[-1] == 1.5


def test_ai_accepts_new_sizes():
    assert ai_actions._size('더더 크게') == 1.35
    assert ai_actions._size('아주 크게') == 1.5
    assert dict(style_menu.SIZES) == ai_actions.SIZE_VALUES
