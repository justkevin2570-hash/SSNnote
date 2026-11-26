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


def test_memo_window_min_width_follows_text_size(tmp_path, monkeypatch):
    import db, window
    from ui.chrome import EdgeHandle
    from ui.common import px
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    monkeypatch.setenv('APPDATA', str(tmp_path))
    db.init_db()
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    try:
        w.apply_state(100, 100, 320, 400, 0, '', 1.0, 0)
        base = px(EdgeHandle.MIN_W)              # Windows 배율만큼 (테스트 화면은 ~100dpi)
        assert w.minimumWidth() == base
        w._set_scale(1.5)
        assert w.minimumWidth() == round(base * 1.5) or abs(w.minimumWidth() - base * 1.5) <= 1
        assert w.width() >= w.minimumWidth()
        w._set_scale(1.0)
        assert w.minimumWidth() == base   # 줄이면 최소 폭도 돌아온다 (창 폭은 그대로)
    finally:
        w._set_scale(1.0)
        w.close()
