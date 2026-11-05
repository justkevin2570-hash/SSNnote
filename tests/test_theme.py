import colorsys

import pytest

import ui.theme as theme


def _rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _lum(h):
    def ch(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(x) for x in _rgb(h))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a, b):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def lightness(h):
    return colorsys.rgb_to_hls(*(x / 255 for x in _rgb(h)))[1]


def test_light_mode_is_identity(monkeypatch):
    monkeypatch.setattr(theme, 'DARK', False)
    for v in ['#333', 'white', 'rgba(0,0,0,0.12)', '#FDCFE8']:
        assert theme.tc(v) == v and theme.tct(v) == v and theme.note_color(v) == v


def test_dark_color_rules():
    assert lightness(theme.dark_color('white')) < 0.2          # 흰 바탕 → 어둡게
    assert lightness(theme.dark_color('#333')) > 0.7           # 진한 글자 → 밝게
    assert lightness(theme.dark_color('#5a4000')) > 0.7
    assert theme.dark_color('#4a90d9') == '#4a90d9'            # 중간 밝기 강조색은 유지
    assert theme.dark_color('rgba(0,0,0,0.12)') == 'rgba(255,255,255,0.12)'
    assert theme.dark_color('rgba(255,255,255,0.5)') == 'rgba(0,0,0,0.25)'
    assert theme.dark_color('transparent') == 'transparent'    # 못 읽는 값은 그대로


@pytest.mark.parametrize('pastel', ['#e8b5a8', '#a8b8e8', '#ccc', '#ddd'])
def test_pastel_text_stays_visible(pastel):
    out = theme.dark_text_color(pastel)
    assert 0.38 <= lightness(out) <= 0.7                       # 사라지지도, 튀지도 않게


def test_note_colors_are_deep_and_distinct(monkeypatch):
    monkeypatch.setattr(theme, 'DARK', True)
    shown = [theme.note_color(c) for c in theme.PALETTE]
    assert all(lightness(c) < 0.32 for c in shown)
    assert len(set(shown)) == len(shown)


def test_dark_tokens_meet_contrast():
    d = theme._DARK
    for fg in ('INK', 'TEXT_WARM', 'TEXT_ON_NOTE'):
        for bg in ('BG_WARM', 'SURFACE', 'SURFACE_WARM'):
            assert contrast(d[fg], d[bg]) >= 4.5, (fg, bg, contrast(d[fg], d[bg]))
    assert contrast(d['TEXT_WARM_MUTED'], d['BG_WARM']) >= 4.5
    # 어두운 메모(노랑) 위 기본 글자
    note = '#4b4b20'
    assert contrast(d['TEXT_ON_NOTE'], note) >= 4.5


def test_light_tokens_unchanged_in_TL():
    assert theme.TL.ACCENT == '#d4b800' and theme.TL.BG_WARM == '#faf6f0'
