"""업무창 글자 크기를 따라 메뉴 글자도 커진다 (하위 메뉴 · 업무 ⋮ 메뉴 포함)."""
import pytest

pytest.importorskip('PyQt5')


@pytest.fixture(scope='module')
def app():
    from PyQt5.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def test_menu_follows_owner_scale(app):
    from PyQt5.QtWidgets import QWidget
    from ui import menu_style as ms
    owner = QWidget()
    owner._scale = 1.3
    child = QWidget(owner)
    m = ms.make_menu(child)
    assert ms._scale_of(m) == 1.3 and abs(m.font().pointSizeF() - 13.0) < 0.01
    sub = ms.add_submenu(m, '글자 크기')
    assert ms._scale_of(sub) == 1.3 and abs(sub.font().pointSizeF() - 13.0) < 0.01
    plain = ms.make_menu()
    assert ms._scale_of(plain) == 1.0 and abs(plain.font().pointSizeF() - 10.0) < 0.01


def test_bigger_scale_makes_taller_items(app):
    from PyQt5.QtWidgets import QWidget
    from ui import menu_style as ms
    hs = []
    for k in (1.0, 1.4):
        owner = QWidget()
        owner._scale = k
        m = ms.make_menu(owner)
        ms.add_action(m, '오늘 한눈에 보기', icon='ph.squares-four')
        m.ensurePolished()
        hs.append((m.sizeHint().height(), m.sizeHint().width()))
    assert hs[1][0] > hs[0][0] and hs[1][1] > hs[0][1]
