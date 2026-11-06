"""화면 크기에 맞추기 (노트북·모니터 분리 대비)."""
import os
import pytest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
pytest.importorskip('PyQt5')


_APP = None


@pytest.fixture(scope='module')
def ag():
    global _APP
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])     # 참조를 붙잡아 둬야 앱이 사라지지 않음
    return _APP.primaryScreen().availableGeometry()


def test_big_saved_window_shrinks_into_screen(ag):
    from ui.common import fit_to_screen
    x, y, w, h = fit_to_screen(ag.left() + 50, ag.top() + 50, ag.width() * 2, ag.height() * 2)
    assert w <= ag.width() * 0.92 + 1 and h <= ag.height() * 0.92 + 1
    assert ag.left() <= x and x + w <= ag.left() + ag.width()
    assert ag.top() <= y and y + h <= ag.top() + ag.height()


def test_window_on_missing_monitor_comes_back(ag):
    from ui.common import fit_to_screen
    x, y, w, h = fit_to_screen(ag.left() + ag.width() + 900, ag.top() + 40, 300, 200)   # 떼어 낸 오른쪽 모니터
    assert (w, h) == (300, 200)
    assert ag.left() <= x and x + w <= ag.left() + ag.width()


def test_window_that_fits_is_untouched(ag):
    from ui.common import fit_to_screen
    r = (ag.left() + 10, ag.top() + 10, 200, 150)
    assert fit_to_screen(*r) == r


def test_minimum_size_respected(ag):
    from ui.common import fit_to_screen
    _, _, w, h = fit_to_screen(0, 0, 50, 50, min_w=120, min_h=90)
    assert (w, h) == (120, 90)


def test_default_geometry_is_centered_and_fits(ag):
    from ui.common import default_geometry
    x, y, w, h = default_geometry(5000, 5000)
    assert w <= ag.width() and h <= ag.height()
    assert abs((x - ag.left()) - (ag.left() + ag.width() - x - w)) <= 1


def test_high_dpi_rounding_is_whole_number():
    """소수점 배율(1.25배)은 Qt 5에서 글씨가 흐려지므로 정수 배율만 쓴다."""
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
    from ui.common import setup_high_dpi
    if not hasattr(QApplication, 'setHighDpiScaleFactorRoundingPolicy'):
        pytest.skip('Qt 5.14 미만')
    setup_high_dpi()
    assert QApplication.highDpiScaleFactorRoundingPolicy() == Qt.HighDpiScaleFactorRoundingPolicy.Floor
