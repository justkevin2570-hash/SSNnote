"""트레이 아이콘은 assets/tray.ico (크기별 그림), 설치 파일에도 들어감."""
import os
import sys
from unittest import mock


def test_tray_icon_loads_ico():
    sys.modules.setdefault('winreg', mock.MagicMock())
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtGui import QIcon
    QApplication.instance() or QApplication([])
    import main
    icon = main._tray_icon(QIcon())
    assert not icon.isNull() and len(icon.availableSizes()) >= 4


def test_tray_icon_bundled():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    assert os.path.exists(os.path.join(root, 'assets', 'tray.ico'))
    spec = open(os.path.join(root, 'SSNnote.spec'), encoding='utf-8').read()
    assert "('assets/tray.ico', 'assets')" in spec
