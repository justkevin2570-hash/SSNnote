import os
import sys

# 프로젝트 루트를 import 경로에 추가 (python -m pytest tests 로 실행)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Windows가 아닌 곳에서도 autostart(winreg)를 불러올 수 있게
if sys.platform != 'win32':
    import unittest.mock as _um
    sys.modules.setdefault('winreg', _um.MagicMock())

# Qt 앱은 시험 내내 하나만 — 파일마다 만들었다 지우면(붙잡아 두지 않으면) 앞 앱에 딸린 타이머 ·
# 창이 지워진 채 남아, Windows에서 엉뚱한 테스트에서 프로세스가 끊길 수 있다 (v3.66 빌드에서 겪음)
try:
    from PyQt5.QtWidgets import QApplication as _QApp
    _SESSION_APP = _QApp.instance() or _QApp([])
except Exception:          # PyQt5가 없는 곳에서는 Qt 테스트가 알아서 건너뜀
    _SESSION_APP = None
