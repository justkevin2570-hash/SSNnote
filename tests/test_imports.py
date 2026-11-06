"""앱의 모든 모듈을 실제로 불러온다 — 빌드 환경에 빠진 패키지(예: v2.68의 numpy)를 릴리스 전에 잡기."""
import importlib
import sys
import unittest.mock as um

import pytest

pytest.importorskip('PyQt5')

MODULES = ['db', 'dashboard', 'doc_ai', 'privacy', 'school_profile', 'korean_search', 'claude_connect',
           'ai_client', 'updater', 'capture', 'window', 'document_editor', 'budget', 'selfcheck',
           'ui.common', 'ui.theme', 'ui.widgets', 'ui.dialogs', 'ui.chrome', 'ui.task_row', 'ui.document_row',
           'ui.calendar_view', 'ui.capture_dialog', 'ui.profile_dialog', 'ui.dashboard_dialog',
           'ui.command_palette', 'ui.claude_connect_dialog']


@pytest.fixture(autouse=True)
def _no_winreg_on_linux(monkeypatch):
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())


@pytest.mark.parametrize('name', MODULES)
def test_module_imports(name):
    importlib.import_module(name)


def test_numpy_complete():
    np = pytest.importorskip('numpy')
    assert hasattr(np, 'ndarray')


def test_document_editor_opens_without_numpy(monkeypatch):
    """numpy가 없어도 공문 작성 창 모듈은 열려야 한다 (유사 공문 찾기만 꺼짐)."""
    monkeypatch.setitem(sys.modules, 'numpy', None)
    monkeypatch.delitem(sys.modules, 'rag', raising=False)
    monkeypatch.delitem(sys.modules, 'document_editor', raising=False)
    importlib.import_module('document_editor')
