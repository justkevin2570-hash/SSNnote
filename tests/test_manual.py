import os
import re
import sys
import unittest.mock as um

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANUAL = os.path.join(ROOT, 'docs', 'manual.md')
_APP = None


def _github_anchor(heading: str) -> str:
    """GitHub 마크다운 제목 → 링크 이름 (소문자, 문장부호 제거, 공백 → -)."""
    s = heading.strip().lower()
    s = re.sub(r'[^\w\- ]', '', s)
    return s.replace(' ', '-')


def test_manual_toc_links_point_to_real_headings():
    text = open(MANUAL, encoding='utf-8').read()
    anchors = {_github_anchor(h) for h in re.findall(r'^#{2,3} (.+)$', text, re.M)}
    links = re.findall(r'\]\(#([^)]+)\)', text)
    assert links, '목차 링크가 없음'
    missing = [l for l in links if l not in anchors]
    assert not missing, missing


def test_manual_mentions_current_shortcuts_not_old_ones():
    text = open(MANUAL, encoding='utf-8').read()
    for key in ('Ctrl+Shift+E', 'Ctrl+Shift+X', 'Ctrl+K', 'Ctrl+Shift+F', 'Ctrl+Shift+R'):
        assert key in text
    assert 'ctrl + s' not in text.lower()


def test_manual_relative_links_exist():
    text = open(MANUAL, encoding='utf-8').read()
    for rel in re.findall(r'\]\((\.\./[^)#]+)\)', text):
        assert os.path.exists(os.path.normpath(os.path.join(ROOT, 'docs', rel))), rel


@pytest.fixture
def win(monkeypatch, tmp_path):
    pytest.importorskip('PyQt5')
    if sys.platform != 'win32':
        monkeypatch.setitem(sys.modules, 'winreg', um.MagicMock())
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = QApplication.instance() or QApplication([])     # 참조를 붙잡아 둬야 창을 만들 수 있음
    import window
    return window.MemoWindow(db.get_first_window_id())


def test_help_opens_manual_url(win):
    import window
    assert window.MemoWindow.MANUAL_URL.endswith('/blob/master/docs/manual.md')
    with um.patch('PyQt5.QtGui.QDesktopServices.openUrl', return_value=True) as op:
        win.show_help()
    assert op.call_args[0][0].toString() == window.MemoWindow.MANUAL_URL


def test_help_shows_address_when_browser_fails(win):
    from PyQt5.QtWidgets import QMessageBox
    shown = []
    with um.patch('PyQt5.QtGui.QDesktopServices.openUrl', return_value=False), \
         um.patch.object(QMessageBox, 'exec_', lambda self: shown.append(self.text()) or 0):
        win.show_help()
    assert shown and 'docs/manual.md' in shown[0]


def test_manual_documents_claude_subscription_and_new_features():
    text = open(MANUAL, encoding='utf-8').read()
    for heading in ('### Claude 구독으로 AI 쓰기', '### 공문 선택 부분 다듬기',
                    '### 예산 AI에게 지시하기 (잔액 맞추기, 되돌리기)'):
        assert heading in text, heading
    sub = text.split('### Claude 구독으로 AI 쓰기', 1)[1].split('\n## ', 1)[0]
    for phrase in ('구독 사용량에서 차감', '본인만 사용', '저장하지 않습니다', '[설치]', '[Claude로 로그인]', '[연결 확인]'):
        assert phrase in sub, phrase
    assert '--bare' not in text and 'setup-token' not in text          # 안내에 금지 방식을 권하지 않음
