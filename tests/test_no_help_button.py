"""창 제목줄의 ? (도움말) 단추는 앱의 모든 창에서 없앤다 — QApplication을 만들기 전에 한 번 설정."""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _src(name):
    with open(os.path.join(ROOT, name), encoding='utf-8') as f:
        return f.read()


def test_main_disables_help_button_before_app_is_created():
    s = _src('main.py')
    attr = s.index('AA_DisableWindowContextHelpButton')
    assert attr < s.index('app = QApplication(sys.argv)')


def test_update_windows_use_one_title():
    import updater
    assert updater.UPDATE_TITLE == '서서니 노트 업데이트'
    s = _src('updater.py')
    assert 'SSNnote 업데이트' not in s and "'업데이트 실패'" not in s
