"""고배율 화면에서 흐려 보이던 그림들: 앱 아이콘 · icon.ico · QSS 그림(@2x · @3x) · 공문 창 아이콘 라벨."""
import os

import pytest

pytest.importorskip('PyQt5')


@pytest.fixture(scope='module')
def app():
    from PyQt5.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def test_app_icon_has_every_size(app):
    from ui.common import app_icon, APP_ICON_SIZES
    icon = app_icon()
    got = {(s.width(), s.height()) for s in icon.availableSizes()}
    assert {(n, n) for n in APP_ICON_SIZES} <= got           # 16px 한 장을 늘려 쓰지 않음


def test_exe_icon_file_has_large_sizes():
    from PIL import Image
    path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'icon.ico')
    sizes = Image.open(path).info['sizes']
    for n in (16, 24, 32, 48, 256):                           # 바탕화면 · 탐색기 · 작업 표시줄용
        assert (n, n) in sizes


def test_qss_png_has_retina_variants(app, tmp_path, monkeypatch):
    import tempfile
    from PIL import Image
    monkeypatch.setattr(tempfile, 'gettempdir', lambda: str(tmp_path))
    from ui.common import icon_png, px
    path = icon_png('fa5s.check', '#FFFFFF', 10)
    s = px(10)
    base = path[:-4]
    assert Image.open(path).size == (s, s)
    assert Image.open(base + '@2x.png').size == (2 * s, 2 * s)   # Qt가 200% 화면에서 고름
    assert Image.open(base + '@3x.png').size == (3 * s, 3 * s)


def test_doc_editor_icons_are_vector_labels(app, tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    monkeypatch.setenv('APPDATA', str(tmp_path))
    import document_editor as de
    monkeypatch.setattr(de, 'load_ai_mode', lambda: 'none')
    monkeypatch.setattr(de.DocumentEditorWindow, '_STATE_FILE', str(tmp_path / 'st.json'))
    w = de.DocumentEditorWindow()
    try:
        # 머리줄 · 요청 칸 아이콘은 한 번 만든 그림(pixmap)이 아니라 그릴 때마다 벡터로
        for lab in (w._app_icon, w._spark):
            assert lab.pixmap() is None or lab.pixmap().isNull()
            assert hasattr(lab, 'setIcon')
        from ui.common import px
        assert w.btn_more.iconSize().width() == px(16) and w.btn_send.iconSize().width() == px(15)
    finally:
        w.close()
