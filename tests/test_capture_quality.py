"""캡처 글자 인식 개선: 화면 배율과 관계없이 실제 픽셀로 자르기 · 여러 판을 읽고 가장 그럴듯한 결과 고르기."""
import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def app():
    global _APP
    from PyQt5.QtWidgets import QApplication
    _APP = QApplication.instance() or QApplication([])
    return _APP


def test_best_text_prefers_doc_number():
    from capture import best_ocr_text
    texts = ['교무기획부 -1Z34 (2O26 . 3 , 4.)', '교무기획부-1234(2026. 3. 4.)', '']
    assert best_ocr_text(texts) == '교무기획부-1234(2026. 3. 4.)'
    assert best_ocr_text(['가나다', '가나다라마']) == '가나다라마'      # 한글을 더 많이 읽은 쪽
    assert best_ocr_text(['가나 다', '가나@#$ 다']) == '가나 다'          # 이상한 기호가 많으면 깎음
    assert best_ocr_text([]) == ''


def test_variants(app):
    from PyQt5.QtGui import QPixmap
    from capture import ocr_variants
    v = ocr_variants(QPixmap(200, 40))
    assert len(v) == 4 and all(isinstance(b, bytes) and b[:4] == b'\x89PNG' for b in v)
    assert len(ocr_variants(QPixmap(2000, 300))) == 1                 # 큰 캡처는 한 판만


def test_worker_picks_best(app, monkeypatch):
    import capture
    outs = iter(['교무기획부 -1Z34', '교무기획부-1234(2026. 3. 4.)', '', '흐림'])
    monkeypatch.setattr(capture, '_winrt_ocr', lambda png: next(outs))
    got = []
    w = capture.OcrWorker([b'a', b'b', b'c', b'd'])
    w.result.connect(lambda t, e: got.append((t, e)))
    w.run()
    assert got == [('교무기획부-1234(2026. 3. 4.)', '')]


def test_worker_error_when_nothing(app, monkeypatch):
    import capture

    def boom(png):
        raise RuntimeError('언어팩 없음')
    monkeypatch.setattr(capture, '_winrt_ocr', boom)
    got = []
    w = capture.OcrWorker(b'a')
    w.result.connect(lambda t, e: got.append((t, e)))
    w.run()
    assert got[0][0] == '' and '언어팩 없음' in got[0][1]


def test_capture_debug_keeps_only_last(tmp_path, monkeypatch, app):
    from PyQt5.QtGui import QPixmap
    import capture_debug
    monkeypatch.setenv('APPDATA', str(tmp_path))
    capture_debug.start(QPixmap(30, 10), '첫 캡처')
    capture_debug.log('하나')
    capture_debug.start(QPixmap(40, 10), '둘째')
    log = (tmp_path / 'SSNnote' / 'capture_debug' / 'log.txt').read_text(encoding='utf-8')
    assert '둘째' in log and '첫 캡처' not in log and '하나' not in log
    assert (tmp_path / 'SSNnote' / 'capture_debug' / 'capture.png').exists()


def test_overlay_pads_5px(app):
    from PyQt5.QtCore import QPoint, Qt, QEvent
    from PyQt5.QtGui import QPixmap, QMouseEvent
    from capture import ScreenCaptureOverlay
    ov = ScreenCaptureOverlay(QPixmap(800, 600))
    got = []
    ov.region_captured.connect(got.append)
    ov.mousePressEvent(QMouseEvent(QEvent.MouseButtonPress, QPoint(100, 100), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier))
    ov.mouseReleaseEvent(QMouseEvent(QEvent.MouseButtonRelease, QPoint(160, 130), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier))
    assert (got[0].width(), got[0].height()) == (61 + 10, 31 + 10)     # 사방 5픽셀씩
