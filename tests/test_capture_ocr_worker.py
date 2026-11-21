"""글자 인식 스레드: 결과 신호(result)와 스레드 종료(finished)가 따로 — 끝나기 전에 지우지 않는다."""
import pytest

pytest.importorskip('PyQt5')


def test_ocr_worker_result_signal_is_separate(monkeypatch):
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtGui import QPixmap
    from PyQt5.QtCore import QThread
    app = QApplication.instance() or QApplication([])
    import capture
    assert capture.OcrWorker.finished is QThread.finished          # QThread 종료 신호를 덮어쓰지 않음
    monkeypatch.setattr(capture, '_winrt_ocr', lambda b: '공문 제목')
    got = []
    pm = QPixmap(10, 10)
    w = capture.run_ocr(pm, got.append, got.append)
    assert w in capture._OCR_ALIVE
    w.wait(3000)
    for _ in range(20):
        app.processEvents()
    assert got == ['공문 제목'] and w not in capture._OCR_ALIVE
