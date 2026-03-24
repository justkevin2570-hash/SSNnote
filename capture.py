import io

from PyQt5.QtCore import Qt, QRect, QPoint, pyqtSignal, QThread
from PyQt5.QtGui import QPainter, QColor, QPen, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget


class ScreenCaptureOverlay(QWidget):
    """전체화면 반투명 오버레이. 마우스 드래그로 영역을 선택하면 region_captured 시그널 emit."""

    region_captured = pyqtSignal(QPixmap)
    cancelled = pyqtSignal()

    def __init__(self, screenshot: QPixmap):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self._screenshot = screenshot
        self._start: QPoint | None = None
        self._end: QPoint | None = None
        self._captured = False
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setCursor(Qt.CrossCursor)

        virtual = QRect()
        for screen in QApplication.screens():
            virtual = virtual.united(screen.geometry())
        self.setGeometry(virtual)
        # showFullScreen()은 외부에서 시그널 연결 후 호출

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(self.rect(), self._screenshot)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))
        if self._start and self._end:
            sel = QRect(self._start, self._end).normalized()
            painter.drawPixmap(sel, self._screenshot, sel)
            pen = QPen(QColor('#f7c948'), 2)
            painter.setPen(pen)
            painter.drawRect(sel)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._start = event.pos()
            self._end = event.pos()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and self._start:
            self._end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._start:
            self._end = event.pos()
            sel = QRect(self._start, self._end).normalized()
            if sel.width() > 5 and sel.height() > 5:
                cropped = self._screenshot.copy(sel)
                self._captured = True
                self.region_captured.emit(cropped)  # emit 먼저, close 나중
            self.close()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()

    def closeEvent(self, event):
        if not self._captured:
            self.cancelled.emit()
        super().closeEvent(event)


class OcrWorker(QThread):
    """별도 스레드에서 OCR을 실행하고 결과를 시그널로 메인 스레드에 전달."""

    finished = pyqtSignal(str, str)  # (text, error_msg)

    def __init__(self, png_bytes: bytes):
        super().__init__()
        self._png_bytes = png_bytes

    def run(self):
        text = ''
        error_msg = ''
        try:
            import winocr
            from PIL import Image, ImageEnhance, ImageFilter
            img = Image.open(io.BytesIO(self._png_bytes))

            # OCR 정확도 향상을 위한 전처리
            img = img.convert('L')                              # 그레이스케일
            img = img.resize((img.width * 3, img.height * 3),  # 3배 확대
                             Image.LANCZOS)
            img = ImageEnhance.Contrast(img).enhance(2.0)      # 대비 강화
            img = img.filter(ImageFilter.SHARPEN)               # 선명화

            result = winocr.recognize_pil_sync(img, 'ko')
            if isinstance(result, dict):
                raw = result.get('text', '').strip()
                text = _normalize_doc_number(raw)
            else:
                error_msg = f'예상치 못한 OCR 반환 형식: {type(result)}'
        except AssertionError:
            error_msg = ('한국어 OCR 언어 팩이 설치되어 있지 않습니다.\n'
                         'Windows 설정 → 시간 및 언어 → 언어 → 한국어 추가 후\n'
                         '선택적 기능에서 "기본 입력" 항목을 설치해 주세요.')
        except Exception as e:
            error_msg = f'OCR 오류: {e}'
        self.finished.emit(text, error_msg)


def run_ocr(pixmap: QPixmap, callback, error_callback=None):
    """pixmap을 OCR하여 결과를 callback(text)으로 전달. 에러 시 error_callback(msg) 호출."""
    # QPixmap → PNG bytes (메인 스레드에서 미리 변환)
    png_bytes = _pixmap_to_png_bytes(pixmap)

    worker = OcrWorker(png_bytes)

    def _on_finished(text, error_msg):
        if error_msg:
            if error_callback:
                error_callback(error_msg)
        else:
            callback(text)
        # 스레드 객체 정리
        worker.deleteLater()

    worker.finished.connect(_on_finished)
    worker.start()
    return worker  # 호출자가 참조를 유지할 수 있도록 반환


def _pixmap_to_png_bytes(pixmap: QPixmap) -> bytes:
    from PyQt5.QtCore import QBuffer, QIODevice
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    pixmap.save(buf, 'PNG')
    buf.close()
    return bytes(buf.data())


def _normalize_doc_number(text: str) -> str:
    """
    공문번호를 규칙에 맞게 정규화한다.

    목표 형식: 부서명-번호(YYYY. MM. DD.)
      - 부서명: 한글+영숫자, 사이 공백 제거
      - 번호:   숫자, 부서명과 '-'로 연결, 공백 없음
      - 날짜:   4자리년도. 2자리월. 2자리일.  (점 뒤 한 칸 띄어쓰기)
      - 닫는 괄호 이후 문자는 모두 삭제
    """
    import re

    # 0) 줄바꿈·탭 → 공백
    text = text.replace('\n', ' ').replace('\t', ' ')

    # 1) 닫는 괄호 이후 모두 삭제
    if ')' in text:
        text = text[:text.index(')') + 1]

    # 2) '(' 기준으로 앞부분(부서명-번호)과 뒷부분(날짜) 분리
    if '(' not in text:
        # 괄호가 없으면 부서명-번호 부분만 정리 후 반환
        prefix = _clean_prefix(text)
        return prefix

    paren_idx = text.index('(')
    prefix_raw = text[:paren_idx]
    date_raw   = text[paren_idx + 1:].rstrip(')')  # 괄호 안 내용만

    # 3) 부서명-번호 정리
    prefix = _clean_prefix(prefix_raw)

    # 4) 날짜 숫자 추출 → YYYY. MM. DD. 형식
    # 공백 제거 후 파싱 ('1 1' → '11' 처리)
    date_raw = date_raw.replace(' ', '')
    nums = re.findall(r'\d+', date_raw)
    if len(nums) >= 3:
        y  = nums[0].zfill(4)
        mo = nums[1].zfill(2)
        d  = nums[2].zfill(2)
        date_part = f'{y}. {mo}. {d}.'
    else:
        date_part = date_raw.strip()

    return f'{prefix}({date_part})'


def _clean_prefix(text: str) -> str:
    """부서명-번호 부분의 공백을 정리한다."""
    import re
    # 한글 사이 공백 제거
    text = re.sub(r'(?<=[\uAC00-\uD7A3])\s+(?=[\uAC00-\uD7A3])', '', text)
    # 한글과 영숫자 사이 공백 제거
    text = re.sub(r'(?<=[\uAC00-\uD7A3])\s+(?=[A-Za-z0-9])', '', text)
    text = re.sub(r'(?<=[A-Za-z0-9])\s+(?=[\uAC00-\uD7A3])', '', text)
    # 하이픈 주변 공백 제거
    text = re.sub(r'\s*-\s*', '-', text)
    return text.strip()


def grab_fullscreen() -> QPixmap:
    """현재 가상 데스크탑 전체를 QPixmap으로 캡처."""
    virtual = QRect()
    for screen in QApplication.screens():
        virtual = virtual.united(screen.geometry())
    result = QPixmap(virtual.size())
    result.fill(Qt.black)
    painter = QPainter(result)
    for screen in QApplication.screens():
        geo = screen.geometry()
        pix = screen.grabWindow(0)
        painter.drawPixmap(geo.x() - virtual.x(), geo.y() - virtual.y(), pix)
    painter.end()
    return result
