import io
import asyncio

from PyQt5.QtCore import Qt, QRect, QPoint, pyqtSignal, QThread
from PyQt5.QtGui import QPainter, QColor, QPen, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget


async def _winrt_ocr(png_bytes: bytes) -> str:
    """Windows 내장 OCR로 텍스트 인식."""
    from PIL import Image
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

    # WinRT OCR은 Bgra8/Gray8만 지원. Qt PNG는 Rgba8 → BMP로 변환하면 BGRA8로 디코딩됨.
    from PIL import ImageFilter
    img = Image.open(io.BytesIO(png_bytes)).convert('RGBA')
    w, h = img.size
    TARGET_H = 300
    MAX_DIM  = 4800
    if h < TARGET_H:
        scale = max(2, TARGET_H // h)
        scale = min(scale, MAX_DIM // max(w, h))
        if scale > 1:
            img = img.resize((w * scale, h * scale), Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    bmp_io = io.BytesIO()
    img.save(bmp_io, format='BMP')
    image_bytes = bmp_io.getvalue()

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(bytearray(image_bytes))
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language('ko'))
    if engine is None:
        engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise RuntimeError('한국어 OCR 언어팩이 없습니다. Windows 설정 → 시간 및 언어 → 언어에서 한국어를 추가하세요.')

    result = await engine.recognize_async(bitmap)
    # result.text가 일부 winrt 버전에서 빈 문자열 반환 → lines에서 직접 추출
    line_texts = [line.text for line in result.lines]
    return ' '.join(line_texts) if line_texts else result.text


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
            raw = asyncio.run(_winrt_ocr(self._png_bytes))
            text = raw.strip()
            if not text:
                error_msg = f'OCR 원문: [{raw}]'
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
    공문번호 정규화: 부서명-번호(YYYY. MM. DD.)

    패턴: 한글부서명 + '-' + 숫자번호 + '(' + YYYY. MM. DD. + ')'
    OCR 오인식 보정(O→0, l→1 등) 및 유연한 구분자 인식 포함.
    """
    import re

    # 0) 공백 정규화
    text = re.sub(r'[\n\t\r]+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # WinRT OCR이 한글 자모 사이에 공백 삽입 → 제거 (예: '중 등 교 육 과' → '중등교육과')
    text = re.sub(r'(?<=[가-힣]) (?=[가-힣])', '', text)

    # 숫자 오인식: 두 자리 숫자가 한자/특수문자로 오인식될 때 보정
    text = text.replace('田', '99')  # '99' → 한자 '田'으로 오인식

    # 전각 괄호 → 반각
    text = text.translate(str.maketrans('（）【】〔〕', '()()()'))
    # 대시 변형 → 하이픈
    for ch in '—–‑−_':
        text = text.replace(ch, '-')
    # 쉼표/중간점 → 온점 (날짜 구분자 오인식 대비)
    for ch in '·。,、':
        text = text.replace(ch, '.')

    # OCR 숫자 오인식 보정 (숫자 위치에 알파벳이 올 때)
    _digit_fix = str.maketrans({'O': '0', 'o': '0', 'I': '1', 'l': '1',
                                 '|': '1', 'S': '5', 'B': '8', 'Z': '2'})

    # 1) 한글 부서명 추출
    m = re.match(r'([가-힣]+)', text)
    if not m:
        return text.strip()
    dept = m.group(1)
    rest = text[m.end():].lstrip()

    # 2) 하이픈
    hm = re.match(r'-\s*', rest)
    if not hm:
        return dept
    rest = rest[hm.end():]

    # 3) 번호(숫자) — OCR 보정 후 추출 (WinRT OCR이 숫자 사이에 공백 삽입 가능)
    # '(' 이전까지의 모든 숫자를 수집 (공백 무시)
    rest_fixed = rest.translate(_digit_fix)
    m_paren = re.search(r'[\(\[\{（]', rest_fixed)
    num_end = m_paren.start() if m_paren else len(rest_fixed)
    number = ''.join(re.findall(r'\d', rest_fixed[:num_end]))
    if not number:
        return dept
    rest = rest_fixed[num_end:].strip()

    # 4) 여는 괄호 (없어도 진행, 전각·대괄호 허용)
    rest = re.sub(r'^[\(\[\{（]\s*', '', rest)

    # 5) 날짜 숫자 3개 추출 (년·월·일)
    nums = re.findall(r'\d+', rest.translate(_digit_fix))
    if len(nums) >= 3:
        y  = nums[0].zfill(4)
        mo = nums[1].zfill(2)
        d  = nums[2].zfill(2)
        # 연도 보정: OCR이 '2'를 'O'→'0'으로 오인식하면 20xx 범위 벗어남
        y_int = int(y)
        if y_int < 1990 or y_int > 2099:
            y = '20' + y[2:]
        return f'{dept}-{number}({y}. {mo}. {d}.)'

    # 날짜 없으면 부서명-번호만 반환
    return f'{dept}-{number}'


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
