import asyncio

from PyQt5.QtCore import Qt, QRect, QPoint, pyqtSignal, QThread
from PyQt5.QtGui import QPainter, QColor, QPen, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget


async def _winrt_recognize_async(png_bytes: bytes):
    """Windows 내장 OCR 결과(OcrResult). 글자와 위치(단어별 사각형)를 모두 담고 있다."""
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png_bytes)
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bmp = await decoder.get_software_bitmap_async()

    eng = OcrEngine.try_create_from_language(Language('ko'))
    if not eng:
        eng = OcrEngine.try_create_from_user_profile_languages()
    if not eng:
        raise RuntimeError('OCR 언어팩이 없습니다.')

    return await eng.recognize_async(bmp)


async def _winrt_ocr_async(png_bytes: bytes) -> str:
    """winrt 패키지를 통해 Windows 내장 OCR로 텍스트 인식."""
    result = await _winrt_recognize_async(png_bytes)
    return '\n'.join(line.text for line in result.lines)


async def _winrt_ocr_words_async(png_bytes: bytes) -> list:
    result = await _winrt_recognize_async(png_bytes)
    words = []
    for line in result.lines:
        for w in line.words:
            r = w.bounding_rect
            words.append((w.text, float(r.x), float(r.y), float(r.width), float(r.height)))
    return words


def _winrt_ocr(png_bytes: bytes) -> str:
    """동기 래퍼: 새 이벤트 루프에서 WinRT OCR 실행."""
    return asyncio.run(_winrt_ocr_async(png_bytes))


def ocr_words(png_bytes: bytes) -> list:
    """Windows OCR로 단어와 위치를 읽는다: [(글자, x, y, 폭, 높이)]. 표를 칸별로 나눌 때 쓴다.
    Windows가 아니거나 OCR 언어팩이 없으면 예외."""
    return asyncio.run(_winrt_ocr_words_async(png_bytes))


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
                padding = 5            # 그린 네모보다 사방 5픽셀만 넓게 (글자 끝은 지키고 옆 줄은 덜 들어오게)
                sel = sel.adjusted(-padding, -padding, padding, padding)
                sel = sel.intersected(QRect(0, 0, self._screenshot.width(), self._screenshot.height()))
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

    # 결과 신호 이름을 QThread.finished와 다르게: 예전에는 같은 이름이라 「스레드가 완전히 끝났다」는
    # 신호를 덮어써서, 결과를 받자마자 스레드를 지우다 아직 돌고 있는 스레드를 지워 앱이 멈출 수 있었다
    result = pyqtSignal(str, str)  # (text, error_msg)

    def __init__(self, png_bytes):
        super().__init__()
        # 한 장 또는 여러 판(크기 · 흑백을 달리한 같은 그림) — 여러 판이면 모두 읽고 가장 그럴듯한 글을 고른다
        self._variants = list(png_bytes) if isinstance(png_bytes, (list, tuple)) else [png_bytes]

    def run(self):
        texts, error_msg = [], ''
        for png in self._variants:
            try:
                t = _winrt_ocr(png).strip()
            except Exception as e:
                error_msg = f'OCR 오류: {e}'
                continue
            if t:
                texts.append(t)
        text = best_ocr_text(texts)
        if text:
            error_msg = ''
        elif not error_msg:
            error_msg = 'OCR 텍스트를 인식하지 못했습니다.'
        self.result.emit(text, error_msg)


_OCR_ALIVE = set()     # 끝날 때까지 스레드 객체를 붙잡아 둔다 (중간에 지워지지 않게)


def _ocr_score(text: str) -> float:
    """글자 인식 결과가 얼마나 그럴듯한지: 공문번호 · 날짜 모양이 있으면 크게, 한글이 많을수록, 이상한 기호가 많을수록 깎음."""
    import re
    score = 0.0
    if re.search(r'[가-힣A-Za-z]{2,}\s*-\s*\d{1,7}', text):
        score += 100
    if re.search(r'(19|20)\d{2}\s*[.,]\s*\d{1,2}\s*[.,]\s*\d{1,2}', text):
        score += 40
    score += len(re.findall(r'[가-힣]', text))
    score -= 3 * len(re.findall(r'[^\w\s가-힣().,\-:·/\[\]「」『』~%"\'※○●□■]', text))
    return score


def best_ocr_text(texts) -> str:
    """여러 판을 읽은 결과 중 가장 그럴듯한 글 (점수가 같으면 앞의 것)."""
    best, best_score = '', None
    for t in texts:
        if not t:
            continue
        sc = _ocr_score(t)
        if best_score is None or sc > best_score:
            best, best_score = t, sc
    return best


_OCR_MAX = 4000            # 너무 크면 Windows 글자 인식이 느려지고 거절함


def _scaled(img, k):
    w, h = img.width(), img.height()
    k = min(k, _OCR_MAX / max(w, 1), _OCR_MAX / max(h, 1))
    if k <= 1.01:
        return img
    return img.scaled(round(w * k), round(h * k), Qt.KeepAspectRatio, Qt.SmoothTransformation)


def _bw(img):
    """흑백 · 대비를 또렷하게 (흐린 회색 글자 · 색 바탕 공문용). numpy가 없으면 회색조만."""
    from PyQt5.QtGui import QImage
    g = img.convertToFormat(QImage.Format_Grayscale8)
    try:
        import numpy as np
        w, h, bpl = g.width(), g.height(), g.bytesPerLine()
        ptr = g.constBits()
        ptr.setsize(bpl * h)
        a = np.frombuffer(ptr, np.uint8).reshape(h, bpl)[:, :w].astype(np.float32)
        lo, hi = np.percentile(a, 2), np.percentile(a, 98)
        if hi - lo < 10:
            return g
        a = np.clip((a - lo) * (255.0 / (hi - lo)), 0, 255)
        a = np.where(a < 160, a * 0.6, 255.0)          # 글자는 더 진하게, 바탕은 하얗게
        out = np.ascontiguousarray(a.astype(np.uint8))
        res = QImage(out.data, w, h, w, QImage.Format_Grayscale8)
        return res.copy()
    except Exception:
        return g


def ocr_variants(pixmap: QPixmap) -> list:
    """같은 캡처를 여러 판으로: 3배 · 3배 흑백 · 4배 · 2배 (작은 글자는 판마다 다르게 틀려서 여럿 읽고 고름)."""
    img = pixmap.toImage()
    img.setDevicePixelRatio(1.0)
    small = max(img.width(), img.height()) <= 1400      # 아주 큰 캡처는 3배 한 판만 (시간)
    out = [_scaled(img, 3)]
    if small:
        out += [_bw(_scaled(img, 3)), _scaled(img, 4), _scaled(img, 2)]
    return [_image_to_png_bytes(x) for x in out]


def _image_to_png_bytes(img) -> bytes:
    from PyQt5.QtCore import QBuffer, QIODevice
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    img.save(buf, 'PNG')
    buf.close()
    return bytes(buf.data())


def run_ocr(pixmap: QPixmap, callback, error_callback=None):
    """pixmap을 OCR하여 결과를 callback(text)으로 전달. 에러 시 error_callback(msg) 호출.
    작은 글자도 잘 읽게 크기 · 흑백을 달리한 여러 판을 모두 읽고 가장 그럴듯한 결과를 고른다."""
    worker = OcrWorker(ocr_variants(pixmap))

    def _on_result(text, error_msg):
        try:
            if error_msg:
                if error_callback:
                    error_callback(error_msg)
            else:
                callback(text)
        except RuntimeError:          # 결과를 받을 창이 이미 닫힘
            pass

    def _on_thread_done():            # 스레드가 완전히 끝난 뒤에만 지운다
        _OCR_ALIVE.discard(worker)
        worker.deleteLater()

    _OCR_ALIVE.add(worker)
    worker.result.connect(_on_result)
    worker.finished.connect(_on_thread_done)
    worker.start()
    return worker  # 호출자가 참조를 유지할 수 있도록 반환


def _pixmap_to_png_bytes(pixmap: QPixmap) -> bytes:
    from PyQt5.QtCore import QBuffer, QIODevice
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    pixmap.save(buf, 'PNG')
    buf.close()
    return bytes(buf.data())


def find_doc_number(text: str) -> str:
    """캡처 글자에서 공문번호 줄을 찾아 「부서명-번호(날짜)」로 (못 찾으면 '').
    줄마다 찾고, 앞에 붙은 「관련 · 시행 · 접수 · 문서번호」 이름표는 뺀다. (공문 캡처 · 업무 메모 캡처가 같이 씀)"""
    import re
    lines = [re.sub(r'\s+', ' ', l).strip() for l in (text or '').splitlines()]
    for l in lines:
        if not l:
            continue
        m = re.search(r'[가-힣](?:\s?[가-힣])+\s*[-—–]\s*[\dOolI|SBZ][\dOolI|SBZ ]+', l)
        if m:
            head = re.sub(r'^(?:시\s?행|관\s?련|접\s?수|문\s?서\s?번\s?호)\s+', '', l[m.start():])
            cand = _normalize_doc_number(head)
            if re.search(r'-\d', cand):
                return cand
    return ''


def _normalize_doc_number(text: str) -> str:
    """
    공문번호 정규화: 부서명-번호(YYYY. MM. DD.)

    패턴: 한글부서명 + '-' + 숫자번호 + '(' + YYYY. MM. DD. + ')'
    OCR 오인식 보정(O→0, l→1 등) 및 유연한 구분자 인식 포함.
    """
    import re

    # 0) 공백 정규화
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
        mo = int(nums[1])     # 공문 표기: '2026. 11. 3.' (월·일 앞에 0을 붙이지 않음)
        d  = int(nums[2])
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
    """현재 가상 데스크탑 전체를 QPixmap으로 캡처. (v3.60에서 배율에 맞춰 바꿨다가 노트북에서 영역이 어긋나 되돌림)"""
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
