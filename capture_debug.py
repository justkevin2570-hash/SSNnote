"""캡처 확인 기록 — 글자 인식이 안 될 때 원인을 찾으려고, 마지막 캡처 한 번의 과정만 이 PC에 남긴다.

%APPDATA%\\SSNnote\\capture_debug\\ 에 캡처 그림(capture.png)과 단계별 기록(log.txt).
새로 캡처할 때마다 지우고 새로 쓴다 (예전 캡처는 남지 않음). 밖으로 보내지 않는다.
"""
import os
import time

_DIR = None


def folder() -> str:
    base = os.path.join(os.environ.get('APPDATA', '.'), 'SSNnote', 'capture_debug')
    os.makedirs(base, exist_ok=True)
    return base


def start(pixmap=None, note=''):
    """새 캡처 — 지난 기록을 지우고 그림을 남긴다."""
    try:
        d = folder()
        for name in ('capture.png', 'log.txt'):
            p = os.path.join(d, name)
            if os.path.exists(p):
                os.remove(p)
        if pixmap is not None:
            pixmap.save(os.path.join(d, 'capture.png'), 'PNG')
            log(f'캡처 {pixmap.width()}x{pixmap.height()} dpr={pixmap.devicePixelRatio()} {note}')
    except Exception:
        pass


def log(msg: str):
    try:
        with open(os.path.join(folder(), 'log.txt'), 'a', encoding='utf-8') as f:
            f.write(time.strftime('%H:%M:%S ') + str(msg).replace('\r', '') + '\n')
    except Exception:
        pass
