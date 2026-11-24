"""개인정보 보호 모드에서 캡처: 개인정보가 안 보이는 캡처만 그림으로, 보이면 가린 글자만 (설정은 그대로 켜짐)."""
import pytest

pytest.importorskip('PyQt5')


@pytest.fixture
def env(monkeypatch):
    import ai_client
    import school_profile

    class Img:
        supports_image = True
    state = {'ad': Img, 'privacy': True, 'mode': 'claude_code'}
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: state['ad'])
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: state['mode'])
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: state['privacy'])
    return state


def _ok(text):
    from ui.capture_panel import CapturePanel
    return CapturePanel._clean_enough_for_image(text)


def test_clean_capture_goes_as_image(env):
    assert _ok('교육과정과-12873(2026. 11. 18.) 2027학년도 고교학점제 운영 계획 제출 안내') is True


def test_personal_info_stays_text(env):
    assert _ok('3반 5번 홍길동 010-1234-5678') is False
    assert _ok('보호자: 김철수') is False


def test_unreadable_or_text_only_ai(env):
    assert _ok('') is False and _ok('ab') is False                 # 못 읽었으면 확인 불가 → 그림 안 보냄

    class NoImg:
        supports_image = False
    env['ad'] = NoImg
    assert _ok('교육과정과-12873(2026. 11. 18.)') is False


def test_privacy_off_uses_normal_path(env):
    env['privacy'] = False                                          # 보호 모드가 꺼져 있으면 원래부터 그림
    assert _ok('교육과정과-12873') is False


def test_worker_respects_image_ok(monkeypatch):
    import ai_client
    import doc_ai
    sent = {}

    class Ad:
        supports_image = True

        @staticmethod
        def get_api_key():
            return 'x'

        @staticmethod
        def complete(system, prompt, image_b64=None):
            sent['image'] = image_b64
            return '{"title": "제목"}'
    monkeypatch.setattr(ai_client, 'get_adapter', lambda: Ad)
    monkeypatch.setattr(ai_client, 'text_only_required', lambda: True)
    w = doc_ai.DocAnalysisWorker(ocr_text='', image_ok=True)
    w._image_b64 = 'aGk='
    got = []
    w.analyzed.connect(lambda a, info: got.append(info['mode']))
    w.failed.connect(lambda m: got.append('fail:' + m))
    w.run()
    assert got == ['image'] and sent['image'] == 'aGk='
