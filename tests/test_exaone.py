"""EXAONE(LG AI연구원) 연결 — 가짜 서버로 요청·응답, 생각 과정 거르기, 이미지·개인정보 규칙, 키 30일 안내."""
import json
import threading
from datetime import date
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

pytest.importorskip('openai')
pytest.importorskip('PyQt5')


class _Fake(BaseHTTPRequestHandler):
    seen = []

    def log_message(self, *a):
        pass

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        _Fake.seen.append((self.path, self.headers.get('Authorization'), body))
        auth = self.headers.get('Authorization')
        if auth == 'Bearer bad':
            return self._json(401, {'error': {'message': 'invalid key'}, 'result_code': 401})
        if auth == 'Bearer limited':
            return self._json(429, {'error': {'message': 'slow down'}})
        if self.path != '/v1/chat/completions' or body.get('model') != 'k-exaone_v2.5':
            return self._json(404, {'error': {'message': 'no model'}})
        asked = body['messages'][-1]['content']
        if not body.get('stream'):
            if 'LENGTH' in asked:
                msg, fin = {'role': 'assistant', 'content': '', 'reasoning_content': '생각만 하다가…'}, 'length'
            else:
                msg, fin = {'role': 'assistant', 'content': '\n\n2. 장소: 시청각실', 'reasoning_content': '생각'}, 'stop'
            return self._json(200, {'id': 'c', 'object': 'chat.completion', 'created': 0, 'model': body['model'],
                                    'choices': [{'index': 0, 'message': msg, 'finish_reason': fin}],
                                    'usage': {'prompt_tokens': 1, 'completion_tokens': 1, 'total_tokens': 2}})
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.end_headers()
        deltas = [{'role': 'assistant', 'content': None, 'reasoning_content': '먼저 생각'},
                  {'reasoning_content': '을 해 보면'}, {'content': '\n\n'}, {'content': '2. 장소: '}, {'content': '시청각실'}]
        for i, d in enumerate(deltas):
            chunk = {'id': 'c', 'object': 'chat.completion.chunk', 'created': 0, 'model': body['model'],
                     'choices': [{'index': 0, 'delta': d, 'finish_reason': None}]}
            self.wfile.write(b'data: ' + json.dumps(chunk, ensure_ascii=False).encode() + b'\n\n')
        end = {'id': 'c', 'object': 'chat.completion.chunk', 'created': 0, 'model': body['model'],
               'choices': [{'index': 0, 'delta': {}, 'finish_reason': 'stop'}]}
        self.wfile.write(b'data: ' + json.dumps(end).encode() + b'\n\ndata: [DONE]\n\n')


@pytest.fixture
def ai(tmp_path, monkeypatch):
    srv = HTTPServer(('127.0.0.1', 0), _Fake)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    import ai_client
    monkeypatch.setattr(ai_client, 'APPDATA_DIR', str(tmp_path))
    for name in ('AI_MODE_FILE', 'EXAONE_KEY_FILE', 'EXAONE_MODEL_FILE', 'EXAONE_KEY_DATE_FILE'):
        monkeypatch.setattr(ai_client, name, str(tmp_path / (name.lower() + '.txt')))
    monkeypatch.setattr(ai_client, 'EXAONE_BASE_URL', f'http://127.0.0.1:{srv.server_port}/v1')
    monkeypatch.setattr(ai_client, 'load_exaone_key', lambda: 'good')
    _Fake.seen = []
    yield ai_client
    srv.shutdown()


def test_registered_as_a_provider(ai):
    assert 'exaone' in ai._AI_MODES and ai.PROVIDER_LABELS['exaone'] == 'EXAONE'
    ai.save_ai_mode('exaone')
    assert ai.get_adapter() is ai.ExaoneAdapter
    assert ai.load_exaone_model() == 'k-exaone_v2.5'
    assert ai.ai_status_text() == 'EXAONE · k-exaone_v2.5'


def test_connection_test(ai):
    assert ai.test_exaone('good', 'k-exaone_v2.5') == (True, '연결 성공')
    path, auth, body = _Fake.seen[-1]
    assert path == '/v1/chat/completions' and auth == 'Bearer good'
    assert body['reasoning_effort'] == 'low'                        # 생각 정도를 낮춰 빠르게
    ok, msg = ai.test_exaone('bad', 'k-exaone_v2.5')
    assert not ok and '30일' in msg and 'docs.exaone.ai' in msg      # 키 기간 안내
    ok, msg = ai.test_exaone('good', 'k-exaone_old')
    assert not ok and '모델 이름이 바뀌었을 수' in msg
    ok, msg = ai.test_exaone('limited', 'k-exaone_v2.5')
    assert not ok and '사용 한도' in msg


def test_complete_trims_and_leaves_room_for_thinking(ai):
    out = ai.ExaoneAdapter.complete('지시', '장소를 알려줘')
    assert out == '2. 장소: 시청각실'                                  # 앞의 줄바꿈 제거, 생각 과정은 받지 않음
    body = _Fake.seen[-1][2]
    assert body['max_tokens'] >= 8000                               # 생각 과정도 max_tokens에 들어감
    assert body['messages'][0] == {'role': 'system', 'content': '지시'}


def test_complete_explains_empty_answer_from_thinking(ai):
    with pytest.raises(RuntimeError, match='생각하는 데'):
        ai.ExaoneAdapter.complete('지시', 'LENGTH')


def test_images_are_refused(ai):
    with pytest.raises(RuntimeError, match='이미지를 읽지 못해요'):
        ai.ExaoneAdapter.complete('지시', '분석', image_b64='aGk=')
    assert _Fake.seen == []                                          # 서버로 보내지도 않음


def test_stream_hides_thinking_and_leading_blank(ai):
    chunks, done, errors = [], [], []
    ai.ExaoneAdapter.stream_text('장소는?', chunks.append, done.append, errors.append, system='지시')
    assert errors == [] and chunks == ['2. 장소: ', '시청각실'] and done == ['2. 장소: 시청각실']
    body = _Fake.seen[-1][2]
    assert body['stream'] is True and body['reasoning_effort'] == 'low' and body['max_tokens'] >= 8000


def test_stream_error_is_friendly(ai, monkeypatch):
    monkeypatch.setattr(ai, 'load_exaone_key', lambda: 'bad')
    errors = []
    ai.ExaoneAdapter.stream_text('x', lambda t: None, lambda t: None, errors.append)
    assert errors and '30일' in errors[0]


def test_always_masks_and_text_only(ai, monkeypatch):
    import school_profile
    ai.save_ai_mode('exaone')
    for privacy in (True, False):                                    # 보호 모드를 꺼도 EXAONE은 항상 가림 · 글자만
        monkeypatch.setattr(school_profile, 'privacy_enabled', lambda p=privacy: p)
        assert ai.masking_required() is True and ai.text_only_required() is True
    monkeypatch.setattr(ai, 'load_claude_key', lambda: 'k')
    ai.save_ai_mode('claude')                                        # 다른 AI는 보호 모드를 따름
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: False)
    assert ai.masking_required() is False and ai.text_only_required() is False
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: True)
    assert ai.masking_required() is True and ai.text_only_required() is True
    ai.save_ai_mode('none')
    monkeypatch.setattr(school_profile, 'privacy_enabled', lambda: False)
    assert ai.text_only_required() is False


def test_key_age_notice(ai, tmp_path):
    ai.save_ai_mode('exaone')
    (tmp_path / 'exaone_key_date_file.txt').write_text('2026-10-10', encoding='utf-8')
    assert ai.exaone_key_age_days(date(2026, 11, 9)) == 30
    assert '30일이 지났' in ai.exaone_key_notice(date(2026, 11, 9))
    assert '26일째' in ai.exaone_key_notice(date(2026, 11, 5))
    assert ai.exaone_key_notice(date(2026, 10, 20)) == ''            # 아직 이름
    ai.save_ai_mode('claude')
    assert ai.exaone_key_notice(date(2026, 11, 9)) == ''             # EXAONE을 안 쓰면 안내 없음


def test_saving_key_records_the_day(ai, tmp_path):
    ai.save_exaone_key('abc', persist=False)
    assert ai.exaone_key_age_days() == 0
