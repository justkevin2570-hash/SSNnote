"""Claude(Anthropic) · ChatGPT(OpenAI Responses) 어댑터를 가짜 서버로 확인."""
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

pytest.importorskip('openai')
pytest.importorskip('anthropic')
pytest.importorskip('PyQt5')


class _Fake(BaseHTTPRequestHandler):
    seen = []

    def log_message(self, *a):
        pass

    def _send(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _bad_key(self):
        return (self.headers.get('Authorization') == 'Bearer bad' or self.headers.get('x-api-key') == 'bad')

    def do_GET(self):
        _Fake.seen.append(('GET', self.path, None))
        if self._bad_key():
            return self._send(401, {'error': {'type': 'authentication_error', 'message': 'bad'}})
        if self.path.startswith('/v1/models') and 'anthropic-version' in {k.lower() for k in self.headers}:
            return self._send(200, {'data': [
                {'type': 'model', 'id': 'claude-sonnet-5-5', 'display_name': 'S', 'created_at': '2026-09-01T00:00:00Z'},
                {'type': 'model', 'id': 'claude-haiku-5-5', 'display_name': 'H', 'created_at': '2026-08-01T00:00:00Z'}],
                'has_more': False, 'first_id': 'a', 'last_id': 'b'})
        return self._send(200, {'object': 'list', 'data': [
            {'id': 'gpt-5.4-mini', 'object': 'model', 'created': 300, 'owned_by': 'openai'},
            {'id': 'gpt-6-sol', 'object': 'model', 'created': 400, 'owned_by': 'openai'},
            {'id': 'gpt-realtime-2', 'object': 'model', 'created': 500, 'owned_by': 'openai'},
            {'id': 'gpt-image-2', 'object': 'model', 'created': 500, 'owned_by': 'openai'},
            {'id': 'text-embedding-3-large', 'object': 'model', 'created': 100, 'owned_by': 'openai'}]})

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        _Fake.seen.append(('POST', self.path, body))
        if self._bad_key():
            return self._send(401, {'error': {'type': 'authentication_error', 'message': 'bad'}})
        if self.path == '/v1/messages':
            return self._send(200, {
                'id': 'm', 'type': 'message', 'role': 'assistant', 'model': body['model'],
                'content': [{'type': 'thinking', 'thinking': '', 'signature': 's'},
                            {'type': 'text', 'text': '{"title": "회의 안내"}'}],
                'stop_reason': 'end_turn', 'stop_sequence': None, 'usage': {'input_tokens': 1, 'output_tokens': 1}})
        if self.path == '/v1/responses':
            resp = {'id': 'r', 'object': 'response', 'created_at': 0, 'model': body['model'], 'status': 'completed',
                    'output': [{'type': 'message', 'id': 'o', 'role': 'assistant', 'status': 'completed',
                                'content': [{'type': 'output_text', 'text': '2. 장소: 시청각실', 'annotations': []}]}],
                    'parallel_tool_calls': False, 'tool_choice': 'auto', 'tools': []}
            if not body.get('stream'):
                return self._send(200, resp)
            self.send_response(200)
            self.send_header('Content-Type', 'text/event-stream')
            self.end_headers()
            n = 0
            for ev in ([{'type': 'response.created', 'response': {**resp, 'status': 'in_progress', 'output': []}}]
                       + [{'type': 'response.output_text.delta', 'item_id': 'o', 'output_index': 0,
                           'content_index': 0, 'delta': p, 'logprobs': []} for p in ('2. 장소: ', '시청각실')]
                       + [{'type': 'response.completed', 'response': resp}]):
                n += 1
                self.wfile.write(f"event: {ev['type']}\ndata: ".encode()
                                 + json.dumps({**ev, 'sequence_number': n}, ensure_ascii=False).encode() + b'\n\n')
            return
        self._send(404, {'error': {'message': 'nope'}})


@pytest.fixture
def ai(tmp_path, monkeypatch):
    srv = HTTPServer(('127.0.0.1', 0), _Fake)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f'http://127.0.0.1:{srv.server_port}'
    import ai_client
    monkeypatch.setattr(ai_client, 'APPDATA_DIR', str(tmp_path))
    for name in ('AI_MODE_FILE', 'CLAUDE_KEY_FILE', 'CLAUDE_MODEL_FILE', 'OPENAI_KEY_FILE', 'OPENAI_MODEL_FILE'):
        monkeypatch.setattr(ai_client, name, str(tmp_path / (name.lower() + '.txt')))
    monkeypatch.setattr(ai_client, 'CLAUDE_BASE_URL', base)
    monkeypatch.setattr(ai_client, 'OPENAI_BASE_URL', base + '/v1')
    monkeypatch.setattr(ai_client, 'load_claude_key', lambda: 'good')
    monkeypatch.setattr(ai_client, 'load_openai_key', lambda: 'good')
    _Fake.seen = []
    yield ai_client
    srv.shutdown()


def test_only_supported_providers(ai):
    assert ai._AI_MODES == ('none', 'claude', 'openai', 'opencode', 'claude_code')
    for gone in ('GeminiAdapter', 'NvidiaAdapter', 'OllamaAdapter', 'is_local_adapter', 'AiImageWorker'):
        assert not hasattr(ai, gone)


def test_legacy_mode_turns_ai_off_with_notice(ai, tmp_path):
    (tmp_path / 'ai_mode_file.txt').write_text('gemini', encoding='utf-8')
    assert ai.load_ai_mode() == 'none' and ai.get_adapter() is None
    assert ai.legacy_ai_name() == 'Gemini'
    ai.save_ai_mode('claude')
    assert ai.legacy_ai_name() == '' and ai.get_adapter() is ai.ClaudeAdapter


def test_defaults(ai):
    assert ai.load_claude_model() == 'claude-haiku-5-5'
    assert ai.load_openai_model() == ai.OPENAI_DEFAULT_MODEL


def test_claude_models_test_and_complete(ai):
    assert ai.fetch_claude_models('good') == ['claude-sonnet-5-5', 'claude-haiku-5-5']
    assert ai.fetch_claude_models('bad') == ai.CLAUDE_DEFAULT_MODELS        # 실패하면 예비 목록
    assert ai.test_claude('good', 'claude-haiku-5-5') == (True, '연결 성공')
    ok, msg = ai.test_claude('bad', 'claude-haiku-5-5')
    assert not ok and 'API 키가 맞지 않습니다' in msg
    assert ai.ClaudeAdapter.complete('sys', '분석', image_b64='aGk=') == '{"title": "회의 안내"}'
    body = _Fake.seen[-1][2]
    assert body['model'] == 'claude-haiku-5-5' and body['messages'][0]['content'][0]['type'] == 'image'


def test_openai_models_are_filtered_newest_first(ai):
    assert ai.fetch_openai_models('good') == ['gpt-6-sol', 'gpt-5.4-mini']
    assert ai.fetch_openai_models('bad') == ai.OPENAI_DEFAULT_MODELS


def test_openai_uses_responses_api(ai):
    assert ai.test_openai('good', 'gpt-5.4-mini') == (True, '연결 성공')
    assert _Fake.seen[-1][1] == '/v1/responses'
    ok, msg = ai.test_openai('bad', 'gpt-5.4-mini')
    assert not ok and 'API 키가 맞지 않습니다' in msg

    assert ai.OpenAIAdapter.complete('지시', '분석', image_b64='aGk=') == '2. 장소: 시청각실'
    body = _Fake.seen[-1][2]
    assert body['instructions'] == '지시' and body['max_output_tokens'] == 4000
    assert [c['type'] for c in body['input'][0]['content']] == ['input_image', 'input_text']

    chunks, done, err = [], [], []
    ai.OpenAIAdapter.stream_text('장소 바꿔줘', chunks.append, done.append, err.append, system='지시')
    assert err == [] and ''.join(chunks) == '2. 장소: 시청각실' and done == ['2. 장소: 시청각실']
    assert _Fake.seen[-1][2]['stream'] is True


def test_friendly_error_messages(ai):
    class E(Exception):
        def __init__(self, code, text=''):
            super().__init__(text)
            self.status_code = code
    assert '크레딧' in ai.friendly_error(E(429, 'insufficient_quota'))
    assert '찾을 수 없습니다' in ai.friendly_error(E(404), 'gpt-x')
    assert '일시적인' in ai.friendly_error(E(503))
    assert '요금제' in ai.friendly_error(E(401), provider='opencode')
