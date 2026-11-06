import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

pytest.importorskip('openai')
pytest.importorskip('PyQt5')


class _Fake(BaseHTTPRequestHandler):
    """OpenCode 흉내: 세션 헤더 없으면 400, 키가 'bad'면 401, 그 밖엔 정상 응답."""
    seen = []
    models = {'data': []}

    def log_message(self, *a):
        pass

    def _send(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _stream(self, body):
        """OpenCode식 스트리밍: 생각(reasoning_content) → 답(content) → 내용 없는 마지막 조각 → [DONE]."""
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.end_headers()

        def ev(obj):
            self.wfile.write(b'data: ' + json.dumps(obj, ensure_ascii=False).encode() + b'\n\n')
        base = {'id': 'c', 'object': 'chat.completion.chunk', 'created': 0, 'model': body['model']}
        for d in [{'role': 'assistant', 'reasoning_content': '요청을 살펴보면…'},
                  {'reasoning_content': ' 장소만 바꾸면 됨.'},
                  {'content': '2. 장소: '}, {'content': '시청각실'}]:
            ev({**base, 'choices': [{'index': 0, 'delta': d, 'finish_reason': None}]})
        ev({**base, 'choices': [{'index': 0, 'delta': {}, 'finish_reason': 'stop'}]})
        ev({**base, 'choices': [], 'usage': {'prompt_tokens': 5, 'completion_tokens': 9, 'total_tokens': 14}})
        self.wfile.write(b'data: [DONE]\n\n')

    def _messages(self, body):
        """Claude(messages) 방식: 생각(thinking) 블록 → 답(text) 블록. 모델 'refuse-me'는 거절."""
        stop = 'refusal' if body['model'] == 'refuse-me' else 'end_turn'
        text = '' if stop == 'refusal' else '2. 장소: 시청각실'
        if not body.get('stream'):
            return self._send(200, {
                'id': 'm', 'type': 'message', 'role': 'assistant', 'model': body['model'],
                'content': [{'type': 'thinking', 'thinking': '', 'signature': 'sig'},
                            {'type': 'text', 'text': text or ' '}],
                'stop_reason': stop, 'stop_sequence': None, 'usage': {'input_tokens': 5, 'output_tokens': 7}})
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.end_headers()

        def ev(name, obj):
            self.wfile.write(f'event: {name}\ndata: '.encode() + json.dumps(obj, ensure_ascii=False).encode() + b'\n\n')
        ev('message_start', {'type': 'message_start', 'message': {
            'id': 'm', 'type': 'message', 'role': 'assistant', 'model': body['model'], 'content': [],
            'stop_reason': None, 'stop_sequence': None, 'usage': {'input_tokens': 5, 'output_tokens': 0}}})
        ev('content_block_start', {'type': 'content_block_start', 'index': 0,
                                   'content_block': {'type': 'thinking', 'thinking': '', 'signature': ''}})
        ev('content_block_delta', {'type': 'content_block_delta', 'index': 0,
                                   'delta': {'type': 'signature_delta', 'signature': 'sig'}})
        ev('content_block_stop', {'type': 'content_block_stop', 'index': 0})
        ev('content_block_start', {'type': 'content_block_start', 'index': 1, 'content_block': {'type': 'text', 'text': ''}})
        for part in (['2. 장소: ', '시청각실'] if text else []):
            ev('content_block_delta', {'type': 'content_block_delta', 'index': 1, 'delta': {'type': 'text_delta', 'text': part}})
        ev('content_block_stop', {'type': 'content_block_stop', 'index': 1})
        ev('message_delta', {'type': 'message_delta', 'delta': {'stop_reason': stop, 'stop_sequence': None},
                             'usage': {'output_tokens': 7}})
        ev('message_stop', {'type': 'message_stop'})

    def do_GET(self):
        _Fake.seen.append(('GET', self.path, {k.lower(): v for k, v in self.headers.items()}, None))
        self._send(200, _Fake.models)

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])))
        _Fake.seen.append(('POST', self.path, {k.lower(): v for k, v in self.headers.items()}, body))   # 헤더 이름은 대소문자 무관
        if not self.headers.get('x-opencode-session'):
            return self._send(400, {'error': {'type': 'MissingSessionID', 'message': 'missing'}})
        if self.headers.get('Authorization') == 'Bearer bad' or self.headers.get('x-api-key') == 'bad':
            return self._send(401, {'error': {'type': 'server_error', 'message': 'Invalid credential'}})
        if self.path.endswith('/anthropic/v1/messages'):
            return self._messages(body)
        if body.get('stream'):
            return self._stream(body)
        self._send(200, {'id': 'x', 'object': 'chat.completion', 'created': 0, 'model': body['model'],
                         'choices': [{'index': 0, 'finish_reason': 'stop',
                                      'message': {'role': 'assistant', 'content': '{"title": "회의 안내"}'}}]})


@pytest.fixture
def oc(tmp_path, monkeypatch):
    srv = HTTPServer(('127.0.0.1', 0), _Fake)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f'http://127.0.0.1:{srv.server_port}'
    import ai_client
    monkeypatch.setattr(ai_client, 'APPDATA_DIR', str(tmp_path))
    monkeypatch.setattr(ai_client, 'OPENCODE_PLAN_FILE', str(tmp_path / 'plan.txt'))
    monkeypatch.setattr(ai_client, 'OPENCODE_MODEL_FILE', str(tmp_path / 'model.txt'))
    monkeypatch.setattr(ai_client, 'OPENCODE_VISION_FILE', str(tmp_path / 'vision.json'))
    monkeypatch.setattr(ai_client, 'OPENCODE_MODELS_FILE', str(tmp_path / 'models.json'))
    monkeypatch.setitem(ai_client.OPENCODE_PLANS, 'go', {'label': 'Go', 'base_url': base + '/inference/go/openai/v1',
                                                          'anthropic_base_url': base + '/inference/go/anthropic',
                                                          'models_url': base + '/inference/go/v1/models'})
    monkeypatch.setitem(ai_client.OPENCODE_PLANS, 'zen', {'label': 'Zen', 'base_url': base + '/inference/openai/v1',
                                                           'anthropic_base_url': base + '/inference/anthropic',
                                                           'models_url': base + '/inference/v1/models'})
    monkeypatch.setattr(ai_client, 'load_opencode_key', lambda: 'good-key')
    _Fake.seen = []
    yield ai_client
    srv.shutdown()


def test_real_plan_urls_match_opencode_docs():
    import ai_client
    assert ai_client.OPENCODE_PLANS['go']['base_url'] == 'https://opencode.ai/inference/go/openai/v1'
    assert ai_client.OPENCODE_PLANS['zen']['base_url'] == 'https://opencode.ai/inference/openai/v1'


def test_request_has_session_header_user_agent_and_plan_url(oc):
    oc.save_opencode_plan('go')
    oc.save_opencode_model('kimi-k3')
    out = oc.OpenCodeAdapter.complete('시스템', '공문 분석', image_b64='aGVsbG8=')
    assert out == '{"title": "회의 안내"}'
    method, path, headers, body = _Fake.seen[-1]
    assert path == '/inference/go/openai/v1/chat/completions'
    assert headers['x-opencode-session'].startswith('ssnnote-')
    assert headers['user-agent'].startswith('ssnnote/')
    assert headers['authorization'] == 'Bearer good-key'
    assert body['model'] == 'kimi-k3'
    assert body['messages'][1]['content'][0]['type'] == 'image_url'      # 이미지가 실제로 들어감

    oc.save_opencode_plan('zen')
    oc.OpenCodeAdapter.complete('s', 'p')
    assert _Fake.seen[-1][1] == '/inference/openai/v1/chat/completions'
    # 같은 실행 안에서는 세션 ID가 같아야 함 (OpenCode 권장)
    assert _Fake.seen[-1][2]['x-opencode-session'] == headers['x-opencode-session']


def test_streaming_path_also_sends_headers(oc):
    oc.save_opencode_plan('go')
    client = oc.OpenCodeAdapter._client()
    assert client.base_url.path.rstrip('/') == '/inference/go/openai/v1'
    assert client.default_headers['x-opencode-session'].startswith('ssnnote-')


def test_text_only_model_refuses_image_with_korean_guidance(oc):
    oc.save_opencode_model('glm-5.3')
    oc.save_opencode_vision_models(['kimi-k3'])
    with pytest.raises(RuntimeError, match='이미지를 읽지 못합니다'):
        oc.OpenCodeAdapter.complete('s', 'p', image_b64='aGVsbG8=')
    assert oc.OpenCodeAdapter.complete('s', '글자만') == '{"title": "회의 안내"}'   # 글자 분석은 됨


def test_fetch_models_keeps_chat_completions_and_marks_images(oc):
    _Fake.models = {'data': [
        {'id': 'kimi-k3', 'name': 'Kimi K3', 'type': 'chat', 'api': {'chat-completions': {}},
         'modalities': {'input': ['text', 'image/png']}},
        {'id': 'glm-5.3', 'name': 'GLM-5.3', 'type': 'chat', 'api': {'chat-completions': {}},
         'modalities': {'input': ['text']}},
        {'id': 'qwen3.8-max', 'type': 'chat', 'api': {'messages': {}}, 'modalities': {'input': ['text', 'image/png']}},
        {'id': 'embed-x', 'type': 'embedding', 'api': {'chat-completions': {}}},
        {'id': 'beam-free', 'type': 'chat', 'api': {'chat-completions': {}}, 'modalities': {'input': ['text']},
         'cost': {'text': {'input': {'price': 0}}}},
    ]}
    models = oc.fetch_opencode_models('go')
    assert models == [{'id': 'kimi-k3', 'name': 'Kimi K3', 'image': True, 'protocol': 'chat'},
                      {'id': 'glm-5.3', 'name': 'GLM-5.3', 'image': False, 'protocol': 'chat'},
                      {'id': 'qwen3.8-max', 'name': 'qwen3.8-max', 'image': True, 'protocol': 'messages'}]
    assert _Fake.seen[-1][1] == '/inference/go/v1/models'


def test_fetch_models_falls_back_when_offline(oc, monkeypatch):
    monkeypatch.setitem(oc.OPENCODE_PLANS, 'go', {'label': 'Go', 'base_url': 'http://127.0.0.1:9/v1',
                                                  'models_url': 'http://127.0.0.1:9/v1/models'})
    models = oc.fetch_opencode_models('go')
    assert [m['id'] for m in models] == oc.OPENCODE_DEFAULT_MODELS and models[0]['image'] is None


def test_connection_test_messages(oc):
    assert oc.test_opencode('go', 'good-key', 'kimi-k3') == (True, '연결 성공')
    ok, msg = oc.test_opencode('go', 'bad', 'kimi-k3')
    assert not ok and 'API 키가 맞지 않습니다' in msg


def test_default_model_and_plan(oc):
    assert oc.load_opencode_plan() == 'go'
    assert oc.load_opencode_model() == oc.OPENCODE_DEFAULT_MODEL == 'claude-haiku-5-5'


def test_streaming_reasoning_and_empty_final_chunk(oc):
    """생각 글자는 본문에 안 들어가고, 내용 없는 마지막 조각에도 오류가 나지 않아야 한다."""
    oc.save_opencode_plan('go')
    oc.save_opencode_model('kimi-k3')
    chunks, done, err = [], [], []
    oc.OpenCodeAdapter.stream_text('장소 바꿔줘', chunks.append, done.append, err.append, system='시스템')
    assert err == [] and ''.join(chunks) == '2. 장소: 시청각실' and done == ['2. 장소: 시청각실']
    method, path, headers, body = _Fake.seen[-1]
    assert body['stream'] is True and body['max_tokens'] >= 4000
    assert headers['x-opencode-session'].startswith('ssnnote-')


def _use_claude_model(oc, model='claude-haiku-5-5'):
    oc.save_opencode_plan('go')
    oc.save_opencode_model(model)
    oc.save_opencode_model_info([{'id': model, 'image': True, 'protocol': 'messages'}])


def test_claude_model_streams_via_messages_without_thinking(oc):
    _use_claude_model(oc)
    chunks, done, err = [], [], []
    oc.OpenCodeAdapter.stream_text('장소 바꿔줘', chunks.append, done.append, err.append, system='시스템')
    assert err == [] and ''.join(chunks) == '2. 장소: 시청각실' and done == ['2. 장소: 시청각실']
    method, path, headers, body = _Fake.seen[-1]
    assert path == '/inference/go/anthropic/v1/messages'
    # anthropic SDK 2.x는 키를 Authorization: Bearer로 보냄 (실제 OpenCode가 이 방식으로 키를 읽는 것 확인)
    assert headers['authorization'] == 'Bearer good-key' and headers['x-opencode-session'].startswith('ssnnote-')
    assert body['model'] == 'claude-haiku-5-5' and body['stream'] is True and body['system'] == '시스템'


def test_claude_model_complete_with_image(oc):
    _use_claude_model(oc)
    assert oc.OpenCodeAdapter.complete('s', '공문 분석', image_b64='aGVsbG8=') == '2. 장소: 시청각실'
    body = _Fake.seen[-1][3]
    assert body['messages'][0]['content'][0]['type'] == 'image' and body['max_tokens'] >= 4000


def test_claude_model_refusal_is_explained(oc):
    _use_claude_model(oc, 'refuse-me')
    with pytest.raises(RuntimeError, match='답하지 않았습니다'):
        oc.OpenCodeAdapter.complete('s', 'p')
    chunks, done, err = [], [], []
    oc.OpenCodeAdapter.stream_text('p', chunks.append, done.append, err.append)
    assert done == [] and err and '답하지 않았습니다' in err[0]


def test_connection_test_uses_model_protocol(oc):
    _use_claude_model(oc)
    assert oc.test_opencode('go', 'good-key', 'claude-haiku-5-5') == (True, '연결 성공')
    assert _Fake.seen[-1][1] == '/inference/go/anthropic/v1/messages'
    ok, msg = oc.test_opencode('go', 'bad', 'claude-haiku-5-5')
    assert not ok and 'API 키가 맞지 않습니다' in msg


def test_protocol_guess_without_saved_info(oc):
    assert oc.opencode_model_protocol('claude-haiku-5-5') == 'messages'
    assert oc.opencode_model_protocol('kimi-k3') == 'chat'
