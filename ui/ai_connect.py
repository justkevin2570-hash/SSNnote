"""AI 연결 — 공문 작성 창과 예산 관리 창이 함께 쓰는 연결 메뉴와 연결 순서.

- show_menu(): [AI 설정] 버튼 아래 메뉴 (OpenCode · Claude · ChatGPT · Claude 구독 / AI 끄기 + 창별 추가 항목)
- connect(mode): 고른 AI를 연결 (키 → 모델 → 연결 확인 → 저장). 성공하면 True
- status(): 버튼에 쓸 글과 점 색
- on_change(slot) / notify_changed(): 연결을 바꾸면 알림 → 열려 있는 다른 창도 버튼·입력칸을 바로 맞춘다

설정 대화상자(ApiKeyDialog · ClaudeCodeDialog · _ModelSelectDialog)는 document_editor에 있고, 필요할 때만 불러온다.
"""
from PyQt5.QtCore import QObject, Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QDialog, QInputDialog, QMenu, QMessageBox

from ai_client import (PROVIDER_LABELS, load_ai_mode, save_ai_mode, legacy_ai_name, load_external_model_name,
                       load_claude_key, save_claude_key, load_claude_model, save_claude_model,
                       fetch_claude_models, test_claude,
                       load_openai_key, save_openai_key, load_openai_model, save_openai_model,
                       fetch_openai_models, test_openai,
                       load_opencode_key, save_opencode_key, load_opencode_model, save_opencode_model,
                       OPENCODE_PLANS, load_opencode_plan, save_opencode_plan,
                       fetch_opencode_models, save_opencode_model_info, test_opencode)

MENU_ORDER = ('opencode', 'claude', 'openai', 'claude_code')
TOOLTIP = 'AI 연결 바꾸기 (OpenCode · Claude · ChatGPT · Claude 구독)'

# AI 제공자별 설정 정보
PROVIDERS = {
    'opencode': {'desc': '해당 API키 필요'},
    'claude_code': {'desc': '유료 구독 필요'},
    'claude':   {'desc': '해당 API키 필요',
                 'key_label': 'Anthropic(Claude) API 키를 입력하세요.',
                 'load_key': load_claude_key, 'save_key': save_claude_key,
                 'load_model': load_claude_model, 'save_model': save_claude_model,
                 'fetch': fetch_claude_models, 'test': test_claude},
    'openai':   {'desc': '해당 API키 필요',
                 'key_label': 'OpenAI(ChatGPT) API 키를 입력하세요.',
                 'load_key': load_openai_key, 'save_key': save_openai_key,
                 'load_model': load_openai_model, 'save_model': save_openai_model,
                 'fetch': fetch_openai_models, 'test': test_openai},
}


class _Notifier(QObject):
    changed = pyqtSignal()


_notifier = None


def _get_notifier():
    """알림 객체 (처음 쓸 때 만들고, Qt가 지웠으면 다시 만든다)."""
    global _notifier
    try:
        from PyQt5 import sip
    except ImportError:          # 아주 옛 PyQt5
        import sip
    if _notifier is None or sip.isdeleted(_notifier):
        _notifier = _Notifier()
    return _notifier


def on_change(slot):
    """AI 연결이 바뀔 때 slot()을 부른다."""
    _get_notifier().changed.connect(slot)


def notify_changed():
    _get_notifier().changed.emit()


def status():
    """버튼 표시 → (글, 상태). 상태: 'on' 연결됨 · 'legacy' 다시 연결 필요 · 'off' 꺼짐."""
    mode = load_ai_mode()
    if mode != 'none':
        return f'{PROVIDER_LABELS[mode]} · {load_external_model_name()}', 'on'
    if legacy_ai_name():
        return 'AI 다시 연결 필요', 'legacy'
    return 'AI 꺼짐 · 연결하기', 'off'


def show_menu(parent, anchor, style='', extra=()):
    """연결 메뉴를 anchor 버튼 아래에 띄운다. extra: [(key, 글)] — 창마다 덧붙이는 항목.
    고른 것: 'opencode' 등 제공자 · 'none'(끄기) · extra의 key · None(닫음)."""
    mode = load_ai_mode()
    menu = QMenu(parent)
    if style:
        menu.setStyleSheet(style)
    title = menu.addAction('AI 연결')
    title.setEnabled(False)
    acts = {}
    for key in MENU_ORDER:
        mark = '✓  ' if mode == key else '     '
        acts[menu.addAction(f'{mark}{PROVIDER_LABELS[key]}\t{PROVIDERS[key]["desc"]}')] = key
    menu.addSeparator()
    acts[menu.addAction(('✓  ' if mode == 'none' else '     ') + 'AI 끄기')] = 'none'
    if extra:
        menu.addSeparator()
        for key, text in extra:
            acts[menu.addAction(text)] = key
    action = menu.exec_(anchor.mapToGlobal(anchor.rect().bottomLeft()))
    return acts.get(action)


def turn_off():
    save_ai_mode('none')
    notify_changed()


def connect(parent, mode, say=None):
    """mode 제공자를 연결한다. say(글): 진행·결과 안내 (상태 표시줄 등). 연결되면 True."""
    say = say or (lambda _m: None)
    if mode == 'claude_code':
        ok = _connect_claude_code(parent, say)
    elif mode == 'opencode':
        ok = _connect_opencode(parent, say)
    elif mode in ('claude', 'openai'):
        ok = _connect_keyed(parent, mode, say)
    else:
        return False
    if ok:
        notify_changed()
    return ok


def _busy(say, msg):
    say(msg)
    QApplication.setOverrideCursor(Qt.WaitCursor)
    QApplication.processEvents()


def _idle():
    QApplication.restoreOverrideCursor()


def _connect_keyed(parent, mode, say):
    """Claude·ChatGPT: API 키 → 모델(서버의 최신 목록) → 연결 시험 → 저장."""
    from document_editor import ApiKeyDialog, _ModelSelectDialog
    p = PROVIDERS[mode]
    label = PROVIDER_LABELS[mode]
    dlg = ApiKeyDialog(f'{label} 연결', p['key_label'], p['load_key'](), parent)
    if dlg.exec_() != dlg.Accepted:
        return False
    key = dlg.key()
    if not key:
        say(f'{label} API 키가 비어 있습니다.')
        return False

    _busy(say, f'{label} 모델 목록을 받는 중...')
    try:
        models = p['fetch'](key)
    finally:
        _idle()
    dlg2 = _ModelSelectDialog(f'{label} 모델 선택', models, p['load_model'](), parent,
                              subtitle='목록 위쪽이 최신 모델입니다. 공문 작성에는 빠르고 저렴한 모델로 충분해요.')
    if dlg2.exec_() != dlg2.Accepted or not dlg2.selected_model():
        return False
    model = dlg2.selected_model()

    _busy(say, f'{label} · {model} 연결 확인 중...')
    try:
        ok, msg = p['test'](key, model)
    finally:
        _idle()
    if not ok:
        say(f'{label} 연결 실패')
        QMessageBox.warning(parent, f'{label} 연결 실패', msg)
        return False
    p['save_key'](key, persist=dlg.should_save())
    p['save_model'](model)
    save_ai_mode(mode)
    say(f'{label} · {model} 연결 완료.')
    return True


def _connect_claude_code(parent, say):
    """Claude 구독: 이 PC의 Claude Code로 구독 사용 (설치·로그인·확인·저장은 대화상자 안에서)."""
    from document_editor import ClaudeCodeDialog
    dlg = ClaudeCodeDialog(parent)
    if dlg.exec_() != QDialog.Accepted:
        return False
    say(f'Claude 구독 · {load_external_model_name()} 연결 완료.')
    return True


def _connect_opencode(parent, say):
    """OpenCode: 요금제 → API 키 → 모델(서버의 최신 목록) → 연결 시험 → 저장."""
    from document_editor import ApiKeyDialog, _ModelSelectDialog
    plan_keys = list(OPENCODE_PLANS)
    labels = [OPENCODE_PLANS[k]['label'] for k in plan_keys]
    label, ok = QInputDialog.getItem(parent, 'OpenCode 연결', '사용 중인 OpenCode 요금제를 고르세요:',
                                     labels, plan_keys.index(load_opencode_plan()), False)
    if not ok:
        return False
    plan = plan_keys[labels.index(label)]

    dlg = ApiKeyDialog('OpenCode 연결', 'OpenCode API 키를 입력하세요.', load_opencode_key(), parent)
    if dlg.exec_() != dlg.Accepted:
        return False
    key = dlg.key()
    if not key:
        say('OpenCode API 키가 비어 있습니다.')
        return False

    _busy(say, 'OpenCode 모델 목록을 받는 중...')
    try:
        models = fetch_opencode_models(plan)
    finally:
        _idle()

    def mark(m):
        tags = []
        if m['id'] == 'claude-haiku-5-5':
            tags.append('추천')
        if m['image']:
            tags.append('🖼 이미지')
        if m['id'].startswith('deepseek'):
            tags.append('⚠ 지역 설정 필요')     # OpenCode 기본(유럽·미국)에서는 막혀 있음
        return m['id'] + ('   ' + ' · '.join(tags) if tags else '')
    shown = [mark(m) for m in models]
    current = next((mark(m) for m in models if m['id'] == load_opencode_model()), '')
    dlg2 = _ModelSelectDialog('OpenCode 모델 선택', shown, current, parent,
                              subtitle='🖼 표시가 있는 모델은 캡처한 공문 이미지도 읽을 수 있어요.')
    if dlg2.exec_() != dlg2.Accepted:
        return False
    picked = dlg2.selected_model().split()[0] if dlg2.selected_model() else ''
    if not picked:
        return False

    if any(m['image'] is not None for m in models):      # 서버 목록을 받았을 때만 (요청 방식·이미지 지원)
        save_opencode_model_info(models)
    _busy(say, f'OpenCode · {picked} 연결 확인 중...')
    try:
        ok, msg = test_opencode(plan, key, picked)
    finally:
        _idle()
    if not ok:
        say('OpenCode 연결 실패')
        QMessageBox.warning(parent, 'OpenCode 연결 실패', msg)
        return False

    save_opencode_key(key, persist=dlg.should_save())
    save_opencode_plan(plan)
    save_opencode_model(picked)
    save_ai_mode('opencode')
    say(f'OpenCode · {picked} 연결 완료.')
    return True
