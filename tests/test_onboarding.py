"""첫 사용 안내(코치마크): 처음에만 · 업무를 넣어야 다음 · 파일 놓으면 문구 바뀜 · 건너뛰기/끝."""
import sys

import pytest

pytest.importorskip('PyQt5')
_APP = None


@pytest.fixture
def win(tmp_path, monkeypatch):
    from unittest.mock import MagicMock
    sys.modules.setdefault('winreg', MagicMock())
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    monkeypatch.setenv('APPDATA', str(tmp_path))
    db.init_db()
    from PyQt5.QtWidgets import QApplication
    global _APP
    _APP = app = QApplication.instance() or QApplication([])
    from ui import ai_chat_window as acw
    monkeypatch.setattr(acw, 'auto_open', lambda h: None)
    import window
    w = window.MemoWindow(db.get_first_window_id(), open_windows=[])
    w.resize(420, 460)
    w.show()
    app.processEvents()
    yield w, app
    if getattr(w, '_tour', None) is not None:
        w._tour.finish(mark=False)
    w.close()


def _pump(app, n=5):
    for _ in range(n):
        app.processEvents()


def _finish_step(t):
    """퀘스트를 다 했으면 완료 연출을 건너뛰고 다음 단계로."""
    t._check()
    assert t._burst is not None, f'완료되지 않음: {t._key()} {t.done}'
    t._burst -= 10_000
    t._tick()


def test_full_tour(win, monkeypatch):
    w, app = win
    import PyQt5.QtWidgets as qw
    from ui import onboarding
    t = onboarding.maybe_start(w)
    assert t is not None and w._tour is t and t.intro              # 먼저 유주 인사
    t.begin()
    assert t.step == 0 and not t.intro
    assert t.btn_action.isHidden()                                 # [다음] 같은 단추 없음 — 직접 해야
    # 1 업무 추가: 쓰기 → Tab(연 · 월 · 일 → 시 · 분 → 반복) → Enter, 칸마다 유주가 설명
    from PyQt5.QtWidgets import QDateTimeEdit
    w.input_name.setText('체험학습 계획서 제출')
    t._check()
    assert t.done == [False] * 5                                    # 치자마자는 아직
    t._t += onboarding.TYPE_PAUSE_MS
    t._check()
    assert t.done == [True, False, False, False, False]             # 0.7초 멈추면 ✓
    focus = {'w': w.input_name}
    monkeypatch.setattr(qw.QApplication, 'focusWidget', staticmethod(lambda: focus['w']))
    t._check()
    assert '마감일' in t._line and 'Tab' in t._line
    focus['w'] = w.input_date
    sec = {'d': QDateTimeEdit.YearSection, 't': QDateTimeEdit.HourSection}
    monkeypatch.setattr(w.input_date, 'currentSection', lambda: sec['d'])
    monkeypatch.setattr(w.input_time, 'currentSection', lambda: sec['t'])
    t._check()
    assert '연도' in t._line and not t.done[1]
    sec['d'] = QDateTimeEdit.MonthSection
    t._check()
    assert t._line.startswith('월') and '↑↓' in t._line
    sec['d'] = QDateTimeEdit.DaySection
    t._check()
    assert t._line.startswith('일') and t.done[1]                    # 연 · 월 · 일 ✓
    focus['w'] = w.input_time
    t._check()
    assert '24시간' in t._line and not t.done[2]
    sec['t'] = QDateTimeEdit.MinuteSection
    t._check()
    assert t._line.startswith('분') and t.done[2]
    focus['w'] = w.input_recurrence
    t._check()
    assert '매주' in t._line and t.done == [True, True, True, True, False]   # 반복까지 ✓
    w._add_task()
    _pump(app)
    _finish_step(t)
    assert t._key() == 'row' and t._row().task['name'] == '체험학습 계획서 제출'
    # 2 업무 다루기: 별 → 메모 → ⋮ 메뉴
    import db
    db.update_task(t.task_id, '체험학습 계획서 제출', '', 0, 1, '')
    t._check()
    assert t.done[0] and not t.done[1]
    row = t._row()
    row._expanded = True
    t._check()
    assert t.done[1]
    assert t._target() == t._rect(row.btn_menu)                     # ⋮ 단추만 밝게
    t._menu_seen = True
    _finish_step(t)
    # 3 파일
    assert t._key() == 'file'
    t._row()._attachments.append('C:/문서/계획서.hwp')
    _finish_step(t)
    # 4 아래 줄 버튼: 열기 → 닫기
    assert t._key() == 'bottom'
    state = {'open': True}
    monkeypatch.setattr(t, '_bottom_open', lambda k: state['open'])
    t._check()
    assert t.done == [True, False]
    state['open'] = False
    _finish_step(t)
    # 5 단축키: Ctrl+Shift+X → Ctrl+Shift+E → Ctrl+K (순서대로)
    assert t._key() == 'keys'
    onboarding.notify('restore')                                    # 순서가 앞서면 아직
    t._check()
    assert t.done == [False, False, False]
    onboarding.notify('capture')
    t._check()
    assert t.done == [True, True, False]                            # 앞서 누른 E도 이제 인정
    onboarding.notify('palette')
    _finish_step(t)
    # 6 AI: 먼저 장점 카드 7장 → (연결된 PC) 대화창 열기 → 「…」 → 설정 → AI 연결 → 안내 끝까지
    import ai_client
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'claude_code')
    assert t._key() == 'ai' and t._ai_page == 0
    t._check()
    assert t._burst is None and t._target().isEmpty()
    for _ in range(len(onboarding.AI_PAGES)):
        t.btn_action.click()
    assert t._ai_page is None
    t._check()
    assert t._burst is None and not t.done[0] and 'Alt+4' in t.quests()[0]
    assert t._target() == t._rect(w.btn_ai_bottom)
    monkeypatch.setattr(t, '_bottom_open', lambda k: k == 'ai')
    t._check()
    assert t.done == [True, False, False] and t._target() == t._rect(w.title_bar.btn_menu)
    from ui.ai_settings_window import AiSettingsWindow
    dlg = AiSettingsWindow(w)
    dlg.show()
    _pump(app)
    assert dlg.guide is not None
    t._check()
    assert t.done == [True, True, False] and t._line == onboarding.AI_LINES['guide']
    for _ in range(6):
        dlg.guide.btn_next.click()                                   # 마지막 [다 봤어요] → 창이 닫힘
    _pump(app)
    assert not dlg.isVisible()
    _finish_step(t)
    # 6 메뉴 → 클리어
    assert t._key() == 'menu'
    t._menu_seen = True
    _finish_step(t)
    assert t.clear and t.btn_action.text() == '시작하기' and len(t.completed) == 7
    assert not onboarding.is_done()
    t.btn_action.click()
    _pump(app)
    assert onboarding.is_done() and w._tour is None
    assert onboarding.maybe_start(w) is None                       # 한 번만


def test_enter_without_tab_needs_another(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    w.input_name.setText('하나')
    w._add_task()
    _pump(app)
    t._check()
    assert t.step == 0 and t.done == [True, False, False, False, False] and t._burst is None


def test_skip_step_and_ai_later(win, monkeypatch):
    w, app = win
    import ai_client
    from ui import onboarding
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    t = onboarding.start(w)
    t.begin()
    for _ in range(5):
        t.btn_skip.click()
    assert t._key() == 'ai'
    t._layout_buttons()
    assert t.btn_action.text() == '다음' and t.btn_later.isHidden() and t.btn_skip.isHidden()   # 장점 카드 1장
    t.btn_action.click()
    assert t._ai_page == 1 and t.btn_later.text() == '이전' and not t.btn_later.isHidden()
    t.btn_later.click()                                               # [이전]
    assert t._ai_page == 0 and t._key() == 'ai'
    for _ in range(len(onboarding.AI_PAGES) - 1):
        t.btn_action.click()
    assert t.btn_action.text() == '연결하러 가기'
    t.btn_action.click()
    assert t._ai_page is None and t.btn_action.isHidden()
    assert t.btn_later.text() == '다음에 할게요' and t.btn_skip.isHidden()
    assert t._notes() == [onboarding.AI_NOTE] and t._info() == []
    assert t.quests() == [onboarding.AI_QUESTS['settings'], onboarding.AI_QUESTS['guide']]
    t.btn_later.click()
    assert t._key() == 'menu'


def test_hole_lets_clicks_through(win):
    w, app = win
    from PyQt5.QtCore import QPoint
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    t._tick()
    c = t._hole.center()
    assert not t.mask().contains(c)                          # 뚫린 곳은 아래 칸이 받음
    assert t.mask().contains(QPoint(2, t.height() - 2))


def test_skip_marks_done(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    t.btn_quit.click()
    _pump(app)
    assert onboarding.is_done() and w._tour is None


def test_existing_user_not_shown(win):
    w, app = win
    import db
    from ui import onboarding
    db.add_task(w.window_id, '쓰던 업무', '')
    assert onboarding.maybe_start(w) is None and onboarding.is_done()


def test_replay_from_menu(win):
    w, app = win
    import db
    db.add_task(w.window_id, '쓰던 업무', '')
    w._refresh_tasks()
    _pump(app)
    w.start_tour()
    t = w._tour
    t.begin()
    assert t.step == 0 and t.done == [False] * 5                   # 다시 볼 때도 직접 해야


def test_typing_again_restarts_pause(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    w.input_name.setText('체')
    t._check()
    t._t += 500
    w.input_name.setText('체험')                                   # 다시 치면 처음부터
    t._check()
    t._t += 500
    t._check()
    assert not t.done[0]                                            # 다시 친 뒤 0.5초 — 아직
    t._t += 300
    t._check()
    assert t.done[0]                                                # 0.8초 — ✓


def test_card_goes_beside_window_when_room(win, monkeypatch):
    w, app = win
    from PyQt5.QtCore import QPoint
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    monkeypatch.setattr(t, '_side_spot', lambda pw, ph: QPoint(900, 100))
    t._tick()
    assert t._outside and t.btn_quit.parent() is t.panel and t._card.left() < 20
    monkeypatch.setattr(t, '_side_spot', lambda pw, ph: None)      # 자리가 없으면 창 안쪽
    t._tick()
    assert not t._outside and t.btn_quit.parent() is t and t.panel.isHidden()


def test_intro_waits_for_start_button(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    assert t.intro and t.btn_action.text() == '시작하기' and t.btn_quit.isHidden() and t.btn_skip.isHidden()
    assert t._typed == len(t._line)                                 # 글은 한 번에
    t._t += 60_000
    t._tick()
    assert t.intro                                                  # 기다려도 저절로 넘어가지 않음


def test_intro_start_button(win):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.btn_action.click()
    assert not t.intro and t._key() == 'add'


# ── v3.73: 자세한 Tab 설명 · ⋮ 강조 · 파일 놓는 칸 · Alt+1 표시 · AI 설명 ──
def _tour_with_task(w, app):
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    w.input_name.setText('계획서')
    w._add_task()
    _pump(app)
    return t


def test_add_step_notes_and_calendar(win, monkeypatch):
    w, app = win
    from ui import onboarding
    t = onboarding.start(w)
    t.begin()
    notes = ' '.join(t._notes())
    assert 'Shift+Tab' in notes and '달력' in notes and 'Esc' in notes
    w.input_name.setText('체험학습')
    t._check()
    t._t += onboarding.TYPE_PAUSE_MS
    t._check()
    assert t.done[0]
    monkeypatch.setattr(w._cal_popup, 'isVisible', lambda: True)    # 마감일 칸을 눌러 달력이 뜸
    t._check()
    assert t._line == onboarding.ADD_LINES['calendar']
    monkeypatch.setattr(w._cal_popup, 'isVisible', lambda: False)   # 날짜를 눌러 골랐음
    w._date_touched = True
    t._check()
    assert t.done[1]                                                # 달력으로 골라도 마감일 ✓
    w._add_task()
    _pump(app)
    t._check()
    assert t._burst is not None                                     # 마감일까지 했으면 Enter로 완료


def test_row_step_shows_controls_and_waits_menu_close(win, monkeypatch):
    w, app = win
    import db
    import PyQt5.QtWidgets as qw
    t = _tour_with_task(w, app)
    t.show_step(1)
    row = t._row()
    db.update_task(row.task['id'], row.task['name'], '', 0, 1, '')
    row._expanded = True
    t._check()
    assert row.btn_menu._row_active                                 # ⋮ 늘 보이게
    t._menu_seen = True
    sentinel = object()
    monkeypatch.setattr(qw.QApplication, 'activePopupWidget', staticmethod(lambda: sentinel))
    t._check()
    assert not t.done[2] and 'C' in t._line and t._burst is None    # 메뉴가 열린 동안은 설명만
    monkeypatch.setattr(qw.QApplication, 'activePopupWidget', staticmethod(lambda: None))
    t._check()
    assert t._burst is not None                                     # 닫은 뒤 ✓
    t.finish(mark=False)
    _pump(app)
    assert not row.btn_menu._row_active                             # 끝내면 원래대로 (마우스를 올렸을 때만)


def test_file_step_drop_zone(win):
    w, app = win
    from PyQt5.QtCore import Qt
    t = _tour_with_task(w, app)
    t.show_step(2)
    t._tick()
    dz = t.drop_zone
    assert dz.isVisible() and dz.testAttribute(Qt.WA_TransparentForMouseEvents)   # 끌어 놓기는 업무 줄로 통과
    row = t._row()
    assert row.geometry().isValid() and t._rect(row).contains(dz.geometry().center())
    assert '바탕화면' in t._line and t.btn_later.text() == '다음에 할게요'
    t._drag_over = True
    t._tick()
    assert dz.hover
    row._attachments.append('C:/바탕화면/사진.png')
    t._check()
    assert '더블클릭' in t._line                                     # 붙으면 여는 법 안내
    t._burst -= 10_000
    t._tick()
    assert t._key() == 'bottom' and not dz.isVisible()


def test_bottom_step_hole_first_button_and_panel_on_top(win):
    w, app = win
    from PyQt5.QtCore import Qt
    t = _tour_with_task(w, app)
    t.show_step(3)
    k, label = t._bottom_first()
    assert t._target() == t._rect(w._bottom_btns[k])                # 첫 번째 단추만
    assert '을(를)' not in t._line and f'{label} 창' in t._line
    assert t._keep_panel_top                                        # 카드를 다른 창 위로 계속 올림
    assert not (t.panel.windowFlags() & Qt.WindowStaysOnTopHint)    # 창 속성은 안 바꿈 (바꾸면 드물게 튕김)
    t.show_step(4)
    assert not t._keep_panel_top


def test_ai_step_not_connected_open_window_counts(win, monkeypatch):
    w, app = win
    import ai_client
    from ui import onboarding
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    from ui.ai_settings_window import AiSettingsWindow
    t = _tour_with_task(w, app)
    t.show_step(onboarding.STEPS.index('ai'))
    t._check()
    assert t._burst is None and t.done == [False, False]
    dlg = AiSettingsWindow(w)                                         # 장점 카드 중에 바로 열어도 됨
    dlg.show()
    _pump(app)
    assert t._ai_page is None and dlg.guide is not None and dlg.guide.page == 0
    assert dlg.guide.ring.isVisible() and not dlg.guide.ring.target.isEmpty()     # 창이 뜨는 도중에 붙여도 테두리가 보임
    t._check()
    assert t.done == [True, False] and t._burst is None
    dlg.close()                                                       # 끝까지 안 보고 닫으면 아직
    _pump(app)
    t._check()
    assert t.done == [True, False] and t._burst is None
    dlg = AiSettingsWindow(w)
    dlg.show()
    _pump(app)
    for _ in range(5):
        dlg.guide.btn_next.click()
    assert dlg.guide.btn_next.text() == '다 봤어요'
    dlg.close()                                                       # 마지막 장까지 봤으면 [닫기]로 닫아도 ✓
    _pump(app)
    t._check()
    assert t._burst is not None


def test_tour_cycles_no_crash(win):
    """단계를 앞뒤로 오가며 그리고 · 끝내기를 반복해도 오류 없음."""
    w, app = win
    from ui import onboarding
    for _ in range(2):
        t = _tour_with_task(w, app)
        for i in range(len(onboarding.STEPS)):
            t.show_step(i)
            for _ in range(3):
                t._tick()
            t.update()
            t.panel.update()
            _pump(app)
            w.grab()
        t.finish(mark=False)
        _pump(app)


def test_ai_settings_guide_pages_and_rings(win, monkeypatch):
    """AI 연결 창 안내 띠: ①~⑥ 장마다 강조할 칸이 있고, 튜토리얼이 아닐 땐 띠가 없다."""
    w, app = win
    import ai_client
    from ui import ai_settings_window as asw
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    plain = asw.AiSettingsWindow(w)
    plain.show()
    _pump(app)
    assert plain.guide is None
    plain.close()
    dlg = asw.AiSettingsWindow(w)
    g = dlg.start_guide()
    dlg.show()
    _pump(app)
    assert not g.btn_prev.isEnabled()
    for i in range(len(asw.GUIDE)):
        g.go(i)
        _pump(app)
        assert g.count.text() == f'{i + 1} / {len(asw.GUIDE)}'
        assert not g.ring.target.isEmpty(), asw.GUIDE[i][2]
        assert dlg.rect().contains(g.ring.target.center())
        assert g.text.height() >= g.text.heightForWidth(g.text.width()) - 1   # 설명이 잘리지 않음
        dlg.grab()
    assert g.seen_all and g.btn_next.text() == '다 봤어요'
    g.btn_prev.click()
    assert g.page == len(asw.GUIDE) - 2
    dlg.close()


def test_ai_benefit_pages_paint(win, monkeypatch):
    """장점 카드 7장을 모두 그려도 오류 없음 · 장마다 유주 말이 바뀜."""
    w, app = win
    import ai_client
    from ui import onboarding
    monkeypatch.setattr(ai_client, 'load_ai_mode', lambda: 'none')
    t = _tour_with_task(w, app)
    t.show_step(onboarding.STEPS.index('ai'))
    seen = set()
    for i in range(len(onboarding.AI_PAGES)):
        assert t._ai_page == i and t._line == onboarding.AI_PAGES[i]['say']
        seen.add(t._line)
        for _ in range(2):
            t._tick()
        t.update(); t.panel.update()
        _pump(app)
        w.grab(); t.panel.grab()
        assert t._card_height() > 0
        t.btn_action.click()
    assert len(seen) == len(onboarding.AI_PAGES) and t._ai_page is None
