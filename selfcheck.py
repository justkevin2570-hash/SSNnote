"""selfcheck.py — 빌드된 프로그램 자체 점검 (`SSNnote.exe --self-check 결과파일`).

자동 빌드가 릴리스 전에 실행한다. 빌드에 빠진 모듈이 있으면 실제 사용자가 창을 열 때에야
드러나므로(예: v2.68의 numpy 누락), 여기서 주요 창을 실제로 만들어 본다.
사용자 데이터는 건드리지 않도록 임시 폴더를 APPDATA로 쓴다. 결과는 파일에 쓴다(창 프로그램이라 화면 출력이 없음).
"""

import os
import sys
import tempfile
import traceback


def run(result_path: str) -> int:
    lines, ok = [], True

    def check(name, fn):
        nonlocal ok
        try:
            fn()
            lines.append(f'OK   {name}')
        except Exception:
            ok = False
            lines.append(f'FAIL {name}\n' + traceback.format_exc())

    os.environ['APPDATA'] = tempfile.mkdtemp(prefix='ssnnote_selfcheck_')

    def numpy_ok():
        import numpy
        assert hasattr(numpy, 'ndarray'), 'numpy가 불완전하게 들어 있음'

    def modules_ok():
        import db, dashboard, doc_ai, privacy, school_profile, korean_search  # noqa: F401
        import claude_connect, ai_client, updater, capture  # noqa: F401
        db.init_db()

    check('numpy', numpy_ok)
    check('기본 모듈', modules_ok)

    try:
        from PyQt5.QtWidgets import QApplication
        if QApplication.instance() is None:
            from ui.common import setup_high_dpi
            setup_high_dpi()        # 실제 앱과 같은 화면 배율 설정으로 점검
        app = QApplication.instance() or QApplication(['SSNnote'])

        def windows_ok():
            import db
            import window
            import document_editor
            import budget
            from ui.history_dialog import HistoryDialog
            w = window.MemoWindow(db.get_first_window_id())
            e = document_editor.DocumentEditorWindow()
            b = budget.BudgetWindow()
            h = HistoryDialog(w, w.window_id, lambda r: None)
            h.show_tab(1)
            for x in (h, w, e, b):
                x.close()
            app.processEvents()

        def dialogs_ok():
            from ui.dashboard_dialog import DashboardDialog
            DashboardDialog().close()

        def capture_panel_ok():
            import document_editor
            from PyQt5.QtGui import QPixmap
            e = document_editor.DocumentEditorWindow()
            e._show_panel(True)
            e.panel.fill({'title': '점검', 'doc_number': '중등교육과-1(2026. 1. 1.)', 'summary': ['a']})
            e.panel._pixmaps = [QPixmap(10, 10), QPixmap(10, 20)]
            assert e.panel._combined_pixmap().height() == 46, '캡처 이어 붙이기 실패'
            assert e._received_context(), '받은 공문 칸 내용이 AI에 전달되지 않음'
            assert 'a' in e.panel.summary(), '요약 칸이 비어 있음'
            e.ref_input.setText('중등교육과-1(2026. 1. 1.)')
            e._copy_ref()
            e.close()

        def motion_ok():
            """급함 표시가 실제로 움직이는지: 오늘 마감 업무의 왼쪽 세로선이 시간이 흐르며 다시 그려져야 한다."""
            import time
            from datetime import date
            import db
            import window
            from ui import motion
            from ui.task_row import TaskRow
            wid = db.get_first_window_id()
            tid = db.add_task(wid, '점검용 오늘 업무', date.today().isoformat())
            w = window.MemoWindow(wid)
            w.show()
            row = next(r for r in w._task_rows() if r.task['id'] == tid)
            assert row._urg.get('breathe') and row._line, f'오늘 업무에 세로선이 없음: {row._urg}'
            assert not row._dday_lbl._breathe, 'D-day 표시는 움직이지 않아야 함'
            count = {'n': 0}
            orig = TaskRow.paintEvent

            def counting(self, e):
                if self is row:
                    count['n'] += 1
                orig(self, e)
            TaskRow.paintEvent = counting
            try:
                end = time.time() + 0.8
                while time.time() < end:
                    app.processEvents()
                    time.sleep(0.01)
            finally:
                TaskRow.paintEvent = orig
            assert motion.clock()._timer.isActive(), '움직임 시계가 멈춰 있음'
            assert count['n'] >= 5, f'세로선이 다시 그려지지 않음 ({count["n"]}번)'
            w.close()
            db.delete_task(tid)

        def extras_ok():
            """새 화면들이 설치 파일에 들어 있고 뜨는지: 사용법 안내 · 버튼 순서 · AI 연결 창."""
            import db
            import window
            from ui import onboarding, bottom_bar
            from ui.ai_settings_window import AiSettingsWindow
            w = window.MemoWindow(db.get_first_window_id())
            w.show()
            t = onboarding.start(w)
            assert t.step == 0, '사용법 안내가 시작되지 않음'
            t.finish(mark=False)
            bottom_bar.BottomOrderDialog(w).close()
            AiSettingsWindow(w).close()
            w.close()
            app.processEvents()

        check('메모 창·지난 기록·공문 작성 창·예산 창', windows_ok)
        check('급한 업무 움직임 (왼쪽 세로선)', motion_ok)
        check('대시보드', dialogs_ok)
        check('공문 작성 창의 받은 공문 칸', capture_panel_ok)
        check('사용법 안내·버튼 순서·AI 연결 창', extras_ok)

        # 글자 인식 진단 (결과만 기록, 실패해도 점검은 통과): 실제 캡처처럼 화면에 그린 글자를 단계별로 읽어 본다
        try:
            import time
            import capture
            from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont
            pm = QPixmap(360, 40)
            pm.fill(QColor('white'))
            pnt = QPainter(pm)
            f = QFont('Malgun Gothic')
            f.setPixelSize(13)
            pnt.setFont(f)
            pnt.setPen(QColor('black'))
            pnt.drawText(6, 26, 'Notice 2026 school-12873 (2026. 11. 18.)')
            pnt.end()
            try:
                from winrt.windows.media.ocr import OcrEngine
                lines.append(f'INFO OCR 최대 크기: {OcrEngine.max_image_dimension}')
            except Exception as e:
                lines.append(f'INFO OCR 최대 크기 확인 실패: {e!r}')
            try:
                full = capture.grab_fullscreen()
                lines.append(f'INFO 화면 캡처: {full.width()}x{full.height()} dpr={full.devicePixelRatio()} '
                             f'화면={[(sc.geometry().getRect(), sc.devicePixelRatio()) for sc in app.screens()]}')
            except Exception as e:
                lines.append(f'INFO 화면 캡처 오류: {e!r}')
            variants = capture.ocr_variants(pm)
            for i, png in enumerate(variants):
                t0 = time.time()
                try:
                    txt = capture._winrt_ocr(png)
                    lines.append(f'INFO OCR 판{i} ({len(png)}B, {time.time() - t0:.1f}s): {txt!r}')
                except Exception as e:
                    lines.append(f'INFO OCR 판{i} 오류: {e!r}')
            got = []
            w = capture.OcrWorker(variants)
            w.result.connect(lambda t, e: got.append((t, e)))
            w.run()
            lines.append(f'INFO OCR 최종: {got!r}')
            old = pm.scaled(pm.width() * 3, pm.height() * 3)
            try:
                lines.append(f'INFO OCR 예전 방식: {capture._winrt_ocr(capture._pixmap_to_png_bytes(old))!r}')
            except Exception as e:
                lines.append(f'INFO OCR 예전 방식 오류: {e!r}')
        except Exception:
            lines.append('INFO OCR 진단 오류\n' + traceback.format_exc())
    except Exception:
        ok = False
        lines.append('FAIL Qt 시작\n' + traceback.format_exc())

    lines.append('RESULT ' + ('OK' if ok else 'FAIL'))
    with open(result_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    return 0 if ok else 1
