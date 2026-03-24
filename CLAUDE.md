# Project Context: Tmemo (Python + PyQt5)

## 1. Core Concept
- Windows 전용 경량 스티커 메모 앱 ("서서니 메모").
- 고등학교 행정 업무 및 수학 연구(유자-pi) 관련 빠른 메모 작성에 특화.

## 2. 기술 스택

| 항목 | 내용 |
|------|------|
| **언어** | Python |
| **GUI 프레임워크** | PyQt5 |
| **데이터베이스** | SQLite (Python 내장 `sqlite3` 모듈) |
| **DB 저장 경로** | `%APPDATA%/Tmemo/tmemo.db` |
| **패키징** | PyInstaller (`Tmemo.spec`) |
| **가상환경** | `.venv` (Python venv) |

## 3. 파일 구조

```
main.py       - 진입점: QApplication, 시스템 트레이, 창 관리, 단일 인스턴스 잠금
window.py     - MemoWindow(QMainWindow), EdgeHandle(리사이즈), UI 전체
db.py         - SQLite CRUD (windows / tasks / task_history 테이블)
autostart.py  - Windows 레지스트리 자동 시작 관리 (HKCU\...\Run)
Tmemo.spec    - PyInstaller 빌드 설정
```

## 4. DB 스키마

- **windows**: id, x, y, width, height, collapsed, color
- **tasks**: id, window_id, name, deadline, strikethrough
- **task_history**: id, window_id, name, deadline, strikethrough, cleared_at

## 5. 구현된 주요 기능

- **Window Shade (말아올리기):** 타이틀바 더블클릭 시 창을 타이틀바 높이(40px)로 접음/펼침
- **Always on Top:** 토글로 창을 항상 최상위에 고정
- **프레임리스 UI:** 커스텀 타이틀바, 드래그 이동, 가장자리 리사이즈(EdgeHandle)
- **창 스냅:** 창끼리 가까워지면 자동으로 달라붙는 기능 (SNAP_THRESHOLD=20px)
- **파스텔 색상 팔레트:** 11가지 배경색 지원
- **D-day 표시:** 태스크 마감일 기준 D-day 자동 계산
- **태스크 관리:** 추가/삭제/수정/취소선, 마감일 정렬
- **태스크 기록:** 완료된 태스크를 task_history에 보관
- **다중 창:** 여러 메모 창을 동시에 운용, 각 창 상태를 DB에 저장
- **시스템 트레이:** 트레이 아이콘으로 새 메모/종료, 좌클릭으로 창 복원
- **단일 인스턴스:** 소켓(포트 47391)으로 중복 실행 방지
- **자동 시작:** Windows 레지스트리로 로그인 시 자동 실행 등록/해제
- **창 상태 영속:** 위치/크기/접힘 상태/색상 모두 DB에 저장, 재시작 후 복원

## 6. 코딩 규칙

- DB 로직(`db.py`), 창 UI(`window.py`), 진입점(`main.py`), 자동시작(`autostart.py`) 역할을 명확히 분리 유지.
- UI는 깔끔하고 방해 요소 없이 유지.
- 새 기능 추가 시 DB 마이그레이션 코드를 `init_db()`에 `ALTER TABLE` 형식으로 추가.

## 7. 빌드

```bash
pyinstaller Tmemo.spec
```
결과물: `dist/Tmemo.exe` (단일 실행 파일, 콘솔 없음)

## 8. 향후 로드맵

- 온라인 동기화 모드(로그인 시스템)로 전환 예정이나, 현재는 로컬 SQLite 성능에 집중.
