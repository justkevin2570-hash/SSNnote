"""ui/theme.py — 디자인 토큰 (색상·글꼴의 단일 출처).

화면 코드에 색상 hex 값을 직접 쓰지 말고 여기 토큰을 쓴다.
  - 정적 문자열:   QColor(T.ACCENT)
  - 스타일시트:    f"border: 1px solid {T.ACCENT};"   (f-string 안에서 QSS 중괄호는 {{ }})

토큰 값은 v2.66 화면과 똑같이 맞춰 두었다. 다크모드·테마 프리셋은
이 값들을 바꿔 끼우는 방식으로 추가한다.
"""


class T:
    # ── 메모(노란 스티커) 계열 ─────────────────────────────────────
    ACCENT               = '#d4b800'  # 입력칸 테두리, 포커스, 스크롤바, 주요 버튼
    ACCENT_HOVER         = '#b89e00'
    ACCENT_SOFT          = '#f1e7b3'  # 같은 노랑 계열의 옅은 단계 (집행률 막대 바탕)
    DIVIDER              = '#D4AF37'  # 제목/날짜/하단 구분선
    BORDER_SOFT          = '#c8b87c'  # 옅은 입력 테두리
    INPUT_BG             = '#fffacd'  # 메모 위 입력칸 배경
    SURFACE              = '#ffffff'  # 카드·입력칸 바탕
    TEXT_ON_NOTE         = '#5a4000'  # 노란 바탕 위 기본 글자
    INK                  = '#1a1410'  # 제목 등 가장 진한 글자
    TEXT_WARM            = '#5a4a3a'
    TEXT_WARM_MUTED      = '#7a6a5a'

    # ── 달력·상세 창 (따뜻한 회색 계열) ───────────────────────────
    SURFACE_WARM         = '#e8ddd5'  # 버튼 배경, 구분선
    SURFACE_WARM_HOVER   = '#ddd0c5'
    SURFACE_WARM_PRESSED = '#d0c0b5'
    BORDER_WARM          = '#f0e8e0'  # 칸 테두리
    BG_WARM              = '#faf6f0'  # 창 배경
    BG_WARM_ALT          = '#f5efe8'

    # ── 상태 색 ──────────────────────────────────────────────────
    HIGHLIGHT            = '#e85d3a'  # 오늘, 중요 업무
    HIGHLIGHT_BG         = '#fff3ee'
    DANGER               = '#e74c3c'  # 마감 임박 D-day, 삭제 hover
    SUNDAY               = '#cc0000'
    SATURDAY             = '#0055cc'
    SATURDAY_SOFT        = '#3a8ee8'

    # ── 중립 회색 스케일 (예산·공문 작성 창, Tailwind slate) ───────
    SLATE_50             = '#f8fafc'
    SLATE_100            = '#f1f5f9'
    SLATE_200            = '#e2e8f0'
    SLATE_300            = '#cbd5e1'
    SLATE_400            = '#94a3b8'
    SLATE_500            = '#64748b'
    SLATE_600            = '#475569'
    SLATE_700            = '#334155'
    SLATE_800            = '#1e293b'
    SLATE_900            = '#0f172a'
    GRAY_200             = '#e5e7eb'

    # ── 포인트 색 스케일 (예산·공문 작성 창, Tailwind indigo) ──────
    INDIGO_300           = '#a5b4fc'
    INDIGO_500           = '#6366f1'
    INDIGO_600           = '#4f46e5'
    INDIGO_700           = '#4338ca'
    INDIGO_900           = '#312e81'
    INDIGO_950           = '#1e1b4b'

    # ── 글꼴 ─────────────────────────────────────────────────────
    FONT_UI              = 'Pretendard'
    FONT_FALLBACK        = 'Malgun Gothic'


# 메모 창 배경 색상표 (메뉴의 색상 선택).
PALETTE = [
    '#FDD663', '#FEFFA7', '#CCFF90', '#A8F0E8',
    '#AECBFA', '#D7AEFB', '#FDCFE8', '#E6C9A8', '#AAAAAA', '#DDDDDD', '#FFFFFF',
]
