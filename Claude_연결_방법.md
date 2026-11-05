# Claude와 서서니 노트 연결하기 (MCP)

Claude 구독(Pro/Max)으로 **Claude Code**나 **Claude 데스크톱 앱**에서 서서니 노트의 업무 목록을 보고, 업무를 추가하고, 공문 초안을 쓸 수 있습니다.
**API 키는 필요 없고**, 서서니 노트 데이터는 PC 안에 그대로 있습니다.

## 할 수 있는 것

| Claude에게 이렇게 말하면 | Claude가 쓰는 기능 |
|---|---|
| "오늘 할 일 정리해줘", "이번 주 마감 뭐 있어?" | 오늘 한눈에 보기 |
| "방과후 관련 업무 찾아줘", "ㄱㅁㅎㅇ 업무 있어?" | 업무 찾기 (초성 검색 가능) |
| "다음 주 금요일 5시까지 교과협의회 계획서 제출, 업무로 넣어줘" | 업무 추가 |
| "교무회의 자료 제출 완료 처리해줘", "마감을 15일로 미뤄줘" | 업무 수정 |
| "학교폭력 실태조사 결과 보고 공문 초안 써줘" | 공문 작성 안내 (학교 서식 + 학교 정보) |

- 추가·수정한 업무는 **켜져 있는 메모 창에 3초 안에 나타납니다.**
- **삭제는 할 수 없습니다** (실수 방지). 삭제는 메모 창에서 직접 해주세요.
- Claude는 추가·수정 전에 이름과 마감일을 한 번 확인합니다.

## 준비물

1. **Claude Pro 또는 Max 구독**
2. **Claude Code** 또는 **Claude 데스크톱 앱** (둘 중 하나)
3. **Python 3.10 이상** — 연결 프로그램(`mcp_server.py`)을 실행하는 데 필요합니다.
   - 확인: 명령 프롬프트에서 `python --version`
4. 연결용 패키지 설치 (서서니 노트 폴더에서 한 번만):

```bash
pip install -r requirements-mcp.txt
```

## 등록 — Claude Code

`C:\경로\SSNnote` 부분을 서서니 노트 폴더 경로로 바꿔서 한 번 실행합니다.
`--scope user`를 붙이면 어느 폴더에서 Claude Code를 켜도 연결됩니다.

```bash
claude mcp add --scope user ssnnote -- python "C:\경로\SSNnote\mcp_server.py"
```

연결 확인 (`ssnnote: ... √ Connected`가 보이면 성공):

```bash
claude mcp list
```

`python` 명령이 안 되면 `python` 대신 `py`를 넣어보세요.

## 등록 — Claude 데스크톱 앱

1. Claude 앱 → **설정 → 개발자 → 설정 편집(Edit Config)** 을 누르면 `claude_desktop_config.json` 파일이 열립니다.
2. 아래 내용을 넣고 저장합니다. 경로의 `\`는 **두 번씩** 씁니다. 이미 `mcpServers`가 있으면 그 안에 `"ssnnote": {...}`만 추가합니다.

```json
{
  "mcpServers": {
    "ssnnote": {
      "command": "python",
      "args": ["C:\\경로\\SSNnote\\mcp_server.py"]
    }
  }
}
```

3. Claude 앱을 **완전히 종료했다가 다시 켭니다** (트레이 아이콘까지 종료).

## 개인정보 안내

- Claude에게 업무를 물어보면, **그 업무 내용(이름·마감·메모)이 Claude 대화에 들어갑니다.**
- 업무 메모에 학생 이름·연락처 같은 개인정보는 되도록 적지 마세요.
- 연결 프로그램은 PC 안의 서서니 노트 DB(`%APPDATA%\SSNnote\ssnnote.db`)만 읽고 쓰며, 다른 곳으로 데이터를 보내지 않습니다.

## 문제 해결

| 증상 | 확인할 것 |
|---|---|
| `claude mcp list`에 연결 실패 | `python "C:\경로\SSNnote\mcp_server.py"`를 직접 실행해 오류가 나는지 확인 (정상이면 아무 출력 없이 대기 — Ctrl+C로 종료) |
| `No module named 'mcp'` | `pip install -r requirements-mcp.txt` 다시 실행 |
| 추가한 업무가 메모 창에 안 보임 | 서서니 노트를 최신 버전으로 실행 중인지 확인 (3초마다 새로고침) |
| Claude가 날짜를 잘못 넣음 | "마감은 2026-11-14 17:00"처럼 날짜를 정확히 말해주세요 |
