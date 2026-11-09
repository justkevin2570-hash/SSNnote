"""budget_image.py — 캡처한 예산 표(이미지)를 세부항목별로 정리한다.

예산 시스템 화면의 표(세부항목 | 원가통계비목 | 선택 | 산출내역 | 예산현액 | 예산잔액)를
캡처해 예산 창에 붙여넣으면, 세부항목마다 묶어 예산 창 형식으로 바꾼다.

읽는 방법 (BudgetImageWorker)
  1) AI 연결 + 개인정보 보호 모드 꺼짐 → 이미지를 AI에 보내 표를 JSON으로 받음 (가장 정확)
  2) AI 연결 + 보호 모드 켜짐(또는 AI가 이미지를 못 읽음) → PC의 Windows OCR로 읽은 글자만 AI에 보냄
  3) AI 없음 → Windows OCR 글자의 위치(열)로 직접 나눔

정리 규칙 (group_rows)
  · '세부항목' 칸은 묶음의 첫 줄에만 있다 → 새 이름이 나올 때까지 아래 줄은 같은 세부항목
  · 화면에는 집행액이 없고 예산잔액만 있다 → 집행액 = 예산현액 - 예산잔액 (잔액이 음수면 초과 집행)
  · 합계·소계 줄은 뺀다. 숫자를 못 읽었거나 이상한 줄은 'warn'에 이유를 남긴다 (확인 화면에서 노란색)

UI 없는 함수(build_prompt, parse_rows, group_rows, rows_from_words, words_to_text)는 tests/에서 검증한다.
"""

import json
import re

SYSTEM = (
    "당신은 한국 학교 예산(사업관리카드) 화면을 표로 옮기는 도우미입니다. "
    "보이는 글자와 숫자만 정확히 옮기고, 보이지 않는 내용은 지어내지 않습니다. 반드시 JSON만 출력합니다."
)

_FORMAT = '{"rows": [{"category": "세부항목", "cost_type": "원가통계비목", "description": "산출내역", ' \
          '"budget": 440000, "remaining": 0}]}'

_RULES = """규칙
- 표의 열: 세부항목 | 원가통계비목 | 선택 | 산출내역 | 예산현액 | 예산잔액 (열 이름이 조금 달라도 같은 뜻이면 맞춰 읽기)
- 한 줄(산출내역 하나)마다 rows 하나. 위에서 아래 순서 그대로.
- '세부항목'은 묶음의 첫 줄에만 적혀 있습니다. 세부항목 칸이 빈 줄은 ""로 두세요 (앞 줄 것을 반복하지 마세요).
- '선택' 열(체크 칸)은 무시하세요.
- budget(예산현액), remaining(예산잔액)은 쉼표 없는 정수. 음수(-14,000)는 음수로. 읽을 수 없으면 null.
- 합계·소계 줄, 머리글 줄은 빼세요.
- 칸이 좁아 잘린 글자는 보이는 만큼만 적고 지어내지 마세요."""


def build_prompt(table_text: str = '') -> str:
    """AI에 보낼 요청. table_text가 있으면 이미지 대신 OCR로 읽은 표 글자를 준다."""
    src = (f"아래는 화면의 예산 표를 PC에서 글자로 읽은 것입니다 (칸 사이는 ' | ', 줄 순서는 화면 그대로).\n"
           f"글자 인식 오류(0과 O, 1과 l 등)는 문맥에 맞게 바로잡으세요.\n\n{table_text}\n\n"
           if table_text else "첨부한 이미지는 학교 예산 시스템 화면의 표입니다.\n\n")
    return f"{src}{_RULES}\n\n다음 형식의 JSON만 출력하세요:\n{_FORMAT}"


# ── 숫자 ───────────────────────────────────────────────────────

_NUM_RE = re.compile(r'^[\-−–△▲]?\(?[\d,\.]+\)?$')


def to_amount(value):
    """'440,000' · '-14,000' · '△14,000' · '(14,000)' · 440000 → 정수. 못 읽으면 None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(round(value))
    text = str(value).strip().replace(' ', '').replace('원', '')
    if not text or not _NUM_RE.match(text):
        return None
    neg = text[0] in '-−–△▲' or (text.startswith('(') and text.endswith(')'))
    digits = re.sub(r'[^\d]', '', text)          # 쉼표를 마침표로 잘못 읽어도(1.350.000) 그대로 숫자
    if not digits:
        return None
    n = int(digits)
    return -n if neg else n


# ── AI 응답 → 줄 ────────────────────────────────────────────────

def _extract_json(raw: str):
    raw = (raw or '').strip()
    raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw)
    try:
        return json.loads(raw)
    except Exception:
        pass
    for open_ch, close_ch in (('{', '}'), ('[', ']')):
        a, b = raw.find(open_ch), raw.rfind(close_ch)
        if a != -1 and b > a:
            try:
                return json.loads(raw[a:b + 1])
            except Exception:
                continue
    raise ValueError('AI 응답에서 표를 찾지 못했습니다. 다시 시도해 주세요.')


_SKIP_WORDS = ('합계', '총계', '소계', '누계')


def _clean(v) -> str:
    return re.sub(r'\s+', ' ', str(v or '')).strip()


def normalize_rows(rows) -> list:
    """[{category, cost_type, description, budget, remaining}] — 빈 줄·합계 줄은 뺀다."""
    out = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        row = {
            'category': _clean(r.get('category')),
            'cost_type': _clean(r.get('cost_type')),
            'description': _clean(r.get('description')),
            'budget': to_amount(r.get('budget')),
            'remaining': to_amount(r.get('remaining')),
        }
        label = (row['category'] + row['cost_type'] + row['description']).replace(' ', '')
        if any(label.startswith(w) or label == w for w in _SKIP_WORDS) or \
                (not row['description'] and any(w in label for w in _SKIP_WORDS)):
            continue
        if not row['description'] and row['budget'] is None:
            continue
        out.append(row)
    return out


def parse_rows(raw: str) -> list:
    data = _extract_json(raw)
    rows = data.get('rows') if isinstance(data, dict) else data
    if not isinstance(rows, list):
        raise ValueError('AI 응답에 표 줄(rows)이 없습니다.')
    return normalize_rows(rows)


# ── 줄 → 세부항목별 묶음 ────────────────────────────────────────

def group_rows(rows) -> list:
    """줄 목록 → [{'name': 세부항목, 'items': [{cost_type, description, budget, spent, warn}]}].

    세부항목 칸이 빈 줄은 바로 위 세부항목에 속한다. 같은 이름이 다시 나오면 그 묶음에 이어 붙인다.
    """
    cats, current = [], None
    for r in rows:
        name = (r.get('category') or '').strip()
        if name and (current is None or name != current['name']):
            current = next((c for c in cats if c['name'] == name), None)
            if current is None:
                current = {'name': name, 'items': []}
                cats.append(current)
        if current is None:
            current = {'name': '미분류', 'items': []}
            cats.append(current)
        budget, rem = r.get('budget'), r.get('remaining')
        warn = []
        if budget is None:
            warn.append('예산현액을 읽지 못했어요')
        if rem is None:
            warn.append('예산잔액을 읽지 못해 집행액을 0으로 두었어요')
        elif budget is not None and rem > budget:
            warn.append('잔액이 예산보다 커요 — 숫자를 확인하세요')
        if not r.get('description'):
            warn.append('산출내역이 비어 있어요')
        b = budget or 0
        current['items'].append({
            'cost_type': r.get('cost_type') or '일반',
            'description': r.get('description') or '(내용 없음)',
            'budget': b,
            'spent': (b - rem) if rem is not None else 0,
            'warn': warn,
        })
    return [c for c in cats if c['items']]


# ── Windows OCR 단어(위치) → 표 ──────────────────────────────────

_HEADERS = (('category', '세부항목'), ('cost_type', '원가통계비목'), ('select', '선택'),
            ('description', '산출내역'), ('budget', '예산현액'), ('remaining', '예산잔액'))


def _lines(words):
    """단어 → 줄 (세로 위치가 비슷한 단어끼리). 각 줄은 x 순서."""
    ws = [w for w in words if str(w[0]).strip()]
    if not ws:
        return []
    hs = sorted(w[4] for w in ws)
    tol = max(4.0, hs[len(hs) // 2] * 0.6)
    lines = []
    for w in sorted(ws, key=lambda w: w[2] + w[4] / 2):
        cy = w[2] + w[4] / 2
        if lines and abs(lines[-1]['cy'] - cy) <= tol:
            ln = lines[-1]
            ln['words'].append(w)
            ln['cy'] = sum(x[2] + x[4] / 2 for x in ln['words']) / len(ln['words'])
        else:
            lines.append({'cy': cy, 'words': [w]})
    for ln in lines:
        ln['words'].sort(key=lambda w: w[1])
    return [ln['words'] for ln in lines]


def _header_columns(line):
    """머리글 줄에서 각 열 이름의 가로 가운데 위치. 글자 사이가 벌어져 있어도('예 산 현 액') 찾는다."""
    chars, xs = [], []
    for text, x, _y, w, _h in line:
        t = re.sub(r'\s', '', str(text))
        for i, ch in enumerate(t):
            chars.append(ch)
            xs.append(x + w * (i + 0.5) / max(1, len(t)))
    s = ''.join(chars)
    found = {}
    for key, name in _HEADERS:
        i = s.find(name)
        if i != -1:
            found[key] = sum(xs[i:i + len(name)]) / len(name)
    return found if 'description' in found and 'budget' in found else None


def _split_cells(line):
    """머리글이 없을 때: 단어 사이가 크게 벌어진 곳에서 칸을 나눈다."""
    hs = sorted(w[4] for w in line)
    gap = max(8.0, hs[len(hs) // 2] * 1.2)
    cells, cur, last_right = [], [], None
    for w in line:
        if cur and w[1] - last_right > gap:
            cells.append(cur)
            cur = []
        cur.append(w)
        last_right = w[1] + w[3]
    if cur:
        cells.append(cur)
    return [' '.join(str(w[0]) for w in c) for c in cells]


def rows_from_words(words) -> list:
    """Windows OCR 단어 [(글자, x, y, 폭, 높이)] → 줄 목록 (AI 없이 정리할 때)."""
    lines = _lines(words)
    cols, start = None, 0
    for i, ln in enumerate(lines):
        cols = _header_columns(ln)
        if cols:
            start = i + 1
            break
    rows = []
    if cols:
        body = lines[start:]
        # 칸 사이의 세로 빈 틈으로 덩어리를 나누고(여러 줄을 겹쳐 보면 칸 안의 띄어쓰기는 메워짐),
        # 덩어리마다 가운데가 가장 가까운 머리글의 열로 정한다. 머리글은 가운데 정렬, 내용은 왼쪽 정렬이라
        # 머리글 위치만으로 경계를 정하면 긴 세부항목 이름이 옆 칸으로 넘어간다.
        segs = []
        for a, b in sorted((w[1], w[1] + w[3]) for ln in body for w in ln):
            if segs and a <= segs[-1][1] + 3:
                segs[-1][1] = max(segs[-1][1], b)
            else:
                segs.append([a, b])

        def col_of(w):
            xc = w[1] + w[3] / 2
            seg = min(segs, key=lambda sg: 0 if sg[0] <= xc <= sg[1] else min(abs(xc - sg[0]), abs(xc - sg[1])))
            inside = [k for k in cols if seg[0] <= cols[k] <= seg[1]]
            if len(inside) >= 2:          # 붙어 있는 두 칸(예: 예산현액·예산잔액)이 한 덩어리 → 단어 위치로
                return min(cols, key=lambda k: abs(cols[k] - xc))
            sc = (seg[0] + seg[1]) / 2
            return min(cols, key=lambda k: abs(cols[k] - sc))
        for ln in body:
            cell = {}
            for w in ln:
                cell.setdefault(col_of(w), []).append(str(w[0]))
            cell.pop('select', None)
            rows.append({k: ' '.join(v) for k, v in cell.items()})
    else:
        for ln in lines:
            texts = _split_cells(ln)
            nums = []
            while texts and to_amount(texts[-1]) is not None and len(nums) < 2:
                nums.insert(0, texts.pop())
            if len(nums) < 2:
                continue
            texts = [t for t in texts if t.strip() not in ('□', '☐', 'ㅁ', '口')]
            row = {'budget': nums[0], 'remaining': nums[1]}
            if len(texts) >= 3:
                row.update(category=texts[0], cost_type=texts[1], description=' '.join(texts[2:]))
            elif len(texts) == 2:
                row.update(cost_type=texts[0], description=texts[1])
            elif texts:
                row.update(description=texts[0])
            rows.append(row)
    return normalize_rows(rows)


def words_to_text(words) -> str:
    """OCR 단어 → 줄마다 칸을 ' | '로 이은 글자 (AI에 이미지 대신 보낼 때)."""
    return '\n'.join(' | '.join(_split_cells(ln)) for ln in _lines(words))


# ── 실행 (뒤에서) ────────────────────────────────────────────────

try:
    from PyQt5.QtCore import QThread, pyqtSignal
except ImportError:          # 테스트 환경 등
    QThread = object

    def pyqtSignal(*_a):     # noqa: N802
        return None


class BudgetImageWorker(QThread):
    """캡처 이미지 → 세부항목별 묶음.

    done(categories: list, method: str)   method = 'ai-image' | 'ai-text' | 'ocr'
    failed(message: str)
    progress(message: str)
    """
    done = pyqtSignal(list, str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, png_b64: str, ocr_png: bytes, parent=None):
        super().__init__(parent)
        self._b64 = png_b64
        self._ocr_png = ocr_png
        self.cancelled = False

    def _ocr(self):
        self.progress.emit('PC에서 표의 글자를 읽는 중…')
        try:
            from capture import ocr_words
            return ocr_words(self._ocr_png)
        except ImportError:
            raise RuntimeError('이 PC에서는 Windows 글자 인식을 쓸 수 없습니다.\n'
                               'AI를 연결하면 이미지로 읽을 수 있어요 (공문 작성 창 → AI 설정).')

    def run(self):
        try:
            from ai_client import get_adapter
            from school_profile import privacy_enabled, mask_for_ai
            adapter = get_adapter()
            has_ai = adapter is not None and bool(adapter.get_api_key())
            rows, method, image_error = None, '', None

            if has_ai and not privacy_enabled():
                self.progress.emit('AI가 캡처한 표를 읽는 중…')
                try:
                    rows = parse_rows(adapter.complete(SYSTEM, build_prompt(), image_b64=self._b64, max_tokens=8000))
                    method = 'ai-image'
                except Exception as e:          # 이미지를 못 읽는 모델 등 → 글자로 다시
                    image_error = e
            if rows is None:
                words = self._ocr()
                if not words:
                    raise RuntimeError('표에서 글자를 읽지 못했습니다. 표가 잘 보이게 다시 캡처해 주세요.')
                if has_ai:
                    text = words_to_text(words)
                    if privacy_enabled():
                        text, _ = mask_for_ai(text)
                    self.progress.emit('AI가 표를 정리하는 중…')
                    rows = parse_rows(adapter.complete(SYSTEM, build_prompt(text), max_tokens=8000))
                    method = 'ai-text'
                else:
                    rows = rows_from_words(words)
                    method = 'ocr'
            cats = group_rows(rows)
            if not cats:
                hint = f'\n({image_error})' if image_error else ''
                raise RuntimeError('표에서 예산 줄을 찾지 못했습니다. 세부항목·산출내역·예산현액·예산잔액이 '
                                   '보이게 캡처했는지 확인해 주세요.' + hint)
            if not self.cancelled:
                self.done.emit(cats, method)
        except Exception as e:
            if not self.cancelled:
                try:
                    from ai_client import friendly_error
                    msg = friendly_error(e) if getattr(e, 'status_code', None) else str(e)
                except Exception:
                    msg = str(e)
                self.failed.emit(msg)
