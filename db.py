import sqlite3
import os

# APPDATA가 없는 환경(일부 프로그램이 줄인 환경 변수로 실행할 때)에서도 늘 같은 DB를 쓰도록
_APPDATA = os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')
DB_DIR  = os.path.join(_APPDATA, 'SSNnote')
DB_PATH = os.path.join(DB_DIR, 'ssnnote.db')


def _connect():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with _connect() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS windows (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                x         INTEGER DEFAULT 100,
                y         INTEGER DEFAULT 100,
                width     INTEGER DEFAULT 320,
                height    INTEGER DEFAULT 400,
                collapsed INTEGER DEFAULT 0,
                color     TEXT    DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS tasks (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                window_id     INTEGER NOT NULL,
                name          TEXT NOT NULL,
                deadline      TEXT NOT NULL DEFAULT '',
                strikethrough INTEGER NOT NULL DEFAULT 0,
                priority      INTEGER NOT NULL DEFAULT 0,
                recurrence    TEXT NOT NULL DEFAULT '',
                notes         TEXT NOT NULL DEFAULT '',
                related_no    TEXT NOT NULL DEFAULT '',
                attachments   TEXT NOT NULL DEFAULT '[]'
            );

            CREATE TABLE IF NOT EXISTS task_history (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                window_id     INTEGER NOT NULL,
                name          TEXT NOT NULL,
                deadline      TEXT NOT NULL DEFAULT '',
                strikethrough INTEGER NOT NULL DEFAULT 0,
                cleared_at    TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS documents (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                window_id  INTEGER NOT NULL,
                title      TEXT NOT NULL DEFAULT '',
                doc_number TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS official_documents (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                title      TEXT NOT NULL DEFAULT '',
                doc_number TEXT NOT NULL DEFAULT '',
                content    TEXT NOT NULL DEFAULT '',
                doc_type   TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS doc_embeddings (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                doc_id      INTEGER NOT NULL,
                title       TEXT NOT NULL DEFAULT '',
                content     TEXT NOT NULL DEFAULT '',
                doc_type    TEXT NOT NULL DEFAULT '',
                embedding   BLOB,
                created_at  TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS budget_categories (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT NOT NULL UNIQUE,
                sort_order INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS budget_items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                category_id INTEGER NOT NULL,
                cost_type   TEXT NOT NULL DEFAULT '',
                description TEXT NOT NULL DEFAULT '',
                budget      INTEGER NOT NULL DEFAULT 0,
                spent       INTEGER NOT NULL DEFAULT 0,
                completed   INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS budget_expenses (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id    INTEGER NOT NULL,
                amount     INTEGER NOT NULL DEFAULT 0,
                purpose    TEXT NOT NULL DEFAULT '',
                spent_on   TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
            );
        """)

        # 예산 세부항목: 전체 합계에서 빼기 · 그룹 (v3.01)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS budget_groups (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT NOT NULL UNIQUE,
                sort_order INTEGER NOT NULL DEFAULT 0,
                collapsed  INTEGER NOT NULL DEFAULT 0
            )""")
        bcols = [r[1] for r in conn.execute('PRAGMA table_info(budget_categories)').fetchall()]
        if 'excluded' not in bcols:
            conn.execute('ALTER TABLE budget_categories ADD COLUMN excluded INTEGER NOT NULL DEFAULT 0')
        if 'group_id' not in bcols:
            conn.execute('ALTER TABLE budget_categories ADD COLUMN group_id INTEGER DEFAULT NULL')

        # tasks 테이블에 window_id 컬럼 없으면 추가 (기존 DB 마이그레이션)
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        if 'window_id' not in cols:
            conn.execute('ALTER TABLE tasks ADD COLUMN window_id INTEGER NOT NULL DEFAULT 0')

        wcols = [r[1] for r in conn.execute('PRAGMA table_info(windows)').fetchall()]
        if 'color' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN color TEXT DEFAULT ''")
        if 'scale' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN scale REAL DEFAULT 1.0")
        if 'memo_mode' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN memo_mode INTEGER DEFAULT 0")
        if 'memo_text' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN memo_text TEXT DEFAULT ''")
        if 'merge_group_id' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN merge_group_id INTEGER DEFAULT NULL")
        if 'row_gap' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN row_gap INTEGER DEFAULT 0")
        if 'font' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN font TEXT DEFAULT ''")   # 창마다 고른 글꼴 ('' = Pretendard)

        if 'strikethrough' not in cols:
            conn.execute('ALTER TABLE tasks ADD COLUMN strikethrough INTEGER NOT NULL DEFAULT 0')

        hcols = [r[1] for r in conn.execute('PRAGMA table_info(task_history)').fetchall()]
        if 'strikethrough' not in hcols:
            conn.execute('ALTER TABLE task_history ADD COLUMN strikethrough INTEGER NOT NULL DEFAULT 0')

        # priority, recurrence 컬럼 마이그레이션
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        if 'priority' not in cols:
            conn.execute('ALTER TABLE tasks ADD COLUMN priority INTEGER NOT NULL DEFAULT 0')
        hcols = [r[1] for r in conn.execute('PRAGMA table_info(task_history)').fetchall()]
        if 'priority' not in hcols:
            conn.execute('ALTER TABLE task_history ADD COLUMN priority INTEGER NOT NULL DEFAULT 0')
        if 'recurrence' not in cols:
            conn.execute("ALTER TABLE tasks ADD COLUMN recurrence TEXT NOT NULL DEFAULT ''")
        if 'recurrence' not in hcols:
            conn.execute("ALTER TABLE task_history ADD COLUMN recurrence TEXT NOT NULL DEFAULT ''")

        # notes 컬럼 마이그레이션
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        hcols = [r[1] for r in conn.execute('PRAGMA table_info(task_history)').fetchall()]
        if 'notes' not in cols:
            conn.execute("ALTER TABLE tasks ADD COLUMN notes TEXT NOT NULL DEFAULT ''")
        if 'notes' not in hcols:
            conn.execute("ALTER TABLE task_history ADD COLUMN notes TEXT NOT NULL DEFAULT ''")

        # related_no 컬럼 마이그레이션 (공문 관련번호)
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        hcols = [r[1] for r in conn.execute('PRAGMA table_info(task_history)').fetchall()]
        if 'related_no' not in cols:
            conn.execute("ALTER TABLE tasks ADD COLUMN related_no TEXT NOT NULL DEFAULT ''")
        if 'related_no' not in hcols:
            conn.execute("ALTER TABLE task_history ADD COLUMN related_no TEXT NOT NULL DEFAULT ''")

        # attachments 컬럼 마이그레이션 (업무 붙임파일 — 파일 경로 JSON 배열)
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        if 'attachments' not in cols:
            conn.execute("ALTER TABLE tasks ADD COLUMN attachments TEXT NOT NULL DEFAULT '[]'")

        # app_settings 테이블 (학교 프로필, 개인정보 보호 모드 등 앱 설정 — key/value)
        conn.execute(
            "CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL DEFAULT '')"
        )

        # 기존 window_state 마이그레이션
        tables = [r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()]
        if 'window_state' in tables:
            count = conn.execute('SELECT COUNT(*) FROM windows').fetchone()[0]
            if count == 0:
                row = conn.execute('SELECT * FROM window_state WHERE id=1').fetchone()
                if row:
                    conn.execute(
                        'INSERT INTO windows (x, y, width, height, collapsed) VALUES (?,?,?,?,?)',
                        (row['x'], row['y'], row['width'], row['height'], row['collapsed'])
                    )

        # 배경색 새 팔레트(v3.06) — 예전 색을 가장 가까운 새 색으로 (창 · AI 대화 창)
        from ui.theme import OLD_COLOR_MAP
        for old, new in OLD_COLOR_MAP.items():
            conn.execute('UPDATE windows SET color=? WHERE UPPER(color)=?', (new, old))
            conn.execute("UPDATE app_settings SET value=? WHERE key='ai_chat_color' AND UPPER(value)=?", (new, old))

        # 창이 하나도 없으면 기본 창 생성
        count = conn.execute('SELECT COUNT(*) FROM windows').fetchone()[0]
        if count == 0:
            conn.execute('INSERT INTO windows (x, y, width, height) VALUES (100, 100, 320, 400)')


def get_all_windows():
    with _connect() as conn:
        rows = conn.execute('SELECT * FROM windows').fetchall()
        return [dict(r) for r in rows]


def get_first_window_id():
    with _connect() as conn:
        row = conn.execute('SELECT MIN(id) AS mid FROM windows').fetchone()
        return row['mid'] if row else None


def create_window(x=130, y=130, width=320, height=400):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO windows (x, y, width, height) VALUES (?,?,?,?)',
            (x, y, width, height)
        )
        return cur.lastrowid


def update_window(window_id, x, y, width, height, collapsed, color='', scale=1.0, row_gap=0):
    with _connect() as conn:
        conn.execute(
            'UPDATE windows SET x=?,y=?,width=?,height=?,collapsed=?,color=?,scale=?,row_gap=? WHERE id=?',
            (x, y, width, height, 1 if collapsed else 0, color, scale, row_gap, window_id)
        )


def get_window_font(window_id) -> str:
    with _connect() as conn:
        row = conn.execute('SELECT font FROM windows WHERE id=?', (window_id,)).fetchone()
        return (row['font'] or '') if row else ''


def set_window_font(window_id, font):
    with _connect() as conn:
        conn.execute('UPDATE windows SET font=? WHERE id=?', (font or '', window_id))


def set_window_merge_group(window_id, group_id):
    with _connect() as conn:
        conn.execute('UPDATE windows SET merge_group_id=? WHERE id=?', (group_id, window_id))


def clear_window_merge_group(window_id):
    with _connect() as conn:
        conn.execute('UPDATE windows SET merge_group_id=NULL WHERE id=?', (window_id,))


def get_next_merge_group_id():
    with _connect() as conn:
        row = conn.execute('SELECT MAX(merge_group_id) FROM windows').fetchone()
        return (row[0] or 0) + 1


def delete_window(window_id):
    with _connect() as conn:
        conn.execute('DELETE FROM windows WHERE id=?', (window_id,))
        conn.execute('DELETE FROM tasks WHERE window_id=?', (window_id,))
        conn.execute('DELETE FROM task_history WHERE window_id=?', (window_id,))
        conn.execute('DELETE FROM documents WHERE window_id=?', (window_id,))


def get_tasks(window_id):
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM tasks WHERE window_id=? ORDER BY priority DESC, CASE WHEN deadline='' THEN 1 ELSE 0 END, deadline ASC",
            (window_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def get_all_tasks():
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM tasks WHERE deadline != '' ORDER BY deadline ASC"
        ).fetchall()
        return [dict(r) for r in rows]


def get_task(task_id):
    """업무 한 건 (없으면 None)."""
    with _connect() as conn:
        row = conn.execute('SELECT * FROM tasks WHERE id=?', (task_id,)).fetchone()
        return dict(row) if row else None


def get_open_tasks():
    """완료(취소선) 표시가 없는 모든 업무 + 창 색상 — 오늘 대시보드용."""
    with _connect() as conn:
        rows = conn.execute(
            """SELECT t.*, w.color FROM tasks t JOIN windows w ON t.window_id = w.id
               WHERE t.strikethrough = 0"""
        ).fetchall()
        return [dict(r) for r in rows]


def get_tasks_by_date(date_str: str):
    with _connect() as conn:
        rows = conn.execute(
            """SELECT t.id, t.window_id, t.name, t.deadline, t.strikethrough,
                      t.priority, t.recurrence, t.notes, t.related_no, w.color
               FROM tasks t JOIN windows w ON t.window_id = w.id
               WHERE t.deadline LIKE ? AND t.deadline != ''
               ORDER BY t.priority DESC, t.deadline ASC""",
            (date_str + '%',)
        ).fetchall()
        return [dict(r) for r in rows]


def get_all_task_history():
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM task_history ORDER BY deadline ASC'
        ).fetchall()
        return [dict(r) for r in rows]


def get_all_tasks_for_calendar():
    with _connect() as conn:
        rows = conn.execute(
            "SELECT SUBSTR(deadline, 1, 10) as d, COUNT(*) as cnt FROM tasks WHERE deadline != '' GROUP BY d"
        ).fetchall()
        return {row['d']: row['cnt'] for row in rows}


def add_task(window_id, name, deadline, strikethrough=0, priority=0, recurrence='', notes='', related_no=''):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO tasks (window_id, name, deadline, strikethrough, priority, recurrence, notes, related_no) VALUES (?,?,?,?,?,?,?,?)',
            (window_id, name, deadline, strikethrough, priority, recurrence, notes, related_no)
        )
        return cur.lastrowid


def delete_task(task_id):
    with _connect() as conn:
        conn.execute('DELETE FROM tasks WHERE id = ?', (task_id,))


def update_task(task_id, name, deadline, strikethrough=0, priority=0, recurrence=''):
    with _connect() as conn:
        conn.execute(
            'UPDATE tasks SET name=?, deadline=?, strikethrough=?, priority=?, recurrence=? WHERE id=?',
            (name, deadline, strikethrough, priority, recurrence, task_id)
        )


def add_task_history(window_id, name, deadline, strikethrough=0, priority=0, recurrence='', notes='', related_no=''):
    """지난 기록에 남기기. 남긴 줄의 id를 돌려준다 (AI가 지운 업무를 되돌릴 때 씀)."""
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO task_history (window_id, name, deadline, strikethrough, priority, recurrence, notes, related_no) VALUES (?,?,?,?,?,?,?,?)',
            (window_id, name, deadline, strikethrough, priority, recurrence, notes, related_no)
        )
        return cur.lastrowid


def set_task_window(task_id, window_id):
    """업무를 다른 업무창으로 옮기기."""
    with _connect() as conn:
        conn.execute('UPDATE tasks SET window_id=? WHERE id=?', (window_id, task_id))


def restore_task(task: dict):
    """지운 업무를 원래 값으로 다시 넣기 (id는 새로). 새 id를 돌려준다."""
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO tasks (window_id, name, deadline, strikethrough, priority, recurrence, notes, related_no, '
            'attachments) VALUES (?,?,?,?,?,?,?,?,?)',
            (task['window_id'], task['name'], task.get('deadline', ''), task.get('strikethrough', 0),
             task.get('priority', 0), task.get('recurrence', ''), task.get('notes', ''),
             task.get('related_no', ''), task.get('attachments') or '[]'))
        return cur.lastrowid


def set_task_priority(task_id, priority):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET priority=? WHERE id=?', (priority, task_id))


def set_task_recurrence(task_id, recurrence):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET recurrence=? WHERE id=?', (recurrence, task_id))


def get_task_notes(task_id: int) -> str:
    with _connect() as conn:
        row = conn.execute('SELECT notes FROM tasks WHERE id=?', (task_id,)).fetchone()
        return row['notes'] if row else ''


def set_task_notes(task_id: int, notes: str):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET notes=? WHERE id=?', (notes, task_id))


def get_task_related_no(task_id: int) -> str:
    with _connect() as conn:
        row = conn.execute('SELECT related_no FROM tasks WHERE id=?', (task_id,)).fetchone()
        return row['related_no'] if row else ''


def set_task_related_no(task_id: int, related_no: str):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET related_no=? WHERE id=?', (related_no, task_id))


def get_task_attachments(task_id: int) -> str:
    with _connect() as conn:
        row = conn.execute('SELECT attachments FROM tasks WHERE id=?', (task_id,)).fetchone()
        return row['attachments'] if row else '[]'


def set_task_attachments(task_id: int, attachments_json: str):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET attachments=? WHERE id=?', (attachments_json, task_id))


def search_tasks_all(query):
    with _connect() as conn:
        rows = conn.execute(
            """SELECT t.id, t.window_id, t.name, t.deadline, t.strikethrough,
                      t.priority, t.recurrence, w.color
               FROM tasks t JOIN windows w ON t.window_id = w.id
               WHERE t.name LIKE ? COLLATE NOCASE
               ORDER BY t.priority DESC,
                        CASE WHEN t.deadline='' THEN 1 ELSE 0 END,
                        t.deadline ASC""",
            (f'%{query}%',)
        ).fetchall()
        return [dict(r) for r in rows]


def get_window_memo(window_id):
    with _connect() as conn:
        row = conn.execute(
            'SELECT memo_mode, memo_text FROM windows WHERE id=?', (window_id,)
        ).fetchone()
        if row:
            return row['memo_mode'] or 0, row['memo_text'] or ''
        return 0, ''


def set_window_memo(window_id, memo_mode, memo_text):
    with _connect() as conn:
        conn.execute(
            'UPDATE windows SET memo_mode=?, memo_text=? WHERE id=?',
            (memo_mode, memo_text, window_id)
        )


def get_task_history(window_id):
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM task_history WHERE window_id=? ORDER BY cleared_at DESC',
            (window_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def delete_task_history(history_id):
    with _connect() as conn:
        conn.execute('DELETE FROM task_history WHERE id=?', (history_id,))


def get_documents(window_id):
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM documents WHERE window_id=? ORDER BY created_at ASC',
            (window_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def add_document(window_id, title='', doc_number=''):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO documents (window_id, title, doc_number) VALUES (?, ?, ?)',
            (window_id, title, doc_number)
        )
        return cur.lastrowid


def update_document(doc_id, title, doc_number):
    with _connect() as conn:
        conn.execute(
            'UPDATE documents SET title=?, doc_number=? WHERE id=?',
            (title, doc_number, doc_id)
        )


def delete_document(doc_id):
    with _connect() as conn:
        conn.execute('DELETE FROM documents WHERE id=?', (doc_id,))


# ── official_documents CRUD ──────────────────────────────────

def save_official_document(title, doc_number, content, doc_type):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO official_documents (title, doc_number, content, doc_type) VALUES (?, ?, ?, ?)',
            (title, doc_number, content, doc_type)
        )
        return cur.lastrowid


def get_official_documents():
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM official_documents ORDER BY created_at DESC'
        ).fetchall()
        return [dict(r) for r in rows]


def update_official_document(doc_id, title, doc_number, content):
    """저장한 공문 고치기 (지난 기록 → 작성 공문 → 수정). 종류(doc_type)는 그대로 둔다."""
    with _connect() as conn:
        conn.execute('UPDATE official_documents SET title=?, doc_number=?, content=? WHERE id=?',
                     (title, doc_number, content, doc_id))


def delete_official_document(doc_id):
    with _connect() as conn:
        conn.execute('DELETE FROM official_documents WHERE id=?', (doc_id,))
        conn.execute('DELETE FROM doc_embeddings WHERE doc_id=?', (doc_id,))


# ── doc_embeddings CRUD ─────────────────────────────────────────

def save_embedding(doc_id: int, title: str, content: str, doc_type: str, embedding_bytes: bytes):
    with _connect() as conn:
        conn.execute(
            'INSERT INTO doc_embeddings (doc_id, title, content, doc_type, embedding) VALUES (?,?,?,?,?)',
            (doc_id, title, content, doc_type, embedding_bytes)
        )


def get_all_embeddings():
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM doc_embeddings ORDER BY created_at DESC'
        ).fetchall()
        return [dict(r) for r in rows]


def delete_embedding(emb_id: int):
    with _connect() as conn:
        conn.execute('DELETE FROM doc_embeddings WHERE id=?', (emb_id,))


# ── 예산 관리 CRUD ──────────────────────────────────────────────

def get_budget_categories():
    with _connect() as conn:
        rows = conn.execute(
            'SELECT * FROM budget_categories ORDER BY sort_order ASC, id ASC'
        ).fetchall()
        return [dict(r) for r in rows]


def add_budget_category(name):
    with _connect() as conn:
        try:
            row = conn.execute('SELECT MAX(sort_order) FROM budget_categories').fetchone()
            next_order = (row[0] or 0) + 1
            cur = conn.execute(
                'INSERT INTO budget_categories (name, sort_order) VALUES (?, ?)',
                (name, next_order)
            )
            return cur.lastrowid
        except sqlite3.IntegrityError:
            return None


def rename_budget_category(cat_id, new_name):
    with _connect() as conn:
        try:
            conn.execute('UPDATE budget_categories SET name=? WHERE id=?', (new_name, cat_id))
            return True
        except sqlite3.IntegrityError:
            return False


def delete_budget_category(cat_id):
    with _connect() as conn:
        conn.execute('DELETE FROM budget_expenses WHERE item_id IN '
                     '(SELECT id FROM budget_items WHERE category_id=?)', (cat_id,))
        conn.execute('DELETE FROM budget_items WHERE category_id=?', (cat_id,))
        conn.execute('DELETE FROM budget_categories WHERE id=?', (cat_id,))


def set_budget_category_excluded(cat_id, excluded):
    """세부항목을 전체 합계에서 뺄지(1) 넣을지(0)."""
    with _connect() as conn:
        conn.execute('UPDATE budget_categories SET excluded=? WHERE id=?', (1 if excluded else 0, cat_id))


def set_budget_category_group(cat_id, group_id):
    """세부항목을 그룹에 넣기 (None이면 그룹에서 빼기). 세부항목 하나는 그룹 하나에만."""
    with _connect() as conn:
        conn.execute('UPDATE budget_categories SET group_id=? WHERE id=?', (group_id, cat_id))


def get_budget_groups():
    with _connect() as conn:
        rows = conn.execute('SELECT * FROM budget_groups ORDER BY sort_order ASC, id ASC').fetchall()
        return [dict(r) for r in rows]


def add_budget_group(name):
    """새 그룹. 같은 이름이 있으면 None."""
    with _connect() as conn:
        try:
            row = conn.execute('SELECT MAX(sort_order) FROM budget_groups').fetchone()
            cur = conn.execute('INSERT INTO budget_groups (name, sort_order) VALUES (?, ?)',
                               (name, (row[0] or 0) + 1))
            return cur.lastrowid
        except sqlite3.IntegrityError:
            return None


def rename_budget_group(group_id, new_name):
    with _connect() as conn:
        try:
            conn.execute('UPDATE budget_groups SET name=? WHERE id=?', (new_name, group_id))
            return True
        except sqlite3.IntegrityError:
            return False


def delete_budget_group(group_id):
    """그룹만 지운다. 안의 세부항목은 그대로 두고 그룹에서만 뺀다."""
    with _connect() as conn:
        conn.execute('UPDATE budget_categories SET group_id=NULL WHERE group_id=?', (group_id,))
        conn.execute('DELETE FROM budget_groups WHERE id=?', (group_id,))


def set_budget_group_collapsed(group_id, collapsed):
    with _connect() as conn:
        conn.execute('UPDATE budget_groups SET collapsed=? WHERE id=?', (1 if collapsed else 0, group_id))


def get_budget_items_in_total():
    """전체 합계에 들어가는 예산 항목만 (합계에서 뺀 세부항목의 항목은 제외)."""
    with _connect() as conn:
        rows = conn.execute(
            'SELECT i.* FROM budget_items i JOIN budget_categories c ON c.id = i.category_id '
            'WHERE c.excluded = 0 ORDER BY i.id ASC').fetchall()
        return [dict(r) for r in rows]


def get_budget_items(category_id=None):
    with _connect() as conn:
        if category_id is None:
            rows = conn.execute(
                'SELECT * FROM budget_items ORDER BY id ASC'
            ).fetchall()
        else:
            rows = conn.execute(
                'SELECT * FROM budget_items WHERE category_id=? ORDER BY id ASC',
                (category_id,)
            ).fetchall()
        return [dict(r) for r in rows]


def add_budget_item(category_id, cost_type, description, budget, spent):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO budget_items (category_id, cost_type, description, budget, spent) VALUES (?,?,?,?,?)',
            (category_id, cost_type, description, budget, spent)
        )
        return cur.lastrowid


def update_budget_item(item_id, category_id, cost_type, description, budget, spent):
    with _connect() as conn:
        conn.execute(
            'UPDATE budget_items SET category_id=?, cost_type=?, description=?, budget=?, spent=? WHERE id=?',
            (category_id, cost_type, description, budget, spent, item_id)
        )


def delete_budget_item(item_id):
    with _connect() as conn:
        conn.execute('DELETE FROM budget_expenses WHERE item_id=?', (item_id,))
        conn.execute('DELETE FROM budget_items WHERE id=?', (item_id,))


def set_budget_item_spent(item_id, spent):
    with _connect() as conn:
        conn.execute('UPDATE budget_items SET spent=? WHERE id=?', (spent, item_id))


def toggle_budget_item_completed(item_id):
    with _connect() as conn:
        row = conn.execute('SELECT completed FROM budget_items WHERE id=?', (item_id,)).fetchone()
        if not row:
            return 0
        new_val = 0 if row['completed'] else 1
        conn.execute('UPDATE budget_items SET completed=? WHERE id=?', (new_val, item_id))
        return new_val


def clear_budget_data(keep_groups=False):
    """예산 전부 지우기. keep_groups면 그룹(이름·순서)은 남긴다 (전체 교체로 다시 가져올 때)."""
    with _connect() as conn:
        conn.execute('DELETE FROM budget_expenses')
        conn.execute('DELETE FROM budget_items')
        conn.execute('DELETE FROM budget_categories')
        if not keep_groups:
            conn.execute('DELETE FROM budget_groups')


# ── 사용 내역 (명목별 지출 기록) ─────────────────────────────────

def get_budget_expenses():
    """모든 사용 내역 (최근 날짜 먼저)."""
    with _connect() as conn:
        rows = conn.execute('SELECT * FROM budget_expenses ORDER BY spent_on DESC, id DESC').fetchall()
        return [dict(r) for r in rows]


def _insert_expense(conn, item_id, amount, purpose, spent_on):
    cur = conn.execute(
        'INSERT INTO budget_expenses (item_id, amount, purpose, spent_on) VALUES (?,?,?,?)',
        (item_id, amount, purpose.strip(), spent_on))
    return cur.lastrowid


def restore_budget_item(item: dict, expenses=()):
    """지운 예산 항목을 원래 값(완료·사용 내역 포함)으로 다시 넣기 (id는 새로). 새 id를 돌려준다."""
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO budget_items (category_id, cost_type, description, budget, spent, completed) '
            'VALUES (?,?,?,?,?,?)',
            (item['category_id'], item.get('cost_type', ''), item.get('description', ''),
             int(item.get('budget') or 0), int(item.get('spent') or 0), int(item.get('completed') or 0)))
        new_id = cur.lastrowid
        for ex in expenses:
            _insert_expense(conn, new_id, ex['amount'], ex.get('purpose', ''), ex.get('spent_on', ''))
        return new_id


def set_budget_item_completed(item_id, completed):
    with _connect() as conn:
        conn.execute('UPDATE budget_items SET completed=? WHERE id=?', (1 if completed else 0, item_id))


def add_budget_expense(item_id, amount, purpose, spent_on):
    """사용 내역을 적고 집행액을 그만큼 더한다 (한 번에)."""
    with _connect() as conn:
        eid = _insert_expense(conn, item_id, amount, purpose, spent_on)
        conn.execute('UPDATE budget_items SET spent = spent + ? WHERE id=?', (amount, item_id))
        return eid


def save_budget_item_edit(item_id, category_id, cost_type, description, budget, spent, expense=None):
    """항목 수정 + (있으면) 이번에 쓴 금액의 사용 내역 기록. 집행액은 spent로 정한다 (이미 더해진 값)."""
    with _connect() as conn:
        conn.execute(
            'UPDATE budget_items SET category_id=?, cost_type=?, description=?, budget=?, spent=? WHERE id=?',
            (category_id, cost_type, description, budget, spent, item_id))
        if expense:
            _insert_expense(conn, item_id, expense['amount'], expense['purpose'], expense['spent_on'])


def delete_budget_expense(expense_id, adjust_spent=True):
    """사용 내역 한 건 삭제. adjust_spent면 집행액도 그만큼 줄인다 (되돌리기 때는 집행액을 따로 맞춘다)."""
    with _connect() as conn:
        row = conn.execute('SELECT item_id, amount FROM budget_expenses WHERE id=?', (expense_id,)).fetchone()
        if not row:
            return
        conn.execute('DELETE FROM budget_expenses WHERE id=?', (expense_id,))
        if adjust_spent:
            conn.execute('UPDATE budget_items SET spent = spent - ? WHERE id=?', (row['amount'], row['item_id']))


# ── 앱 설정 (key/value) ─────────────────────────────────────────

def get_setting(key: str, default: str = '') -> str:
    with _connect() as conn:
        row = conn.execute('SELECT value FROM app_settings WHERE key=?', (key,)).fetchone()
        return row['value'] if row else default


def set_setting(key: str, value: str):
    with _connect() as conn:
        conn.execute(
            'INSERT INTO app_settings (key, value) VALUES (?, ?) '
            'ON CONFLICT(key) DO UPDATE SET value=excluded.value',
            (key, value)
        )
