import sqlite3
import os

DB_DIR  = os.path.join(os.environ.get('APPDATA', '.'), 'Tmemo')
DB_PATH = os.path.join(DB_DIR, 'tmemo.db')


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
                strikethrough INTEGER NOT NULL DEFAULT 0
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
        """)

        # tasks 테이블에 window_id 컬럼 없으면 추가 (기존 DB 마이그레이션)
        cols = [r[1] for r in conn.execute('PRAGMA table_info(tasks)').fetchall()]
        if 'window_id' not in cols:
            conn.execute('ALTER TABLE tasks ADD COLUMN window_id INTEGER NOT NULL DEFAULT 0')

        wcols = [r[1] for r in conn.execute('PRAGMA table_info(windows)').fetchall()]
        if 'color' not in wcols:
            conn.execute("ALTER TABLE windows ADD COLUMN color TEXT DEFAULT ''")

        if 'strikethrough' not in cols:
            conn.execute('ALTER TABLE tasks ADD COLUMN strikethrough INTEGER NOT NULL DEFAULT 0')

        hcols = [r[1] for r in conn.execute('PRAGMA table_info(task_history)').fetchall()]
        if 'strikethrough' not in hcols:
            conn.execute('ALTER TABLE task_history ADD COLUMN strikethrough INTEGER NOT NULL DEFAULT 0')

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

        # 창이 하나도 없으면 기본 창 생성
        count = conn.execute('SELECT COUNT(*) FROM windows').fetchone()[0]
        if count == 0:
            conn.execute('INSERT INTO windows (x, y, width, height) VALUES (100, 100, 320, 400)')


def get_all_windows():
    with _connect() as conn:
        rows = conn.execute('SELECT * FROM windows').fetchall()
        return [dict(r) for r in rows]


def create_window(x=130, y=130, width=320, height=400):
    with _connect() as conn:
        cur = conn.execute(
            'INSERT INTO windows (x, y, width, height) VALUES (?,?,?,?)',
            (x, y, width, height)
        )
        return cur.lastrowid


def update_window(window_id, x, y, width, height, collapsed, color=''):
    with _connect() as conn:
        conn.execute(
            'UPDATE windows SET x=?,y=?,width=?,height=?,collapsed=?,color=? WHERE id=?',
            (x, y, width, height, 1 if collapsed else 0, color, window_id)
        )


def delete_window(window_id):
    with _connect() as conn:
        conn.execute('DELETE FROM windows WHERE id=?', (window_id,))
        conn.execute('DELETE FROM tasks WHERE window_id=?', (window_id,))
        conn.execute('DELETE FROM task_history WHERE window_id=?', (window_id,))
        conn.execute('DELETE FROM documents WHERE window_id=?', (window_id,))


def get_tasks(window_id):
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM tasks WHERE window_id=? ORDER BY CASE WHEN deadline='' THEN 1 ELSE 0 END, deadline ASC",
            (window_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def add_task(window_id, name, deadline, strikethrough=0):
    with _connect() as conn:
        conn.execute(
            'INSERT INTO tasks (window_id, name, deadline, strikethrough) VALUES (?, ?, ?, ?)',
            (window_id, name, deadline, strikethrough)
        )


def delete_task(task_id):
    with _connect() as conn:
        conn.execute('DELETE FROM tasks WHERE id = ?', (task_id,))


def update_task(task_id, name, deadline, strikethrough=0):
    with _connect() as conn:
        conn.execute('UPDATE tasks SET name=?, deadline=?, strikethrough=? WHERE id=?', (name, deadline, strikethrough, task_id))


def add_task_history(window_id, name, deadline, strikethrough=0):
    with _connect() as conn:
        conn.execute(
            'INSERT INTO task_history (window_id, name, deadline, strikethrough) VALUES (?, ?, ?, ?)',
            (window_id, name, deadline, strikethrough)
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
