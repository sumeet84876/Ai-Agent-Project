"""
memory_store.py
----------------
Persistent chat history using SQLite, stored next to agent.py/app.py.
Survives app restarts. Only cleared when the user explicitly confirms
"Clear ALL history" in the sidebar - mirrors how ChatGPT/Claude keep
history around by default until the user deletes it.
"""

import os
import sqlite3
import time
from contextlib import contextmanager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "chat_memory.db")


@contextmanager
def _connect():
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with _connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                file_ref TEXT,
                created_at REAL NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                session_id TEXT PRIMARY KEY,
                title TEXT,
                created_at REAL NOT NULL
            )
        """)


def create_session(session_id: str, title: str = "New chat"):
    with _connect() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO sessions (session_id, title, created_at) VALUES (?, ?, ?)",
            (session_id, title, time.time()),
        )


def rename_session(session_id: str, title: str):
    with _connect() as conn:
        conn.execute("UPDATE sessions SET title = ? WHERE session_id = ?", (title, session_id))


def list_sessions():
    with _connect() as conn:
        rows = conn.execute(
            "SELECT session_id, title, created_at FROM sessions ORDER BY created_at DESC"
        ).fetchall()
    return [{"session_id": r[0], "title": r[1], "created_at": r[2]} for r in rows]


def add_message(session_id: str, role: str, content: str, file_ref: str = None):
    with _connect() as conn:
        conn.execute(
            "INSERT INTO messages (session_id, role, content, file_ref, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (session_id, role, content, file_ref, time.time()),
        )


def get_messages(session_id: str):
    with _connect() as conn:
        rows = conn.execute(
            "SELECT role, content, file_ref, created_at FROM messages "
            "WHERE session_id = ? ORDER BY id ASC",
            (session_id,),
        ).fetchall()
    return [{"role": r[0], "content": r[1], "file_ref": r[2], "created_at": r[3]} for r in rows]


def clear_all():
    """Explicit, user-triggered wipe. Never called automatically."""
    with _connect() as conn:
        conn.execute("DELETE FROM messages")
        conn.execute("DELETE FROM sessions")