"""
Database layer — SQLite with thread-safe connection pooling.
Handles: users, question sets, questions, answers, leaderboard.
"""

import sqlite3
import json
import threading
from typing import Optional

DB_PATH = "quiz_bot.db"
_local  = threading.local()


def _conn() -> sqlite3.Connection:
    if not hasattr(_local, "conn"):
        _local.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        _local.conn.row_factory = sqlite3.Row
        _local.conn.execute("PRAGMA journal_mode=WAL")   # better concurrency
        _local.conn.execute("PRAGMA synchronous=NORMAL")
    return _local.conn


def init_db():
    c = _conn()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY,
            name     TEXT,
            username TEXT,
            joined   TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS question_sets (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            name    TEXT NOT NULL,
            created TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS questions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            set_id      INTEGER REFERENCES question_sets(id) ON DELETE CASCADE,
            question    TEXT NOT NULL,
            options     TEXT NOT NULL,   -- JSON array
            correct     INTEGER NOT NULL,
            explanation TEXT DEFAULT '',
            timer       INTEGER DEFAULT 20,
            photo_id    TEXT
        );

        CREATE TABLE IF NOT EXISTS answers (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER,
            user_name  TEXT,
            poll_id    TEXT,
            chosen     INTEGER,
            correct    INTEGER,
            time_taken REAL,
            ts         TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS leaderboard (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id  INTEGER,
            user_id  INTEGER,
            name     TEXT,
            score    INTEGER DEFAULT 0,
            correct  INTEGER DEFAULT 0,
            wrong    INTEGER DEFAULT 0,
            ts       TEXT DEFAULT (datetime('now')),
            UNIQUE(chat_id, user_id)
        );
    """)
    c.commit()


# ── Users ────────────────────────────────────────────────────────────────────

def register_user(user_id: int, name: str, username: Optional[str]):
    c = _conn()
    c.execute(
        "INSERT OR IGNORE INTO users(id, name, username) VALUES (?,?,?)",
        (user_id, name, username)
    )
    c.commit()


# ── Sets ─────────────────────────────────────────────────────────────────────

def create_set(name: str) -> int:
    c = _conn()
    cur = c.execute("INSERT INTO question_sets(name) VALUES (?)", (name,))
    c.commit()
    return cur.lastrowid


def get_all_sets() -> list:
    c = _conn()
    rows = c.execute("""
        SELECT s.id, s.name, COUNT(q.id) as count
        FROM question_sets s
        LEFT JOIN questions q ON q.set_id = s.id
        GROUP BY s.id
        ORDER BY s.id DESC
    """).fetchall()
    return [dict(r) for r in rows]


def get_set(set_id: int) -> Optional[dict]:
    c = _conn()
    row = c.execute("SELECT * FROM question_sets WHERE id=?", (set_id,)).fetchone()
    return dict(row) if row else None


# ── Questions ────────────────────────────────────────────────────────────────

def add_question(set_id: int, question: str, options: list,
                 correct: int, explanation: str = "",
                 timer: int = 20, photo_id: Optional[str] = None):
    c = _conn()
    c.execute(
        "INSERT INTO questions(set_id,question,options,correct,explanation,timer,photo_id) VALUES (?,?,?,?,?,?,?)",
        (set_id, question, json.dumps(options, ensure_ascii=False), correct, explanation, timer, photo_id)
    )
    c.commit()


def get_questions(set_id: int) -> list:
    c = _conn()
    rows = c.execute(
        "SELECT * FROM questions WHERE set_id=? ORDER BY id", (set_id,)
    ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["options"] = json.loads(d["options"])
        result.append(d)
    return result


# ── Answers ──────────────────────────────────────────────────────────────────

def record_answer(user_id: int, user_name: str, poll_id: str,
                  chosen: int, correct: int, time_taken: float):
    c = _conn()
    c.execute(
        "INSERT INTO answers(user_id,user_name,poll_id,chosen,correct,time_taken) VALUES (?,?,?,?,?,?)",
        (user_id, user_name, poll_id, chosen, correct, time_taken)
    )
    c.commit()


def reset_session(chat_id: int):
    """Clear per-session data (not leaderboard)."""
    pass   # session is in chat_data; persistent leaderboard kept separately


# ── Leaderboard ──────────────────────────────────────────────────────────────

def save_leaderboard(chat_id: int, sorted_scores: list):
    c = _conn()
    for uid, s in sorted_scores:
        c.execute("""
            INSERT INTO leaderboard(chat_id,user_id,name,score,correct,wrong)
            VALUES (?,?,?,?,?,?)
            ON CONFLICT(chat_id,user_id) DO UPDATE SET
                score  = score  + excluded.score,
                correct= correct+ excluded.correct,
                wrong  = wrong  + excluded.wrong,
                ts     = datetime('now')
        """, (chat_id, uid, s["name"], s["score"], s["correct"], s["wrong"]))
    c.commit()


def get_leaderboard(chat_id: int) -> list:
    c = _conn()
    rows = c.execute("""
        SELECT name, score, correct, wrong
        FROM leaderboard WHERE chat_id=?
        ORDER BY score DESC, wrong ASC LIMIT 50
    """, (chat_id,)).fetchall()
    return [dict(r) for r in rows]


def reset_leaderboard(chat_id: int):
    c = _conn()
    c.execute("DELETE FROM leaderboard WHERE chat_id=?", (chat_id,))
    c.commit()


# Auto-init on import
init_db()


def cleanup_old_answers(days: int = 30):
    """BUG #8 FIX — 30 din purane answers delete karo DB size control ke liye."""
    c = _conn()
    c.execute("DELETE FROM answers WHERE ts < datetime('now', ? || ' days')", (f'-{days}',))
    c.commit()
    logger.info(f'Old answers cleaned up (>{days} days)')
