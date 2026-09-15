"""
Database abstraction layer for OneTrack.
Works with both SQLite (local development) and MySQL (Waifly production).
"""

import os
import json
from datetime import datetime, timedelta

# Detect environment
USE_MYSQL = bool(os.environ.get("DB_HOST"))

if USE_MYSQL:
    import mysql.connector
    from mysql.connector import Error as DBError
else:
    import sqlite3
    DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "onetrack.db")


class CompatibleRow:
    """Row that supports both index and dictionary access."""
    def __init__(self, data, columns):
        self._data = data
        self._columns = columns

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._data[key]
        elif isinstance(key, str):
            idx = self._columns.index(key)
            return self._data[idx]
        raise KeyError(key)

    def __contains__(self, key):
        if isinstance(key, str):
            return key in self._columns
        return False


def get_db():
    """Get database connection."""
    if USE_MYSQL:
        conn = mysql.connector.connect(
            host=os.environ.get("DB_HOST"),
            user=os.environ.get("DB_USER"),
            password=os.environ.get("DB_PASSWORD"),
            database=os.environ.get("DB_NAME"),
            autocommit=False
        )
        return conn
    else:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        return conn


def get_cursor(conn):
    """Get cursor with dictionary-like access."""
    if USE_MYSQL:
        return conn.cursor(dictionary=True)
    else:
        return conn.cursor()


def execute_query(conn, sql, params=None):
    """Execute a query and return results."""
    c = get_cursor(conn)
    if params:
        c.execute(sql, params)
    else:
        c.execute(sql)
    return c


def fetch_one(conn, sql, params=None):
    """Fetch single row as dictionary."""
    c = execute_query(conn, sql, params)
    row = c.fetchone()
    if row is None:
        return None
    if USE_MYSQL:
        return row
    else:
        return dict(row)


def fetch_all(conn, sql, params=None):
    """Fetch all rows as list of dictionaries."""
    c = execute_query(conn, sql, params)
    rows = c.fetchall()
    if USE_MYSQL:
        return rows
    else:
        return [dict(row) for row in rows]


def fetch_all_compatible(conn, sql, params=None):
    """Fetch all rows as CompatibleRow objects (supports both index and dict access)."""
    if USE_MYSQL:
        c = conn.cursor()
        c.execute(sql, params or ())
        rows = c.fetchall()
        columns = [desc[0] for desc in c.description]
        return [CompatibleRow(row, columns) for row in rows]
    else:
        c = conn.cursor()
        c.execute(sql, params or ())
        rows = c.fetchall()
        columns = [desc[0] for desc in c.description]
        return [CompatibleRow(row, columns) for row in rows]


def fetch_one_compatible(conn, sql, params=None):
    """Fetch single row as CompatibleRow object."""
    if USE_MYSQL:
        c = conn.cursor()
        c.execute(sql, params or ())
        row = c.fetchone()
        if row is None:
            return None
        columns = [desc[0] for desc in c.description]
        return CompatibleRow(row, columns)
    else:
        c = conn.cursor()
        c.execute(sql, params or ())
        row = c.fetchone()
        if row is None:
            return None
        columns = [desc[0] for desc in c.description]
        return CompatibleRow(row, columns)


def init_db():
    """Create tables if they don't exist."""
    conn = get_db()
    c = get_cursor(conn)

    if USE_MYSQL:
        # MySQL table definitions
        c.execute("""
            CREATE TABLE IF NOT EXISTS responses (
                id INT AUTO_INCREMENT PRIMARY KEY,
                timestamp VARCHAR(255),
                date VARCHAR(20),
                q1 TEXT,
                q2 TEXT,
                q3 TEXT,
                q4 TEXT,
                q5 TEXT,
                q6 TEXT,
                q7 TEXT
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS team_members (
                id INT AUTO_INCREMENT PRIMARY KEY,
                staff_id VARCHAR(50) UNIQUE,
                name VARCHAR(255),
                gender VARCHAR(10),
                picture VARCHAR(255),
                targets TEXT,
                tasks TEXT,
                created_at DATETIME
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(255) UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role VARCHAR(50) NOT NULL DEFAULT 'user',
                display_name VARCHAR(255),
                created_at DATETIME,
                email VARCHAR(255) DEFAULT '',
                reset_code VARCHAR(20) DEFAULT '',
                reset_expiry VARCHAR(50) DEFAULT ''
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                task_key VARCHAR(50) UNIQUE NOT NULL,
                label VARCHAR(100) NOT NULL,
                unit VARCHAR(50) NOT NULL,
                sort_order INT DEFAULT 0,
                active TINYINT(1) DEFAULT 1,
                created_at DATETIME
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """)
    else:
        # SQLite table definitions
        c.execute("""
            CREATE TABLE IF NOT EXISTS responses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                date TEXT,
                q1 TEXT, q2 TEXT, q3 TEXT, q4 TEXT,
                q5 TEXT, q6 TEXT, q7 TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS team_members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                staff_id TEXT UNIQUE,
                name TEXT, gender TEXT, picture TEXT,
                targets TEXT, tasks TEXT, created_at TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                display_name TEXT, created_at TEXT,
                email TEXT DEFAULT '',
                reset_code TEXT DEFAULT '',
                reset_expiry TEXT DEFAULT ''
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_key TEXT UNIQUE NOT NULL,
                label TEXT NOT NULL, unit TEXT NOT NULL,
                sort_order INTEGER DEFAULT 0,
                active INTEGER DEFAULT 1,
                created_at TEXT
            )
        """)

    conn.commit()
    conn.close()


def delete_old_responses(years=2):
    """Delete responses older than specified years."""
    conn = get_db()
    c = get_cursor(conn)

    cutoff = (datetime.now() - timedelta(days=years * 365)).strftime("%d/%m/%Y")

    if USE_MYSQL:
        c.execute("""
            DELETE FROM responses
            WHERE DATE_FORMAT(STR_TO_DATE(date, '%%d/%%m/%%Y'), '%%Y%%m%%d') < %s
        """, [cutoff.replace("/", "")])
    else:
        # SQLite: compare date strings in dd/mm/yyyy format
        c.execute("""
            DELETE FROM responses
            WHERE (substr(date,7,4)||substr(date,4,2)||substr(date,1,2)) < ?
        """, [cutoff.replace("/", "")])

    deleted = c.rowcount
    conn.commit()
    conn.close()
    return deleted


def seed_defaults():
    """Seed default data if tables are empty."""
    from werkzeug.security import generate_password_hash

    conn = get_db()
    c = get_cursor(conn)

    # Check if users table is empty
    result = fetch_one(conn, "SELECT COUNT(*) as cnt FROM users")
    if result["cnt"] == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.execute(
            "INSERT INTO users (username, password_hash, role, display_name, created_at) VALUES (?, ?, ?, ?, ?)",
            ("admin", generate_password_hash("admin123"), "admin", "Admin", now)
        )
        c.execute(
            "INSERT INTO users (username, password_hash, role, display_name, created_at) VALUES (?, ?, ?, ?, ?)",
            ("user", generate_password_hash("user123"), "user", "User", now)
        )

    # Check if tasks table is empty
    result = fetch_one(conn, "SELECT COUNT(*) as cnt FROM tasks")
    if result["cnt"] == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        default_tasks = [
            ("lubes", "LUBES", "LITRE", 1),
            ("combo", "COMBO", "BAG", 2),
            ("pastry", "PASTRY", "PCS", 3),
            ("shellapp", "SHELL APP", "REG", 4),
        ]
        for key, label, unit, order in default_tasks:
            c.execute(
                "INSERT INTO tasks (task_key, label, unit, sort_order, active, created_at) VALUES (?, ?, ?, ?, 1, ?)",
                (key, label, unit, order, now)
            )

    # Check if team_members table is empty
    result = fetch_one(conn, "SELECT COUNT(*) as cnt FROM team_members")
    if result["cnt"] == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        default_members = [
            ("AMIN", "AMIN", "M", '{"lubes":5,"combo":2,"shellapp":2}', '["lubes","combo","shellapp"]'),
            ("SAIDUL", "SAIDUL", "M", '{"lubes":5,"combo":2,"shellapp":2}', '["lubes","combo","shellapp"]'),
            ("ANIK", "ANIK", "F", '{"lubes":4,"combo":2,"shellapp":2}', '["lubes","combo","shellapp"]'),
            ("SIRAZUL", "SIRAZUL", "M", '{"lubes":5,"combo":2,"shellapp":2}', '["lubes","combo","shellapp"]'),
            ("ALIFF", "ALIFF", "M", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
            ("NORAINI", "NORAINI", "F", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
            ("DIMAS", "DIMAS", "M", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
            ("AIREL", "AIREL", "M", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
            ("IBRAHIM", "IBRAHIM", "M", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
            ("HADIF", "HADIF", "M", '{"combo":2,"pastry":2,"shellapp":2}', '["combo","pastry","shellapp"]'),
        ]
        for staff_id, name, gender, targets, tasks in default_members:
            c.execute(
                "INSERT INTO team_members (staff_id, name, gender, targets, tasks, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (staff_id, name, gender, targets, tasks, now)
            )

    conn.commit()
    conn.close()


if __name__ == "__main__":
    print(f"Database mode: {'MySQL' if USE_MYSQL else 'SQLite'}")
    init_db()
    print("Database initialized successfully!")
    seed_defaults()
    print("Default data seeded!")
    deleted = delete_old_responses(years=2)
    print(f"Deleted {deleted} old responses (older than 2 years)")
