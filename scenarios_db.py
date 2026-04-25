"""
scenarios_db.py — SQLite persistence for DCF scenario management.

For production persistence on Render (free tier has ephemeral filesystem),
set the SCENARIOS_DB_PATH environment variable to a path on Render's
persistent disk, or migrate to Render's free Postgres.
"""

import sqlite3
import json
import os
import uuid
import datetime

DB_PATH = os.environ.get("SCENARIOS_DB_PATH") or os.path.join(
    os.path.dirname(__file__), "scenarios.db"
)


def _conn():
    """Return a new connection with WAL mode and row-factory enabled."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    """Create table if not exists.  Call once on app startup."""
    with _conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS scenarios (
                id          TEXT PRIMARY KEY,
                ticker      TEXT NOT NULL,
                name        TEXT NOT NULL,
                created_at  TEXT NOT NULL,
                updated_at  TEXT NOT NULL,
                inputs      TEXT NOT NULL,
                outputs     TEXT NOT NULL,
                UNIQUE(ticker, name)
            )
            """
        )


def save_scenario(ticker, name, inputs, outputs):
    """Insert or replace scenario.  Returns scenario id."""
    ticker = ticker.upper().strip()
    now = datetime.datetime.utcnow().isoformat() + "Z"
    sid = str(uuid.uuid4())
    inputs_json = json.dumps(inputs)
    outputs_json = json.dumps(outputs)

    with _conn() as conn:
        # Check for existing row (upsert by ticker+name)
        row = conn.execute(
            "SELECT id, created_at FROM scenarios WHERE ticker=? AND name=?",
            (ticker, name),
        ).fetchone()
        if row:
            sid = row["id"]
            created = row["created_at"]
            conn.execute(
                """UPDATE scenarios
                   SET inputs=?, outputs=?, updated_at=?
                   WHERE id=?""",
                (inputs_json, outputs_json, now, sid),
            )
        else:
            conn.execute(
                """INSERT INTO scenarios (id, ticker, name, created_at, updated_at, inputs, outputs)
                   VALUES (?,?,?,?,?,?,?)""",
                (sid, ticker, name, now, now, inputs_json, outputs_json),
            )
    return sid


def list_scenarios(ticker):
    """Return list of scenario dicts for a ticker, newest first."""
    ticker = ticker.upper().strip()
    with _conn() as conn:
        if ticker == "ALL":
            rows = conn.execute(
                "SELECT * FROM scenarios ORDER BY updated_at DESC"
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM scenarios WHERE ticker=? ORDER BY updated_at DESC",
                (ticker,),
            ).fetchall()
    return [_row_to_dict(r) for r in rows]


def delete_scenario(ticker, name):
    """Delete by ticker + name."""
    ticker = ticker.upper().strip()
    with _conn() as conn:
        conn.execute(
            "DELETE FROM scenarios WHERE ticker=? AND name=?", (ticker, name)
        )


def get_scenario(ticker, name):
    """Return single scenario dict or None."""
    ticker = ticker.upper().strip()
    with _conn() as conn:
        row = conn.execute(
            "SELECT * FROM scenarios WHERE ticker=? AND name=?", (ticker, name)
        ).fetchone()
    return _row_to_dict(row) if row else None


def _row_to_dict(row):
    """Convert a sqlite3.Row to a plain dict, parsing JSON fields."""
    d = dict(row)
    d["inputs"] = json.loads(d["inputs"])
    d["outputs"] = json.loads(d["outputs"])
    return d
