from __future__ import annotations

import json
import sqlite3
from datetime import UTC, datetime, timedelta
from pathlib import Path

from pydantic import TypeAdapter

from bdr_wizard.models import ImportSession


class SessionRepository:
    def __init__(self, db_path: Path = Path("data/bdr_wizard.sqlite3")) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _init_db(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS import_sessions (
                    id TEXT PRIMARY KEY,
                    payload TEXT NOT NULL,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )

    def save(self, session: ImportSession) -> None:
        payload = session.model_dump_json()
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO import_sessions (id, payload, updated_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(id) DO UPDATE SET
                    payload = excluded.payload,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (session.id, payload),
            )

    def get(self, session_id: str) -> ImportSession | None:
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT payload FROM import_sessions WHERE id = ?",
                (session_id,),
            ).fetchone()
        if row is None:
            return None
        data = json.loads(row[0])
        return TypeAdapter(ImportSession).validate_python(data)

    def delete_older_than(self, days: int) -> int:
        cutoff = (datetime.now(UTC) - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute(
                "DELETE FROM import_sessions WHERE updated_at < datetime(?)",
                (cutoff,),
            )
            return cursor.rowcount
