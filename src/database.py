from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

BASE_DIR = Path(__file__).resolve().parent.parent
DATABASE_DIR = BASE_DIR / "database"
DATABASE_PATH = DATABASE_DIR / "music.db"
SCHEMA_PATH = DATABASE_DIR / "schema.sql"

REQUIRED_TABLES = {
    "countries",
    "languages",
    "artists",
    "songs",
    "song_artists",
    "song_countries",
    "song_languages",
    "rankings",
    "ranking_entries",
    "streaming_entries",
}


@contextmanager
def get_connection(db_path: str | Path | None = None) -> Iterator[sqlite3.Connection]:
    """Открывает соединение и гарантированно включает проверку внешних ключей."""
    path = Path(db_path) if db_path is not None else DATABASE_PATH
    if str(path) != ":memory:":
        path.parent.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(str(path))
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 5000")
    try:
        yield connection
    finally:
        connection.close()


def initialize_database(db_path: str | Path | None = None) -> None:
    """Создаёт согласованную схему; повторный вызов безопасен."""
    schema_path = SCHEMA_PATH
    if not schema_path.exists():
        raise FileNotFoundError(f"Файл схемы базы данных не найден: {schema_path}")

    path = Path(db_path) if db_path is not None else DATABASE_PATH
    if str(path) != ":memory:":
        path.parent.mkdir(parents=True, exist_ok=True)

    schema_sql = schema_path.read_text(encoding="utf-8")
    with get_connection(path) as connection:
        connection.executescript(schema_sql)


def get_tables(db_path: str | Path | None = None) -> list[str]:
    with get_connection(db_path) as connection:
        rows = connection.execute(
            """
            SELECT name
            FROM sqlite_master
            WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
            ORDER BY name;
            """
        ).fetchall()
    return [row["name"] for row in rows]
