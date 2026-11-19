"""Импорт и аналитика ежегодного стримингового Top-5."""
from __future__ import annotations

import math
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from src.database import BASE_DIR, DATABASE_PATH, get_connection, initialize_database
from src.importer import normalize_value

WORKBOOK_PATH = BASE_DIR / "data" / "music_streaming.xlsx"


class StreamingImportError(ValueError):
    """Ошибка структуры или содержимого книги стримингов."""


@dataclass(frozen=True)
class StreamingEntry:
    year: int
    artist: str
    position: int
    listening_minutes: int | None


def read_streaming_workbook(path: str | Path = WORKBOOK_PATH) -> list[StreamingEntry]:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"Книга стримингов не найдена: {path}")
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise StreamingImportError(f"Не удалось открыть Excel-книгу: {exc}") from exc

    entries: list[StreamingEntry] = []
    seen: set[tuple[int, str]] = set()
    try:
        if workbook.sheetnames != ["Стриминги"]:
            raise StreamingImportError("Ожидается один лист с названием «Стриминги»")
        ws = workbook["Стриминги"]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        expected = ("Год", "Исполнитель", "Время прослушивания")
        if header is None or tuple(str(x).strip() if x is not None else "" for x in header[:3]) != expected or len(header) != 3:
            raise StreamingImportError("Ожидаются три столбца: «Год», «Исполнитель», «Время прослушивания»")
        positions: dict[int, int] = {}
        for row_number, row in enumerate(rows, start=2):
            if len(row) < 3 or all(value is None for value in row):
                continue
            if len(row) != 3:
                raise StreamingImportError(f"Строка {row_number}: ожидалось ровно три значения")
            year_value, artist_value, minutes_value = row
            if isinstance(year_value, bool) or not isinstance(year_value, (int, float)) or not float(year_value).is_integer() or not 1900 <= int(year_value) <= 2100:
                raise StreamingImportError(f"Строка {row_number}: некорректный год")
            year = int(year_value)
            if not isinstance(artist_value, str) or not artist_value.strip():
                raise StreamingImportError(f"Строка {row_number}: исполнитель обязателен")
            artist = " ".join(artist_value.split())
            key = (year, normalize_value(artist))
            if key in seen:
                raise StreamingImportError(f"Строка {row_number}: повтор исполнителя «{artist}» за {year} год")
            seen.add(key)
            if minutes_value is None:
                minutes = None
            elif isinstance(minutes_value, bool) or not isinstance(minutes_value, (int, float)) or not math.isfinite(float(minutes_value)) or not float(minutes_value).is_integer() or minutes_value < 0:
                raise StreamingImportError(f"Строка {row_number}: время должно быть целым числом минут или пустым")
            else:
                minutes = int(minutes_value)
            positions[year] = positions.get(year, 0) + 1
            if positions[year] > 5:
                raise StreamingImportError(f"В {year} году допускается не более пяти исполнителей")
            entries.append(StreamingEntry(year, artist, positions[year], minutes))
    finally:
        workbook.close()
    return entries


def import_streaming_workbook(path: str | Path = WORKBOOK_PATH, db_path: str | Path | None = None) -> int:
    """Проверяет книгу и атомарно заменяет импортированную стриминговую историю."""
    entries = read_streaming_workbook(path)
    database_path = Path(db_path) if db_path is not None else DATABASE_PATH
    if not database_path.exists():
        initialize_database(database_path)
    with get_connection(database_path) as connection:
        connection.execute("BEGIN IMMEDIATE")
        try:
            artist_rows = connection.execute("SELECT artist_id, artist_name FROM artists").fetchall()
            artist_ids = {normalize_value(row["artist_name"]): row["artist_id"] for row in artist_rows}
            missing = sorted({entry.artist for entry in entries if normalize_value(entry.artist) not in artist_ids})
            if missing:
                raise StreamingImportError("Исполнители отсутствуют в общем справочнике artists: " + ", ".join(missing))
            connection.execute("DELETE FROM streaming_entries")
            connection.executemany(
                "INSERT INTO streaming_entries(year, artist_id, position, listening_minutes) VALUES (?, ?, ?, ?)",
                [(e.year, artist_ids[normalize_value(e.artist)], e.position, e.listening_minutes) for e in entries],
            )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
    return len(entries)


def get_streaming_statistics(db_path: str | Path | None = None) -> list[dict[str, object]]:
    with get_connection(db_path) as connection:
        rows = connection.execute("""
            SELECT se.year, se.position, a.artist_name AS artist, se.listening_minutes
            FROM streaming_entries se JOIN artists a USING (artist_id)
            ORDER BY se.year, se.position
        """).fetchall()
    return [dict(row) for row in rows]


def main() -> int:
    try:
        count = import_streaming_workbook()
        print(f"Импортировано стриминговых записей: {count}")
        return 0
    except (StreamingImportError, OSError, sqlite3.Error) as exc:
        print(f"Ошибка импорта стримингов: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
