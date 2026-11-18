"""Импорт и анализ прогнозов годовых музыкальных Top-50."""
from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from statistics import mean
from typing import Iterable

from openpyxl import load_workbook

from src.database import DATABASE_PATH, get_connection, initialize_database

BASE_DIR = Path(__file__).resolve().parent.parent
FORECAST_WORKBOOK_PATH = BASE_DIR / "data" / "music_predictions.xlsx"
REQUIRED_COLUMNS = ("Исполнители", "Песня", "Язык", "Страна", "Год")


def _norm(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _split(value: object) -> tuple[str, ...]:
    if value is None:
        return ()
    return tuple(sorted({part.strip() for part in str(value).split(";") if part.strip()}, key=str.casefold))


def _identity(title: str, year: int, artists: Iterable[str]) -> tuple:
    return (_norm(title), int(year), tuple(sorted((_norm(a) for a in artists))))


@dataclass(frozen=True)
class ForecastMetric:
    year: int
    predicted_count: int
    actual_count: int
    hits: int
    misses: int
    unexpected: int
    precision: float
    recall: float
    mae: float | None
    exact_matches: int
    within_5: int
    within_10: int


@dataclass(frozen=True)
class ForecastComparison:
    forecast_year: int
    predicted_position: int
    title: str
    artists: tuple[str, ...]
    release_year: int
    actual_position: int | None
    position_error: int | None
    status: str
    first_actual_year: int | None
    delay_years: int | None
    first_actual_position: int | None = None


def _ensure_forecast_tables(connection: sqlite3.Connection) -> None:
    connection.executescript("""
    CREATE TABLE IF NOT EXISTS forecast_years (
        forecast_year INTEGER PRIMARY KEY CHECK (forecast_year BETWEEN 1900 AND 2100)
    );
    CREATE TABLE IF NOT EXISTS forecast_entries (
        forecast_entry_id INTEGER PRIMARY KEY,
        forecast_year INTEGER NOT NULL,
        predicted_position INTEGER NOT NULL CHECK (predicted_position > 0),
        title TEXT NOT NULL,
        release_year INTEGER NOT NULL CHECK (release_year BETWEEN 1900 AND 2100),
        artists_json TEXT NOT NULL,
        languages_json TEXT NOT NULL,
        countries_json TEXT NOT NULL,
        identity_key TEXT NOT NULL,
        UNIQUE (forecast_year, predicted_position),
        UNIQUE (forecast_year, identity_key),
        FOREIGN KEY (forecast_year) REFERENCES forecast_years(forecast_year) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_forecast_entries_identity ON forecast_entries(identity_key);
    """)


def import_forecasts(
    workbook_path: str | Path = FORECAST_WORKBOOK_PATH,
    db_path: str | Path | None = None,
    replace_existing: bool = True,
) -> int:
    """Импортирует листы с годами; номер строки данных становится прогнозной позицией."""
    path = Path(workbook_path)
    if not path.is_file():
        raise FileNotFoundError(f"Файл прогнозов не найден: {path}")
    book = load_workbook(path, read_only=True, data_only=True)
    parsed: dict[int, list[dict]] = {}
    errors: list[str] = []
    for sheet in book.worksheets:
        try:
            year = int(sheet.title)
        except ValueError:
            errors.append(f"Лист '{sheet.title}': название должно быть годом")
            continue
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if header is None or tuple(str(v).strip() for v in header[:5]) != REQUIRED_COLUMNS:
            errors.append(f"Лист {year}: ожидаются столбцы {', '.join(REQUIRED_COLUMNS)}")
            continue
        entries = []
        seen = set()
        for position, row in enumerate(rows, start=1):
            if not any(v is not None and str(v).strip() for v in row[:5]):
                continue
            if len(row) < 5 or any(row[i] is None or not str(row[i]).strip() for i in range(5)):
                errors.append(f"Лист {year}, строка {position + 1}: заполните все пять полей")
                continue
            artists = _split(row[0])
            try:
                release_year = int(row[4])
            except (TypeError, ValueError):
                errors.append(f"Лист {year}, строка {position + 1}: некорректный год выпуска")
                continue
            if not artists or not 1900 <= release_year <= 2100:
                errors.append(f"Лист {year}, строка {position + 1}: некорректные исполнители или год")
                continue
            title = str(row[1]).strip()
            key = _identity(title, release_year, artists)
            if key in seen:
                errors.append(f"Лист {year}: повтор песни '{title}'")
            seen.add(key)
            entries.append({
                "position": position, "title": title, "release_year": release_year,
                "artists": artists, "languages": _split(row[2]), "countries": _split(row[3]),
                "identity_key": json.dumps(key, ensure_ascii=False),
            })
        parsed[year] = entries
    book.close()
    if errors:
        raise ValueError("Ошибки в книге прогнозов:\n- " + "\n- ".join(errors))
    if not parsed:
        raise ValueError("В книге не найдено ни одного листа с годом")

    initialize_database(db_path)
    with get_connection(db_path) as connection:
        _ensure_forecast_tables(connection)
        for year, entries in parsed.items():
            exists = connection.execute("SELECT 1 FROM forecast_years WHERE forecast_year=?", (year,)).fetchone()
            if exists and not replace_existing:
                continue
            if exists:
                connection.execute("DELETE FROM forecast_years WHERE forecast_year=?", (year,))
            connection.execute("INSERT INTO forecast_years(forecast_year) VALUES (?)", (year,))
            connection.executemany("""INSERT INTO forecast_entries
                (forecast_year,predicted_position,title,release_year,artists_json,languages_json,countries_json,identity_key)
                VALUES (:year,:position,:title,:release_year,:artists,:languages,:countries,:identity_key)""",
                [{**e, "year": year, "artists": json.dumps(e["artists"], ensure_ascii=False),
                  "languages": json.dumps(e["languages"], ensure_ascii=False),
                  "countries": json.dumps(e["countries"], ensure_ascii=False)} for e in entries])
        connection.commit()
    return sum(len(v) for v in parsed.values())


def _actual_entries(connection: sqlite3.Connection) -> tuple[dict[int, dict], dict[tuple, dict[int, int]]]:
    years = {r["ranking_year"]: r["top_size"] for r in connection.execute("SELECT ranking_year,top_size FROM rankings")}
    rows = connection.execute("""SELECT r.ranking_year,re.position,s.title,s.release_year,a.artist_name
        FROM ranking_entries re JOIN rankings r USING(ranking_id) JOIN songs s USING(song_id)
        LEFT JOIN song_artists sa USING(song_id) LEFT JOIN artists a USING(artist_id)
        ORDER BY r.ranking_year,re.position""").fetchall()
    grouped: dict[tuple, dict[int, int]] = {}
    temp: dict[tuple, set[str]] = {}
    meta: dict[tuple, tuple[str,int]] = {}
    for row in rows:
        base = (_norm(row["title"]), int(row["release_year"]))
        temp.setdefault((base, row["ranking_year"], row["position"]), set())
        if row["artist_name"]:
            temp[(base, row["ranking_year"], row["position"])].add(row["artist_name"])
        meta[base] = (row["title"], int(row["release_year"]))
    for (base, year, position), artists in temp.items():
        key = (base[0], base[1], tuple(sorted(_norm(a) for a in artists)))
        grouped.setdefault(key, {})[year] = position
    return years, grouped


def get_forecast_analysis(db_path: str | Path | None = None) -> tuple[list[ForecastMetric], list[ForecastComparison], list[ForecastComparison]]:
    """Возвращает метрики по годам, сравнение прогнозов и неожиданные попадания.

    Отложенное попадание определяется относительно первого прогноза песни.
    Повторные прогнозы остаются отдельными строками и оцениваются в своём году.
    """
    with get_connection(db_path) as connection:
        _ensure_forecast_tables(connection)
        actual_year_sizes, actual = _actual_entries(connection)
        forecast_rows = connection.execute("SELECT * FROM forecast_entries ORDER BY forecast_year,predicted_position").fetchall()
    if not forecast_rows:
        return [], [], []
    forecasts_by_key: dict[str, list[sqlite3.Row]] = {}
    parsed_rows = []
    for row in forecast_rows:
        artists = tuple(json.loads(row["artists_json"]))
        key = _identity(row["title"], row["release_year"], artists)
        forecasts_by_key.setdefault(json.dumps(key, ensure_ascii=False), []).append(row)
        parsed_rows.append((row, key, artists))
    first_forecast = {key: min(r["forecast_year"] for r in rows) for key, rows in forecasts_by_key.items()}
    comparisons: list[ForecastComparison] = []
    metrics: list[ForecastMetric] = []
    unexpected: list[ForecastComparison] = []
    forecast_years = sorted({r["forecast_year"] for r in forecast_rows})
    for year in forecast_years:
        year_rows = [item for item in parsed_rows if item[0]["forecast_year"] == year]
        hits = 0; errors = []; exact = within5 = within10 = 0
        forecast_keys = {item[1] for item in year_rows}
        actual_this_year = {k: positions[year] for k, positions in actual.items() if year in positions}
        for row, key, artists in year_rows:
            actual_pos = actual.get(key, {}).get(year)
            first_actual = min((y for y in actual.get(key, {}) if y >= first_forecast[json.dumps(key, ensure_ascii=False)]), default=None)
            if actual_pos is not None:
                hits += 1
                error = actual_pos - row["predicted_position"]
                errors.append(abs(error))
                exact += error == 0; within5 += abs(error) <= 5; within10 += abs(error) <= 10
                status = "HIT"
            else:
                error = None
                status = "DELAYED" if first_actual is not None else "PENDING"
            comparisons.append(ForecastComparison(year,row["predicted_position"],row["title"],artists,
                row["release_year"],actual_pos,error,status,first_actual,
                first_actual - first_forecast[json.dumps(key, ensure_ascii=False)] if first_actual else None,
                actual.get(key, {}).get(first_actual) if first_actual is not None else None))
        for key, pos in actual_this_year.items():
            if key not in forecast_keys:
                # Unexpected actual entrant; no predicted position exists.
                title = next((r["title"] for r,k,_ in parsed_rows if k == key), "")
                artists = next((a for r,k,a in parsed_rows if k == key), ())
                if not title:
                    with get_connection(db_path) as connection:
                        song = connection.execute("""SELECT s.title,s.release_year,GROUP_CONCAT(a.artist_name,';') artists
                            FROM songs s JOIN ranking_entries re USING(song_id) JOIN rankings r USING(ranking_id)
                            LEFT JOIN song_artists sa USING(song_id) LEFT JOIN artists a USING(artist_id)
                            WHERE r.ranking_year=? AND re.position=? GROUP BY s.song_id""",(year,pos)).fetchone()
                    if song:
                        title=song["title"]; artists=tuple((song["artists"] or "").split(";"))
                unexpected.append(ForecastComparison(year,0,title,artists,key[1],pos,None,"UNEXPECTED",year,0))
        count_actual = len(actual_this_year)
        metrics.append(ForecastMetric(year,len(year_rows),count_actual,hits,len(year_rows)-hits,
            count_actual-hits,hits/len(year_rows) if year_rows else 0,
            hits/count_actual if count_actual else 0,mean(errors) if errors else None,
            int(exact),int(within5),int(within10)))
    comparisons.sort(key=lambda x:(x.forecast_year,x.predicted_position))
    unexpected.sort(key=lambda x:(x.forecast_year,x.actual_position or 999))
    return metrics, comparisons, unexpected


def get_delayed_forecasts(db_path: str | Path | None = None) -> list[ForecastComparison]:
    """Песни, впервые попавшие в рейтинг позже года первого прогноза."""
    _, comparisons, _ = get_forecast_analysis(db_path)
    unique: dict[tuple, ForecastComparison] = {}
    for item in comparisons:
        if item.first_actual_year is not None and item.delay_years and item.delay_years > 0:
            key = _identity(item.title,item.release_year,item.artists)
            current = unique.get(key)
            if current is None or item.forecast_year < current.forecast_year:
                unique[key] = item
    return sorted(unique.values(), key=lambda x:(x.delay_years or 0,x.title.casefold()))
