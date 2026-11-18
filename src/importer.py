from __future__ import annotations

import math
import sqlite3
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .database import BASE_DIR, DATABASE_PATH, REQUIRED_TABLES, get_connection, initialize_database

WORKBOOK_PATH = BASE_DIR / "data" / "music_statistics.xlsx"
TOP_SIZE = 50
MIN_RELEASE_YEAR = 1900
MAX_RELEASE_YEAR = 2100

HEADER_ALIASES = {
    "artists": {"исполнители", "артисты"},
    "title": {"песня", "название"},
    "languages": {"язык", "языки"},
    "countries": {"страна", "страны"},
    "release_year": {"год", "год выпуска"},
}
EXPECTED_COLUMN_ORDER = ("artists", "title", "languages", "countries", "release_year")


class ImporterError(Exception):
    """Базовая ошибка импортёра."""


class WorkbookValidationError(ImporterError):
    def __init__(self, errors: Iterable[str]):
        self.errors = list(errors)
        super().__init__("Книга Excel не прошла проверку")


class ImportConflictError(ImporterError):
    def __init__(self, conflicts: Iterable[str]):
        self.conflicts = list(conflicts)
        super().__init__("Обнаружены конфликты; база данных не изменена")


class DatabaseSchemaError(ImporterError):
    pass


@dataclass(frozen=True)
class SongIdentity:
    normalized_title: str
    release_year: int
    normalized_artists: frozenset[str]


@dataclass(frozen=True)
class ParsedSong:
    ranking_year: int
    position: int
    excel_row: int
    title: str
    release_year: int
    artists: tuple[str, ...]
    languages: tuple[str, ...]
    countries: tuple[str, ...]

    @property
    def identity(self) -> SongIdentity:
        return SongIdentity(
            normalized_title=normalize_value(self.title),
            release_year=self.release_year,
            normalized_artists=frozenset(normalize_value(x) for x in self.artists),
        )

    @property
    def title_year(self) -> tuple[str, int]:
        return normalize_value(self.title), self.release_year


@dataclass
class WorkbookData:
    songs: list[ParsedSong] = field(default_factory=list)
    ranking_years: list[int] = field(default_factory=list)
    top_size: int = TOP_SIZE

    @property
    def unique_songs(self) -> dict[SongIdentity, ParsedSong]:
        result: dict[SongIdentity, ParsedSong] = {}
        for song in self.songs:
            result.setdefault(song.identity, song)
        return result


@dataclass
class DatabaseSong:
    song_id: int
    title: str
    release_year: int
    artists: tuple[str, ...]
    languages: tuple[str, ...]
    countries: tuple[str, ...]

    @property
    def identity(self) -> SongIdentity:
        return SongIdentity(
            normalized_title=normalize_value(self.title),
            release_year=self.release_year,
            normalized_artists=frozenset(normalize_value(x) for x in self.artists),
        )

    @property
    def title_year(self) -> tuple[str, int]:
        return normalize_value(self.title), self.release_year


@dataclass
class ImportResult:
    rankings_imported: int
    songs_added: int
    songs_reused: int
    entries_added: int
    replaced_years: tuple[int, ...] = ()
    dry_run: bool = False


def normalize_value(value: str) -> str:
    """Нормализация только для сопоставления; исходное написание сохраняется."""
    normalized = unicodedata.normalize("NFKC", value)
    return " ".join(normalized.split()).casefold()


def _clean_text(value: object, *, field_name: str, sheet: str, row: int) -> str:
    if not isinstance(value, str):
        raise ValueError(
            f"лист «{sheet}», строка {row}: поле «{field_name}» должно быть текстом"
        )
    cleaned = " ".join(value.split())
    if not cleaned:
        raise ValueError(
            f"лист «{sheet}», строка {row}: поле «{field_name}» обязательно"
        )
    return cleaned


def _split_multi_value(value: object, *, field_name: str, sheet: str, row: int) -> tuple[str, ...]:
    text = _clean_text(value, field_name=field_name, sheet=sheet, row=row)
    parts = [part.strip() for part in text.split(";")]
    if any(not part for part in parts):
        raise ValueError(
            f"лист «{sheet}», строка {row}: в поле «{field_name}» есть пустой элемент между разделителями «;»"
        )

    seen: set[str] = set()
    result: list[str] = []
    for part in parts:
        cleaned = " ".join(part.split())
        key = normalize_value(cleaned)
        if key in seen:
            raise ValueError(
                f"лист «{sheet}», строка {row}: повторяется значение «{cleaned}» в поле «{field_name}»"
            )
        seen.add(key)
        result.append(cleaned)
    return tuple(result)


def _parse_release_year(value: object, *, sheet: str, row: int) -> int:
    if isinstance(value, bool):
        raise ValueError(f"лист «{sheet}», строка {row}: год выпуска должен быть целым числом")

    year: int
    if isinstance(value, int):
        year = value
    elif isinstance(value, float) and math.isfinite(value) and value.is_integer():
        year = int(value)
    elif isinstance(value, str) and value.strip().isdecimal():
        year = int(value.strip())
    else:
        raise ValueError(f"лист «{sheet}», строка {row}: год выпуска должен быть целым числом")

    if not MIN_RELEASE_YEAR <= year <= MAX_RELEASE_YEAR:
        raise ValueError(
            f"лист «{sheet}», строка {row}: год выпуска {year} вне диапазона "
            f"{MIN_RELEASE_YEAR}–{MAX_RELEASE_YEAR}"
        )
    return year


def _parse_ranking_year(sheet_name: str) -> int:
    if not sheet_name.strip().isdecimal():
        raise ValueError(f"название листа «{sheet_name}» должно содержать год рейтинга")
    year = int(sheet_name.strip())
    if not MIN_RELEASE_YEAR <= year <= MAX_RELEASE_YEAR:
        raise ValueError(f"год рейтинга на листе «{sheet_name}» вне диапазона 1900–2100")
    return year


def _resolve_headers(header_values: tuple[object, ...], sheet_name: str) -> list[int]:
    non_empty = [index for index, value in enumerate(header_values) if value is not None and str(value).strip()]
    if len(non_empty) != 5:
        raise ValueError(
            f"лист «{sheet_name}»: ожидается ровно 5 заголовков, найдено {len(non_empty)}"
        )

    normalized_headers = [normalize_value(str(header_values[index])) for index in non_empty]
    resolved: list[int] = []
    for expected in EXPECTED_COLUMN_ORDER:
        aliases = {normalize_value(alias) for alias in HEADER_ALIASES[expected]}
        matching = [
            column_index
            for column_index, header in zip(non_empty, normalized_headers)
            if header in aliases
        ]
        if len(matching) != 1:
            raise ValueError(
                f"лист «{sheet_name}»: не найден однозначный заголовок для поля "
                f"«{expected}»; проверьте названия и порядок пяти столбцов"
            )
        resolved.append(matching[0])

    if resolved != [0, 1, 2, 3, 4]:
        raise ValueError(
            f"лист «{sheet_name}»: ожидаются ровно пять последовательных столбцов "
            "в порядке «Исполнители», «Песня», «Язык», «Страна», «Год»"
        )
    return resolved


def _row_is_empty(values: tuple[object, ...]) -> bool:
    return all(value is None or (isinstance(value, str) and not value.strip()) for value in values)


def read_workbook(excel_path: str | Path, top_size: int = TOP_SIZE) -> WorkbookData:
    """Читает и полностью проверяет Excel-книгу, не обращаясь к SQLite."""
    if isinstance(top_size, bool) or not isinstance(top_size, int) or top_size <= 0:
        raise ValueError("top_size должен быть положительным целым числом")
    path = Path(excel_path)
    if not path.is_file():
        raise FileNotFoundError(f"Excel-файл не найден: {path}")

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise WorkbookValidationError([f"не удалось открыть Excel-книгу: {exc}"]) from exc

    errors: list[str] = []
    conflicts: list[str] = []
    all_songs: list[ParsedSong] = []
    ranking_years: list[int] = []
    seen_ranking_years: set[int] = set()
    workbook_identity_metadata: dict[SongIdentity, tuple[frozenset[str], frozenset[str], ParsedSong]] = {}
    workbook_title_year_artists: dict[tuple[str, int], tuple[frozenset[str], ParsedSong]] = {}

    try:
        if not workbook.sheetnames:
            raise WorkbookValidationError(["книга не содержит листов"])

        for worksheet in workbook.worksheets:
            sheet_name = worksheet.title
            try:
                ranking_year = _parse_ranking_year(sheet_name)
            except ValueError as exc:
                errors.append(str(exc))
                continue

            if ranking_year in seen_ranking_years:
                errors.append(f"год рейтинга {ranking_year} указан более чем на одном листе")
                continue
            seen_ranking_years.add(ranking_year)
            ranking_years.append(ranking_year)

            rows = worksheet.iter_rows(values_only=True)
            try:
                header_values = tuple(next(rows))
            except StopIteration:
                errors.append(f"лист «{sheet_name}»: лист пустой")
                continue

            try:
                column_indexes = _resolve_headers(header_values, sheet_name)
            except ValueError as exc:
                errors.append(str(exc))
                continue

            # Сначала определяем последнюю непустую строку, чтобы не считать
            # форматирование пустых хвостовых строк частью рейтинга.
            data_rows: list[tuple[int, tuple[object, ...]]] = []
            for excel_row, values in enumerate(rows, start=2):
                values_tuple = tuple(values)
                if not _row_is_empty(values_tuple):
                    data_rows.append((excel_row, values_tuple))

            if not data_rows:
                errors.append(f"лист «{sheet_name}»: нет строк с песнями")
                continue

            last_non_empty_row = data_rows[-1][0]
            rows_by_number = {row_number: values for row_number, values in data_rows}
            sheet_songs: list[ParsedSong] = []
            sheet_identities: dict[SongIdentity, ParsedSong] = {}

            for excel_row in range(2, last_non_empty_row + 1):
                values = rows_by_number.get(excel_row)
                if values is None:
                    errors.append(
                        f"лист «{sheet_name}», строка {excel_row}: пустая строка внутри рейтинга; "
                        "позиции определяются порядком строк"
                    )
                    continue

                if len(values) > 5 and any(value is not None and str(value).strip() for value in values[5:]):
                    errors.append(
                        f"лист «{sheet_name}», строка {excel_row}: обнаружены данные за пределами пяти столбцов"
                    )
                    continue

                try:
                    selected = [values[index] if index < len(values) else None for index in column_indexes]
                    artists = _split_multi_value(
                        selected[0], field_name="Исполнители", sheet=sheet_name, row=excel_row
                    )
                    title = _clean_text(
                        selected[1], field_name="Песня", sheet=sheet_name, row=excel_row
                    )
                    languages = _split_multi_value(
                        selected[2], field_name="Язык", sheet=sheet_name, row=excel_row
                    )
                    countries = _split_multi_value(
                        selected[3], field_name="Страна", sheet=sheet_name, row=excel_row
                    )
                    release_year = _parse_release_year(
                        selected[4], sheet=sheet_name, row=excel_row
                    )
                except ValueError as exc:
                    errors.append(str(exc))
                    continue

                song = ParsedSong(
                    ranking_year=ranking_year,
                    position=excel_row - 1,
                    excel_row=excel_row,
                    title=title,
                    release_year=release_year,
                    artists=artists,
                    languages=languages,
                    countries=countries,
                )
                identity = song.identity

                if identity in sheet_identities:
                    first = sheet_identities[identity]
                    errors.append(
                        f"лист «{sheet_name}»: дубликат песни «{title}» в строках "
                        f"{first.excel_row} и {excel_row}"
                    )
                else:
                    sheet_identities[identity] = song

                metadata = (
                    frozenset(normalize_value(x) for x in languages),
                    frozenset(normalize_value(x) for x in countries),
                )
                previous = workbook_identity_metadata.get(identity)
                if previous and previous[:2] != metadata:
                    old_song = previous[2]
                    conflicts.append(
                        f"конфликт метаданных внутри Excel: «{title}» ({release_year}), "
                        f"листы {old_song.ranking_year} и {ranking_year}; "
                        f"на первом листе язык [{'; '.join(old_song.languages)}], "
                        f"страна [{'; '.join(old_song.countries)}]; "
                        f"на втором листе язык [{'; '.join(languages)}], "
                        f"страна [{'; '.join(countries)}]"
                    )
                else:
                    workbook_identity_metadata.setdefault(identity, (*metadata, song))

                title_year = song.title_year
                artist_key = frozenset(normalize_value(x) for x in artists)
                previous_title_year = workbook_title_year_artists.get(title_year)
                if previous_title_year and previous_title_year[0] != artist_key:
                    old_song = previous_title_year[1]
                    conflicts.append(
                        f"возможный конфликт идентичности внутри Excel: «{title}» ({release_year}); "
                        f"на листе {old_song.ranking_year} исполнители "
                        f"[{'; '.join(old_song.artists)}], на листе {ranking_year} — "
                        f"[{'; '.join(artists)}]. Проверьте вручную"
                    )
                else:
                    workbook_title_year_artists.setdefault(title_year, (artist_key, song))

                sheet_songs.append(song)

            if len(sheet_songs) != top_size:
                errors.append(
                    f"лист «{sheet_name}»: ожидалось ровно {top_size} песен, "
                    f"найдено корректно разобранных строк: {len(sheet_songs)}"
                )
            all_songs.extend(sheet_songs)
    finally:
        workbook.close()

    if errors:
        raise WorkbookValidationError(errors)
    if conflicts:
        raise ImportConflictError(conflicts)

    return WorkbookData(
        songs=all_songs,
        ranking_years=sorted(ranking_years),
        top_size=top_size,
    )


def _ensure_database_ready(db_path: str | Path) -> None:
    path = Path(db_path)
    if not path.exists() or path.stat().st_size == 0:
        initialize_database(path)
        return

    with get_connection(path) as connection:
        tables = {
            row["name"]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
            ).fetchall()
        }

    if not tables:
        initialize_database(path)
        return

    missing = REQUIRED_TABLES - tables
    if missing:
        raise DatabaseSchemaError(
            "В существующей базе отсутствуют таблицы: " + ", ".join(sorted(missing))
        )


def _load_database_songs(connection: sqlite3.Connection) -> list[DatabaseSong]:
    song_rows = connection.execute(
        "SELECT song_id, title, release_year FROM songs ORDER BY song_id"
    ).fetchall()
    songs: dict[int, DatabaseSong] = {
        row["song_id"]: DatabaseSong(
            song_id=row["song_id"],
            title=row["title"],
            release_year=row["release_year"],
            artists=(),
            languages=(),
            countries=(),
        )
        for row in song_rows
    }

    relation_queries = (
        (
            """
            SELECT sa.song_id, a.artist_name AS value
            FROM song_artists sa JOIN artists a ON a.artist_id = sa.artist_id
            ORDER BY sa.song_id, a.artist_name
            """,
            "artists",
        ),
        (
            """
            SELECT sl.song_id, l.language_name AS value
            FROM song_languages sl JOIN languages l ON l.language_id = sl.language_id
            ORDER BY sl.song_id, l.language_name
            """,
            "languages",
        ),
        (
            """
            SELECT sc.song_id, c.country_name AS value
            FROM song_countries sc JOIN countries c ON c.country_id = sc.country_id
            ORDER BY sc.song_id, c.country_name
            """,
            "countries",
        ),
    )
    collected: dict[str, dict[int, list[str]]] = {
        name: defaultdict(list) for _, name in relation_queries
    }
    for query, relation_name in relation_queries:
        for row in connection.execute(query).fetchall():
            collected[relation_name][row["song_id"]].append(row["value"])

    for song_id, song in songs.items():
        song.artists = tuple(collected["artists"].get(song_id, []))
        song.languages = tuple(collected["languages"].get(song_id, []))
        song.countries = tuple(collected["countries"].get(song_id, []))
        if not song.artists or not song.languages or not song.countries:
            raise DatabaseSchemaError(
                f"Нарушена целостность каталога: у песни song_id={song_id} "
                "отсутствуют исполнители, язык или страна"
            )

    return list(songs.values())


def _load_entity_ids(
    connection: sqlite3.Connection,
    table: str,
    id_column: str,
    name_column: str,
) -> dict[str, int]:
    result: dict[str, int] = {}
    for row in connection.execute(
        f"SELECT {id_column}, {name_column} FROM {table} ORDER BY {id_column}"
    ).fetchall():
        key = normalize_value(row[name_column])
        if key in result:
            raise DatabaseSchemaError(
                f"В таблице {table} есть значения, совпадающие после нормализации: "
                f"{row[name_column]!r}"
            )
        result[key] = row[id_column]
    return result


def _compare_with_database(
    workbook: WorkbookData,
    database_songs: list[DatabaseSong],
) -> tuple[dict[SongIdentity, DatabaseSong], list[str]]:
    by_identity: dict[SongIdentity, DatabaseSong] = {}
    by_title_year: dict[tuple[str, int], DatabaseSong] = {}
    conflicts: list[str] = []

    for db_song in database_songs:
        if db_song.identity in by_identity:
            conflicts.append(
                f"в базе уже есть дубликаты идентичности песни «{db_song.title}» "
                f"({db_song.release_year}); проверьте записи song_id "
                f"{by_identity[db_song.identity].song_id} и {db_song.song_id}"
            )
        else:
            by_identity[db_song.identity] = db_song

        title_year = db_song.title_year
        previous_title_year = by_title_year.get(title_year)
        if previous_title_year and previous_title_year.identity != db_song.identity:
            conflicts.append(
                f"в базе есть песни с одинаковыми названием и годом выпуска, "
                f"но разными исполнителями: «{db_song.title}» ({db_song.release_year}); "
                f"song_id {previous_title_year.song_id} и {db_song.song_id}"
            )
        else:
            by_title_year[title_year] = db_song

    for song in workbook.unique_songs.values():
        existing = by_identity.get(song.identity)
        if existing:
            old_languages = {normalize_value(x) for x in existing.languages}
            new_languages = {normalize_value(x) for x in song.languages}
            old_countries = {normalize_value(x) for x in existing.countries}
            new_countries = {normalize_value(x) for x in song.countries}
            if old_languages != new_languages or old_countries != new_countries:
                conflicts.append(
                    f"«{song.title}» ({song.release_year}), исполнители: "
                    f"{'; '.join(song.artists)} — метаданные в базе отличаются от Excel; "
                    f"база: язык [{'; '.join(existing.languages)}], страна "
                    f"[{'; '.join(existing.countries)}]; Excel: язык "
                    f"[{'; '.join(song.languages)}], страна [{'; '.join(song.countries)}]"
                )
            continue

        existing_title_year = by_title_year.get(song.title_year)
        if existing_title_year:
            conflicts.append(
                f"возможный конфликт идентичности: «{song.title}» ({song.release_year}); "
                f"в базе исполнители [{'; '.join(existing_title_year.artists)}], "
                f"в Excel [{'; '.join(song.artists)}]. Проверьте вручную"
            )

    return by_identity, conflicts


def _check_existing_ranking_years(
    connection: sqlite3.Connection,
    ranking_years: Iterable[int],
    replace_existing: bool,
) -> tuple[int, ...]:
    requested = sorted(set(ranking_years))
    placeholders = ",".join("?" for _ in requested)
    rows = connection.execute(
        f"SELECT ranking_year FROM rankings WHERE ranking_year IN ({placeholders}) ORDER BY ranking_year",
        requested,
    ).fetchall()
    existing = tuple(row["ranking_year"] for row in rows)
    if existing and not replace_existing:
        years = ", ".join(str(year) for year in existing)
        raise ImportConflictError(
            [
                f"рейтинги за {years} уже существуют. Повторный импорт запрещён; "
                "для полной замены вызовите import_excel с replace_existing=True"
            ]
        )
    return existing


def _get_or_create_entity(
    connection: sqlite3.Connection,
    cache: dict[str, int],
    *,
    table: str,
    name_column: str,
    value: str,
) -> int:
    key = normalize_value(value)
    existing_id = cache.get(key)
    if existing_id is not None:
        return existing_id

    cursor = connection.execute(
        f"INSERT INTO {table} ({name_column}) VALUES (?)",
        (value,),
    )
    entity_id = int(cursor.lastrowid)
    cache[key] = entity_id
    return entity_id


def _insert_song(
    connection: sqlite3.Connection,
    song: ParsedSong,
    artist_ids: dict[str, int],
    language_ids: dict[str, int],
    country_ids: dict[str, int],
) -> int:
    cursor = connection.execute(
        "INSERT INTO songs (title, release_year) VALUES (?, ?)",
        (song.title, song.release_year),
    )
    song_id = int(cursor.lastrowid)

    for artist in song.artists:
        artist_id = _get_or_create_entity(
            connection,
            artist_ids,
            table="artists",
            name_column="artist_name",
            value=artist,
        )
        connection.execute(
            "INSERT INTO song_artists (song_id, artist_id) VALUES (?, ?)",
            (song_id, artist_id),
        )

    for language in song.languages:
        language_id = _get_or_create_entity(
            connection,
            language_ids,
            table="languages",
            name_column="language_name",
            value=language,
        )
        connection.execute(
            "INSERT INTO song_languages (song_id, language_id) VALUES (?, ?)",
            (song_id, language_id),
        )

    for country in song.countries:
        country_id = _get_or_create_entity(
            connection,
            country_ids,
            table="countries",
            name_column="country_name",
            value=country,
        )
        connection.execute(
            "INSERT INTO song_countries (song_id, country_id) VALUES (?, ?)",
            (song_id, country_id),
        )

    return song_id


def import_excel(
    excel_path: str | Path,
    db_path: str | Path | None = None,
    *,
    replace_existing: bool = False,
    dry_run: bool = False,
    top_size: int = TOP_SIZE,
) -> ImportResult:
    """Проверяет книгу и импортирует рейтинги одной транзакцией.

    По умолчанию повторный импорт любого существующего года считается ошибкой.
    Полная замена возможна только при явном replace_existing=True.
    """
    workbook = read_workbook(excel_path, top_size=top_size)
    path = Path(db_path) if db_path is not None else DATABASE_PATH

    if dry_run:
        if not path.exists() or path.stat().st_size == 0:
            database_songs: list[DatabaseSong] = []
            existing_years: tuple[int, ...] = ()
        else:
            with get_connection(path) as connection:
                tables = {
                    row["name"]
                    for row in connection.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                    ).fetchall()
                }
                if not REQUIRED_TABLES.issubset(tables):
                    missing = REQUIRED_TABLES - tables
                    raise DatabaseSchemaError(
                        "В существующей базе отсутствуют таблицы: " + ", ".join(sorted(missing))
                    )
                database_songs = _load_database_songs(connection)
                existing_years = _check_existing_ranking_years(
                    connection, workbook.ranking_years, replace_existing
                )
        if not workbook.ranking_years:
            return ImportResult(0, 0, 0, 0, dry_run=True)
        existing_by_identity, conflicts = _compare_with_database(workbook, database_songs)
        if conflicts:
            raise ImportConflictError(conflicts)
        unique_count = len(workbook.unique_songs)
        songs_added = sum(1 for identity in workbook.unique_songs if identity not in existing_by_identity)
        return ImportResult(
            rankings_imported=len(workbook.ranking_years),
            songs_added=songs_added,
            songs_reused=unique_count - songs_added,
            entries_added=len(workbook.songs),
            replaced_years=existing_years if replace_existing else (),
            dry_run=True,
        )

    _ensure_database_ready(path)

    with get_connection(path) as connection:
        connection.execute("BEGIN IMMEDIATE")
        try:
            existing_years = _check_existing_ranking_years(
                connection, workbook.ranking_years, replace_existing
            )
            if not workbook.ranking_years:
                connection.commit()
                return ImportResult(0, 0, 0, 0, dry_run=False)
            database_songs = _load_database_songs(connection)
            existing_by_identity, conflicts = _compare_with_database(workbook, database_songs)
            if conflicts:
                raise ImportConflictError(conflicts)

            artist_ids = _load_entity_ids(connection, "artists", "artist_id", "artist_name")
            language_ids = _load_entity_ids(connection, "languages", "language_id", "language_name")
            country_ids = _load_entity_ids(connection, "countries", "country_id", "country_name")

            if existing_years and replace_existing:
                placeholders = ",".join("?" for _ in existing_years)
                connection.execute(
                    f"DELETE FROM rankings WHERE ranking_year IN ({placeholders})",
                    existing_years,
                )

            # Одна запись песни на всю книгу, даже если она встречается в нескольких годах.
            song_ids: dict[SongIdentity, int] = {
                identity: db_song.song_id
                for identity, db_song in existing_by_identity.items()
            }
            songs_added = 0
            for identity, song in workbook.unique_songs.items():
                if identity not in song_ids:
                    song_ids[identity] = _insert_song(
                        connection,
                        song,
                        artist_ids,
                        language_ids,
                        country_ids,
                    )
                    songs_added += 1

            ranking_ids: dict[int, int] = {}
            for ranking_year in workbook.ranking_years:
                cursor = connection.execute(
                    "INSERT INTO rankings (ranking_year, top_size) VALUES (?, ?)",
                    (ranking_year, workbook.top_size),
                )
                ranking_ids[ranking_year] = int(cursor.lastrowid)

            for song in workbook.songs:
                connection.execute(
                    """
                    INSERT INTO ranking_entries (ranking_id, song_id, position)
                    VALUES (?, ?, ?)
                    """,
                    (ranking_ids[song.ranking_year], song_ids[song.identity], song.position),
                )

            connection.commit()
        except Exception:
            connection.rollback()
            raise

    return ImportResult(
        rankings_imported=len(workbook.ranking_years),
        songs_added=songs_added,
        songs_reused=len(workbook.unique_songs) - songs_added,
        entries_added=len(workbook.songs),
        replaced_years=existing_years if replace_existing else (),
        dry_run=False,
    )


def _print_result(result: ImportResult) -> None:
    mode = "Проверка без записи" if result.dry_run else "Импорт завершён"
    print(mode)
    print(f"Рейтингов: {result.rankings_imported}")
    print(f"Новых песен: {result.songs_added}")
    print(f"Повторно использовано песен: {result.songs_reused}")
    print(f"Позиций рейтинга: {result.entries_added}")
    if result.replaced_years:
        print("Заменены рейтинги за годы: " + ", ".join(map(str, result.replaced_years)))


def run_project_import(*, replace_existing: bool = False) -> ImportResult:
    """Импортирует книгу проекта; заменяет существующие годы только по явному запросу.

    Книга должна находиться в data/music_statistics.xlsx относительно корня
    репозитория. По умолчанию наличие хотя бы одного загруженного года — ошибка.
    """
    if not WORKBOOK_PATH.is_file():
        raise FileNotFoundError(
            f"Не найдена книга статистики: {WORKBOOK_PATH}. "
            "Поместите Excel-файл в каталог data и назовите его music_statistics.xlsx."
        )
    return import_excel(WORKBOOK_PATH, replace_existing=replace_existing)


def main() -> int:
    """Точка входа без аргументов командной строки."""
    try:
        result = run_project_import(replace_existing=True)
    except WorkbookValidationError as exc:
        print("Ошибки Excel:", file=sys.stderr)
        for error in exc.errors:
            print(f" - {error}", file=sys.stderr)
        return 2
    except ImportConflictError as exc:
        print("Импорт остановлен. База данных не изменена:", file=sys.stderr)
        for conflict in exc.conflicts:
            print(f" - {conflict}", file=sys.stderr)
        return 3
    except (ImporterError, OSError, sqlite3.Error) as exc:
        print(f"Ошибка импорта: {exc}", file=sys.stderr)
        return 4

    if result.rankings_imported == 0:
        print("Новых рейтингов нет. База данных уже актуальна.")
    else:
        _print_result(result)
    return 0
