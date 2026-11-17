"""Аналитика годовых музыкальных рейтингов, хранящихся в SQLite.

Модуль не создаёт дополнительные таблицы: все показатели рассчитываются
по исходным таблицам рейтингов и связям между песнями и исполнителями.
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from statistics import mean
from math import ceil, floor
from typing import TypeAlias

from .database import DATABASE_PATH, REQUIRED_TABLES, get_connection

Year: TypeAlias = int


@dataclass(frozen=True)
class SongStats:
    song_id: int
    title: str
    release_year: int
    artists: tuple[str, ...]
    appearances: int
    best_position: int
    average_position: float
    average_position_with_absences: float
    total_points: int
    positions_by_year: dict[Year, int | None]
    points_by_year: dict[Year, int]


@dataclass(frozen=True)
class ArtistStats:
    artist_id: int
    artist: str
    appearances: int
    best_position: int
    average_position: float
    total_points: int
    appearances_by_year: dict[Year, int]
    best_positions_by_year: dict[Year, int | None]
    average_positions_by_year: dict[Year, float | None]
    points_by_year: dict[Year, int]


@dataclass(frozen=True)
class YearStats:
    year: int
    top_size: int
    entries: int
    new_songs: int
    returning_songs: int
    average_position: float


@dataclass(frozen=True)
class HallOfFameEntry:
    song_id: int
    title: str
    release_year: int
    artists: tuple[str, ...]
    appearances: int
    induction_year: int
    best_position: int
    average_position: float
    total_points: int


@dataclass(frozen=True)
class ArtistStabilityStats:
    """Показатели регулярности присутствия и концентрации баллов исполнителя."""

    artist: str
    years_present: int
    total_ranking_years: int
    presence_rate: float
    longest_streak: int
    current_streak: int
    max_tracks_in_year: int
    average_tracks_per_present_year: float
    top_track_title: str
    top_track_points: int
    top_track_points_share: float


def _read_ranking_data(
    db_path: str | Path | None = None,
) -> tuple[dict[Year, int], list[sqlite3.Row]]:
    path = Path(db_path) if db_path is not None else DATABASE_PATH
    if not path.is_file():
        raise FileNotFoundError(
            f"База данных не найдена: {path}. Сначала запустите импорт рейтингов."
        )

    with get_connection(path) as connection:
        tables = {
            row["name"]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        missing = REQUIRED_TABLES - tables
        if missing:
            raise sqlite3.DatabaseError(
                "В базе отсутствуют обязательные таблицы: " + ", ".join(sorted(missing))
            )

        rankings = {
            row["ranking_year"]: row["top_size"]
            for row in connection.execute(
                "SELECT ranking_year, top_size FROM rankings ORDER BY ranking_year"
            ).fetchall()
        }
        entries = connection.execute(
            """
            SELECT
                s.song_id,
                s.title,
                s.release_year,
                r.ranking_year,
                r.top_size,
                re.position,
                a.artist_id,
                a.artist_name
            FROM ranking_entries AS re
            JOIN rankings AS r ON r.ranking_id = re.ranking_id
            JOIN songs AS s ON s.song_id = re.song_id
            LEFT JOIN song_artists AS sa ON sa.song_id = s.song_id
            LEFT JOIN artists AS a ON a.artist_id = sa.artist_id
            ORDER BY s.song_id, r.ranking_year, re.position, a.artist_name
            """
        ).fetchall()

    return rankings, entries


def get_song_statistics(db_path: str | Path | None = None) -> list[SongStats]:
    """Возвращает сводную статистику песен, отсортированную по набранным баллам.

    Баллы за место = top_size + 1 - position. Отсутствие в годовом рейтинге
    даёт 0 баллов. Для среднего места с учётом отсутствий используется штраф
    top_size + 1 соответствующего года.
    """
    rankings, rows = _read_ranking_data(db_path)
    if not rankings:
        return []

    songs: dict[int, dict] = {}
    for row in rows:
        song = songs.setdefault(
            row["song_id"],
            {
                "title": row["title"],
                "release_year": row["release_year"],
                "artists": set(),
                "positions": {},
            },
        )
        song["positions"][row["ranking_year"]] = row["position"]
        if row["artist_name"] is not None:
            song["artists"].add(row["artist_name"])

    result: list[SongStats] = []
    for song_id, song in songs.items():
        positions = song["positions"]
        appearances = len(positions)
        best_position = min(positions.values())
        average_position = mean(positions.values())
        points_by_year = {
            year: (top_size + 1 - positions[year]) if year in positions else 0
            for year, top_size in rankings.items()
        }
        average_with_absences = mean(
            positions.get(year, top_size + 1)
            for year, top_size in rankings.items()
        )
        result.append(
            SongStats(
                song_id=song_id,
                title=song["title"],
                release_year=song["release_year"],
                artists=tuple(sorted(song["artists"], key=str.casefold)),
                appearances=appearances,
                best_position=best_position,
                average_position=average_position,
                average_position_with_absences=average_with_absences,
                total_points=sum(points_by_year.values()),
                positions_by_year={year: positions.get(year) for year in rankings},
                points_by_year=points_by_year,
            )
        )

    return sorted(
        result,
        key=lambda item: (
            -item.total_points,
            -item.appearances,
            item.best_position,
            item.title.casefold(),
        ),
    )


def get_artist_statistics(db_path: str | Path | None = None) -> list[ArtistStats]:
    """Считает статистику исполнителей; совместные песни засчитываются каждому.

    Показатель appearances — число попаданий песен исполнителя в рейтинги.
    Если у исполнителя несколько песен в одном году, каждая учитывается отдельно.
    """
    rankings, rows = _read_ranking_data(db_path)
    if not rankings:
        return []

    artists: dict[int, dict] = {}
    for row in rows:
        artist_id = row["artist_id"]
        if artist_id is None:
            continue
        artist = artists.setdefault(
            artist_id,
            {"name": row["artist_name"], "positions": []},
        )
        artist["positions"].append((row["ranking_year"], row["position"], row["top_size"]))

    result: list[ArtistStats] = []
    for artist_id, artist in artists.items():
        positions = artist["positions"]
        appearances_by_year = {year: 0 for year in rankings}
        best_positions_by_year: dict[Year, int | None] = {year: None for year in rankings}
        positions_by_year: dict[Year, list[int]] = {year: [] for year in rankings}
        points_by_year = {year: 0 for year in rankings}

        for year, position, top_size in positions:
            appearances_by_year[year] += 1
            positions_by_year[year].append(position)
            current_best = best_positions_by_year[year]
            if current_best is None or position < current_best:
                best_positions_by_year[year] = position
            points_by_year[year] += top_size + 1 - position

        all_positions = [position for _, position, _ in positions]
        result.append(
            ArtistStats(
                artist_id=artist_id,
                artist=artist["name"],
                appearances=len(all_positions),
                best_position=min(all_positions),
                average_position=mean(all_positions),
                total_points=sum(points_by_year.values()),
                appearances_by_year=appearances_by_year,
                best_positions_by_year=best_positions_by_year,
                average_positions_by_year={
                    year: mean(values) if values else None
                    for year, values in positions_by_year.items()
                },
                points_by_year=points_by_year,
            )
        )

    return sorted(
        result,
        key=lambda item: (
            -item.total_points,
            -item.appearances,
            item.best_position,
            item.artist.casefold(),
        ),
    )


def get_artist_stability_statistics(
    db_path: str | Path | None = None,
) -> list[ArtistStabilityStats]:
    """Оценивает устойчивость популярности исполнителей.

    Последовательными считаются соседние рейтинги в имеющемся наборе данных
    (например, 2023 и 2025, если рейтинга за 2024 год нет). Доля присутствия
    показывает, в скольких рейтингах из общего числа появлялся исполнитель.
    Доля лучшего трека — часть баллов исполнителя, набранная его самым
    результативным отдельным треком. Совместные треки засчитываются каждому.
    """
    rankings, _ = _read_ranking_data(db_path)
    if not rankings:
        return []

    artist_stats = get_artist_statistics(db_path)
    songs = get_song_statistics(db_path)
    total_years = len(rankings)
    result: list[ArtistStabilityStats] = []

    for artist in artist_stats:
        active_flags = [artist.appearances_by_year.get(year, 0) > 0 for year in rankings]
        years_present = sum(active_flags)
        longest_streak = current_streak = streak = 0
        for active in active_flags:
            if active:
                streak += 1
                longest_streak = max(longest_streak, streak)
            else:
                streak = 0
        for active in reversed(active_flags):
            if not active:
                break
            current_streak += 1

        artist_songs = [song for song in songs if artist.artist in song.artists]
        best_song = max(artist_songs, key=lambda song: (song.total_points, song.title.casefold()))
        max_tracks = max(artist.appearances_by_year.values(), default=0)
        result.append(
            ArtistStabilityStats(
                artist=artist.artist,
                years_present=years_present,
                total_ranking_years=total_years,
                presence_rate=years_present / total_years if total_years else 0.0,
                longest_streak=longest_streak,
                current_streak=current_streak,
                max_tracks_in_year=max_tracks,
                average_tracks_per_present_year=(artist.appearances / years_present if years_present else 0.0),
                top_track_title=best_song.title,
                top_track_points=best_song.total_points,
                top_track_points_share=(best_song.total_points / artist.total_points if artist.total_points else 0.0),
            )
        )

    return sorted(
        result,
        key=lambda item: (
            -item.presence_rate, -item.longest_streak, -item.current_streak,
            -item.top_track_points, item.artist.casefold(),
        ),
    )


def get_hall_of_fame(
    threshold: int = 3,
    db_path: str | Path | None = None,
) -> list[HallOfFameEntry]:
    """Возвращает песни, попавшие в годовой Top-N не менее threshold раз.

    Годом включения считается год N-го появления песни в рейтинге.
    """
    if threshold < 1:
        raise ValueError("Порог Hall of Fame должен быть не меньше 1")

    songs = get_song_statistics(db_path)
    hall_of_fame: list[HallOfFameEntry] = []
    for song in songs:
        if song.appearances < threshold:
            continue
        years_present = [
            year for year, position in song.positions_by_year.items() if position is not None
        ]
        induction_year = sorted(years_present)[threshold - 1]
        hall_of_fame.append(
            HallOfFameEntry(
                song_id=song.song_id,
                title=song.title,
                release_year=song.release_year,
                artists=song.artists,
                appearances=song.appearances,
                induction_year=induction_year,
                best_position=song.best_position,
                average_position=song.average_position,
                total_points=song.total_points,
            )
        )

    return sorted(
        hall_of_fame,
        key=lambda item: (item.induction_year, -item.total_points, item.title.casefold()),
    )


def get_yearly_statistics(db_path: str | Path | None = None) -> list[YearStats]:
    """Возвращает сводные показатели по каждому годовому рейтингу.

    Новая песня — та, которая впервые появилась в рейтингах именно в этом году.
    Повторная — уже встречавшаяся хотя бы в одном более раннем рейтинге.
    Средняя позиция рассчитывается по всем строкам рейтинга за год.
    """
    rankings, rows = _read_ranking_data(db_path)
    if not rankings:
        return []

    # В выборке одна песня может повторяться из-за нескольких исполнителей.
    # Для годовой статистики учитываем каждую запись рейтинга ровно один раз.
    positions_by_year: dict[int, dict[int, int]] = {year: {} for year in rankings}
    first_appearance: dict[int, int] = {}
    for row in rows:
        song_id = row["song_id"]
        year = row["ranking_year"]
        positions_by_year[year][song_id] = row["position"]
        first_appearance[song_id] = min(first_appearance.get(song_id, year), year)

    result: list[YearStats] = []
    for year, top_size in rankings.items():
        positions = positions_by_year[year]
        new_songs = sum(1 for song_id in positions if first_appearance[song_id] == year)
        returning_songs = len(positions) - new_songs
        result.append(
            YearStats(
                year=year,
                top_size=top_size,
                entries=len(positions),
                new_songs=new_songs,
                returning_songs=returning_songs,
                average_position=mean(positions.values()) if positions else 0.0,
            )
        )
    return result


@dataclass(frozen=True)
class CompositionStats:
    """Распределение уникальных песен по одной категории состава рейтинга."""

    year: int | None  # None означает всю историю рейтингов
    dimension: str
    category: str
    songs_count: int
    share: float


def get_composition_statistics(
    db_path: str | Path | None = None,
    interval_count: int = 7,
) -> list[CompositionStats]:
    """Распределения песен по странам, языкам и году выпуска.

    Для сравнимости интервалы года выпуска рассчитываются один раз по всему
    набору песен и применяются ко всем годовым рейтингам. Внутри категории
    страна/язык песня считается один раз; при нескольких значениях она
    попадает в каждую соответствующую категорию.
    """
    if interval_count < 1:
        raise ValueError("Количество интервалов должно быть не меньше 1")

    rankings, _ = _read_ranking_data(db_path)
    if not rankings:
        return []

    path = Path(db_path) if db_path is not None else DATABASE_PATH
    with get_connection(path) as connection:
        song_rows = connection.execute(
            "SELECT song_id, release_year FROM songs ORDER BY song_id"
        ).fetchall()
        country_rows = connection.execute(
            """SELECT sc.song_id, c.country_name
               FROM song_countries sc JOIN countries c USING(country_id)"""
        ).fetchall()
        language_rows = connection.execute(
            """SELECT sl.song_id, l.language_name
               FROM song_languages sl JOIN languages l USING(language_id)"""
        ).fetchall()
        year_rows = connection.execute(
            """SELECT r.ranking_year, re.song_id
               FROM ranking_entries re JOIN rankings r USING(ranking_id)
               ORDER BY r.ranking_year"""
        ).fetchall()

    release_years = {row["song_id"]: row["release_year"] for row in song_rows}
    if not release_years:
        return []
    min_year = min(release_years.values())
    max_year = max(release_years.values())
    year_span = max_year - min_year + 1
    interval_width = year_span / interval_count

    countries: dict[int, set[str]] = {}
    for row in country_rows:
        countries.setdefault(row["song_id"], set()).add(row["country_name"])
    languages: dict[int, set[str]] = {}
    for row in language_rows:
        languages.setdefault(row["song_id"], set()).add(row["language_name"])

    songs_by_year: dict[int, set[int]] = {year: set() for year in rankings}
    for row in year_rows:
        songs_by_year[row["ranking_year"]].add(row["song_id"])
    populations: list[tuple[int | None, set[int]]] = [
        (year, song_ids) for year, song_ids in songs_by_year.items()
    ]
    populations.append((None, set().union(*songs_by_year.values())))

    result: list[CompositionStats] = []
    dimension_names = {
        "country": "Страна",
        "language": "Язык",
        "decade": "Десятилетие",
        "interval": "Интервал года выпуска",
    }

    for year, song_ids in populations:
        denominator = len(song_ids)
        if not denominator:
            continue
        buckets: dict[str, dict[str, set[int]]] = {key: {} for key in dimension_names}
        for song_id in song_ids:
            for country in countries.get(song_id, {"Не указано"}):
                buckets["country"].setdefault(country, set()).add(song_id)
            for language in languages.get(song_id, {"Не указано"}):
                buckets["language"].setdefault(language, set()).add(song_id)
            release_year = release_years[song_id]
            decade_start = (release_year // 10) * 10
            decade_label = f"{decade_start}-е"
            buckets["decade"].setdefault(decade_label, set()).add(song_id)
            interval_index = min(
                interval_count - 1, floor((release_year - min_year) / interval_width)
            )
            start = min_year + ceil(interval_index * interval_width)
            end = min(max_year, min_year + ceil((interval_index + 1) * interval_width) - 1)
            interval_label = f"{start}–{end}"
            buckets["interval"].setdefault(interval_label, set()).add(song_id)

        for dimension, categories in buckets.items():
            for category, members in sorted(categories.items(), key=lambda item: item[0].casefold()):
                count = len(members)
                result.append(
                    CompositionStats(
                        year=year,
                        dimension=dimension_names[dimension],
                        category=category,
                        songs_count=count,
                        share=count / denominator,
                    )
                )

    return result
