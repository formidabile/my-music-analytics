from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

from src.analytics import (
    get_artist_stability_statistics,
    get_artist_statistics,
    get_composition_statistics,
    get_hall_of_fame,
    get_song_statistics,
    get_song_stability_statistics,
    get_yearly_statistics,
)
from src.database import DATABASE_PATH, get_connection
from src.forecasts import get_delayed_forecasts, get_forecast_analysis
from src.streaming import get_streaming_statistics


def _database_exists() -> bool:
    return DATABASE_PATH.is_file()


def _dict(value: Any) -> dict[str, Any]:
    if hasattr(value, "__dataclass_fields__"):
        return {name: getattr(value, name) for name in value.__dataclass_fields__}
    return dict(value)


def get_overview() -> dict[str, Any]:
    if not _database_exists():
        return {
            "database_ready": False,
            "years": [],
            "songs_count": 0,
            "artists_count": 0,
            "latest_year": None,
            "latest_top": [],
            "top_artists": [],
            "streaming_top": [],
            "forecast_metrics": [],
        }

    songs = get_song_statistics()
    artists = get_artist_statistics()
    years = get_yearly_statistics()
    latest_year = max((item.year for item in years), default=None)
    latest_top: list[dict[str, Any]] = []
    if latest_year is not None:
        with get_connection() as connection:
            rows = connection.execute(
                """SELECT re.position, s.song_id, s.title, s.release_year,
                          GROUP_CONCAT(a.artist_name, '; ') AS artists,
                          r.top_size - re.position + 1 AS points
                   FROM ranking_entries re
                   JOIN rankings r USING (ranking_id)
                   JOIN songs s USING (song_id)
                   LEFT JOIN song_artists sa USING (song_id)
                   LEFT JOIN artists a USING (artist_id)
                   WHERE r.ranking_year = ?
                   GROUP BY re.ranking_id, re.song_id, re.position, s.song_id
                   ORDER BY re.position LIMIT 5""",
                (latest_year,),
            ).fetchall()
        latest_top = [dict(row) for row in rows]

    streaming = get_streaming_statistics()
    streaming_years = [row["year"] for row in streaming]
    latest_streaming_year = max(streaming_years, default=None)
    streaming_top = [row for row in streaming if row["year"] == latest_streaming_year]
    forecast_metrics, _, _ = get_forecast_analysis()
    return {
        "database_ready": True,
        "years": [item.year for item in years],
        "songs_count": len(songs),
        "artists_count": len(artists),
        "latest_year": latest_year,
        "latest_top": latest_top,
        "top_artists": [_dict(item) for item in artists[:5]],
        "streaming_top": streaming_top,
        "forecast_metrics": [_dict(item) for item in forecast_metrics],
        "yearly": [_dict(item) for item in years],
    }


def get_years() -> list[int]:
    if not _database_exists():
        return []
    with get_connection() as connection:
        rows = connection.execute(
            "SELECT ranking_year FROM rankings ORDER BY ranking_year"
        ).fetchall()
    return [int(row["ranking_year"]) for row in rows]


def get_ranking(year: int) -> list[dict[str, Any]]:
    if not _database_exists():
        return []
    with get_connection() as connection:
        rows = connection.execute(
            """SELECT re.position, s.song_id, s.title, s.release_year,
                      GROUP_CONCAT(a.artist_name, '; ') AS artists,
                      r.top_size - re.position + 1 AS points
               FROM ranking_entries re
               JOIN rankings r USING (ranking_id)
               JOIN songs s USING (song_id)
               LEFT JOIN song_artists sa USING (song_id)
               LEFT JOIN artists a USING (artist_id)
               WHERE r.ranking_year = ?
               GROUP BY re.ranking_id, re.song_id, re.position, s.song_id
               ORDER BY re.position""",
            (year,),
        ).fetchall()
    return [dict(row) for row in rows]


def get_ranking_composition(year: int) -> dict[str, Any]:
    """Возвращает возрастной, страновой и языковой состав одного рейтинга."""
    if not _database_exists():
        return {"entries_count": 0, "age": [], "countries": [], "languages": []}

    with get_connection(DATABASE_PATH) as connection:
        entries = connection.execute(
            """SELECT s.song_id, s.release_year
               FROM ranking_entries re
               JOIN rankings r USING (ranking_id)
               JOIN songs s USING (song_id)
               WHERE r.ranking_year = ?
               ORDER BY re.position""",
            (year,),
        ).fetchall()
        country_rows = connection.execute(
            """SELECT sc.song_id, c.country_name AS category
               FROM ranking_entries re
               JOIN rankings r USING (ranking_id)
               JOIN song_countries sc USING (song_id)
               JOIN countries c USING (country_id)
               WHERE r.ranking_year = ?""",
            (year,),
        ).fetchall()
        language_rows = connection.execute(
            """SELECT sl.song_id, l.language_name AS category
               FROM ranking_entries re
               JOIN rankings r USING (ranking_id)
               JOIN song_languages sl USING (song_id)
               JOIN languages l USING (language_id)
               WHERE r.ranking_year = ?""",
            (year,),
        ).fetchall()

    total = len(entries)
    age_labels = [
        "Актуальные хиты", "Устоявшиеся хиты", "Старые хиты",
        "Легендарные хиты", "Классика", "Не указано",
    ]
    age_counts = {label: 0 for label in age_labels}
    for entry in entries:
        release_year = entry["release_year"]
        if release_year is None:
            category = "Не указано"
        else:
            age = year - int(release_year)
            if age < 0:
                category = "Не указано"
            elif age <= 3:
                category = "Актуальные хиты"
            elif age <= 10:
                category = "Устоявшиеся хиты"
            elif age <= 20:
                category = "Старые хиты"
            elif age <= 30:
                category = "Легендарные хиты"
            else:
                category = "Классика"
        age_counts[category] += 1

    def dimension_stats(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
        members: dict[str, set[int]] = {}
        for row in rows:
            members.setdefault(str(row["category"]), set()).add(int(row["song_id"]))
        known_song_ids = {int(entry["song_id"]) for entry in entries}
        assigned = set().union(*members.values()) if members else set()
        if known_song_ids - assigned:
            members["Не указано"] = known_song_ids - assigned
        return [
            {"category": category, "songs_count": len(song_ids),
             "share": len(song_ids) / total if total else 0.0}
            for category, song_ids in sorted(
                members.items(), key=lambda item: (-len(item[1]), item[0].casefold())
            )
        ]

    return {
        "entries_count": total,
        "age": [
            {"category": category, "songs_count": count,
             "share": count / total if total else 0.0}
            for category, count in age_counts.items() if count
        ],
        "countries": dimension_stats(country_rows),
        "languages": dimension_stats(language_rows),
    }


def get_song_detail(song_id: int) -> dict[str, Any] | None:
    songs = get_song_statistics()
    item = next((song for song in songs if song.song_id == song_id), None)
    if item is None:
        return None
    result = _dict(item)
    with get_connection() as connection:
        countries = connection.execute(
            """SELECT c.country_name FROM song_countries sc
               JOIN countries c USING (country_id) WHERE sc.song_id = ?
               ORDER BY c.country_name""", (song_id,),
        ).fetchall()
        languages = connection.execute(
            """SELECT l.language_name FROM song_languages sl
               JOIN languages l USING (language_id) WHERE sl.song_id = ?
               ORDER BY l.language_name""", (song_id,),
        ).fetchall()
    result["countries"] = [row["country_name"] for row in countries]
    result["languages"] = [row["language_name"] for row in languages]
    result["top10_appearances"] = sum(
        position is not None and position <= 10
        for position in result["positions_by_year"].values()
    )
    result["history"] = [
        {
            "year": int(year),
            "position": result["positions_by_year"][year],
            "points": result["points_by_year"][year],
        }
        for year in sorted(result["positions_by_year"])
    ]
    return result


def get_artist_detail(artist_id: int) -> dict[str, Any] | None:
    artists = get_artist_statistics()
    item = next((artist for artist in artists if artist.artist_id == artist_id), None)
    if item is None:
        return None
    result = _dict(item)
    result["songs"] = sorted(
        [_dict(song) for song in get_song_statistics() if item.artist in song.artists],
        key=lambda song: (-song["appearances"], -song["total_points"], song["title"].casefold()),
    )
    result["history"] = [
        {
            "year": int(year),
            "points": result["points_by_year"][year],
            "appearances": result["appearances_by_year"][year],
            "best_position": result["best_positions_by_year"][year],
        }
        for year in sorted(result["points_by_year"])
    ]
    return result


def get_analytics_payload() -> dict[str, Any]:
    return {
        "songs": [_dict(item) for item in get_song_statistics()],
        "artists": [_dict(item) for item in get_artist_statistics()],
        "top_artists_yearly": [
            {
                "artist_id": item.artist_id,
                "artist": item.artist,
                "total_points": item.total_points,
                "points_by_year": item.points_by_year,
            }
            for item in get_artist_statistics()[:6]
        ],
        "hall_of_fame": sorted(
            [_dict(item) for item in get_hall_of_fame()],
            key=lambda item: (-item["appearances"], -item["total_points"], item["title"].casefold()),
        ),
        "stability": [_dict(item) for item in get_artist_stability_statistics()],
        "song_stability": [_dict(item) for item in get_song_stability_statistics()],
        "composition": [_dict(item) for item in get_composition_statistics()],
        "yearly": [_dict(item) for item in get_yearly_statistics()],
    }



def get_streaming_payload() -> dict[str, Any]:
    rows = get_streaming_statistics()
    years = sorted({int(row["year"]) for row in rows})
    all_time: dict[str, dict[str, Any]] = {}
    for row in rows:
        artist = str(row["artist"])
        item = all_time.setdefault(artist, {
            "artist": artist, "years_count": 0, "total_minutes": 0,
            "known_minutes_years": 0, "last_year": None, "last_position": None,
        })
        item["years_count"] += 1
        item["last_year"] = row["year"]
        item["last_position"] = row["position"]
        if row["listening_minutes"] is not None:
            item["total_minutes"] += int(row["listening_minutes"])
            item["known_minutes_years"] += 1
    all_time_rows = sorted(
        all_time.values(),
        key=lambda x: (-(x["total_minutes"] if x["known_minutes_years"] else -1), -x["years_count"], x["artist"].casefold()),
    )
    return {"years": years, "entries": rows, "all_time": all_time_rows}


def get_forecasts_payload() -> dict[str, Any]:
    metrics, comparisons, unexpected = get_forecast_analysis()
    delayed = get_delayed_forecasts()
    return {
        "metrics": [_dict(item) for item in metrics],
        "comparisons": [_dict(item) for item in comparisons],
        "unexpected": [_dict(item) for item in unexpected],
        "delayed": [_dict(item) for item in delayed],
    }
