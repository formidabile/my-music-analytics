from src.database import initialize_database, get_connection
from web.api import data


def _prepare_rankings_database(path):
    initialize_database(path)
    songs = [
        (1, "Актуальный хит", 2023),       # age 3 in 2026, age 4 in 2027
        (2, "Устоявшийся хит", 2022),      # age 4
        (3, "Устоявшийся хит 2", 2016),    # age 10
        (4, "Старый хит", 2015),           # age 11
        (5, "Старый хит 2", 2006),         # age 20
        (6, "Легендарный хит", 2005),      # age 21
        (7, "Легендарный хит 2", 1996),    # age 30
        (8, "Классика", 1995),             # age 31
    ]
    with get_connection(path) as connection:
        connection.executemany(
            "INSERT INTO songs(song_id, title, release_year) VALUES (?, ?, ?)", songs
        )
        connection.executemany(
            "INSERT INTO rankings(ranking_id, ranking_year, top_size) VALUES (?, ?, ?)",
            [(1, 2026, 50), (2, 2027, 50)],
        )
        connection.executemany(
            "INSERT INTO ranking_entries(ranking_id, song_id, position) VALUES (?, ?, ?)",
            [(ranking_id, song_id, song_id) for ranking_id in (1, 2) for song_id in range(1, 9)],
        )
        connection.execute("INSERT INTO countries(country_id, country_name) VALUES (1, 'Россия')")
        connection.execute("INSERT INTO countries(country_id, country_name) VALUES (2, 'США')")
        connection.execute("INSERT INTO languages(language_id, language_name) VALUES (1, 'Русский')")
        connection.execute("INSERT INTO languages(language_id, language_name) VALUES (2, 'Английский')")
        connection.executemany(
            "INSERT INTO song_countries(song_id, country_id) VALUES (?, ?)",
            [(1, 1), (1, 2), (2, 1)],
        )
        connection.executemany(
            "INSERT INTO song_languages(song_id, language_id) VALUES (?, ?)",
            [(1, 1), (2, 2)],
        )
        connection.commit()


def test_ranking_age_categories_use_selected_ranking_year(tmp_path, monkeypatch):
    db_path = tmp_path / "music.db"
    _prepare_rankings_database(db_path)
    monkeypatch.setattr(data, "DATABASE_PATH", db_path)

    result_2026 = data.get_ranking_composition(2026)
    age_2026 = {row["category"]: row["songs_count"] for row in result_2026["age"]}
    assert age_2026 == {
        "Актуальные хиты": 1,
        "Устоявшиеся хиты": 2,
        "Старые хиты": 2,
        "Легендарные хиты": 2,
        "Классика": 1,
    }

    result_2027 = data.get_ranking_composition(2027)
    age_2027 = {row["category"]: row["songs_count"] for row in result_2027["age"]}
    assert age_2027.get("Актуальные хиты", 0) == 0
    assert age_2027["Устоявшиеся хиты"] == 2


def test_country_and_language_shares_count_songs_and_allow_multiple_values(tmp_path, monkeypatch):
    db_path = tmp_path / "music.db"
    _prepare_rankings_database(db_path)
    monkeypatch.setattr(data, "DATABASE_PATH", db_path)

    result = data.get_ranking_composition(2026)
    countries = {row["category"]: row for row in result["countries"]}
    languages = {row["category"]: row for row in result["languages"]}

    assert countries["Россия"]["songs_count"] == 2
    assert countries["США"]["songs_count"] == 1
    assert countries["Не указано"]["songs_count"] == 6
    assert languages["Русский"]["songs_count"] == 1
    assert languages["Английский"]["songs_count"] == 1
    assert languages["Не указано"]["songs_count"] == 6
    assert sum(row["share"] for row in countries.values()) > 1
