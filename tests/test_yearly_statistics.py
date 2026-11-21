from unittest.mock import patch

from src.analytics import get_yearly_statistics


def test_new_repeated_and_returning_songs_are_classified_separately():
    rankings = {2023: 50, 2024: 50, 2025: 50, 2026: 50}
    # Song 1 is present continuously; song 3 returns in 2025 after missing 2024.
    # Song 4 debuts in 2024 and continues; song 5 debuts in 2025 and continues.
    rows = [
        {"ranking_year": 2023, "song_id": 1, "position": 1},
        {"ranking_year": 2023, "song_id": 2, "position": 2},
        {"ranking_year": 2023, "song_id": 3, "position": 3},
        {"ranking_year": 2024, "song_id": 1, "position": 1},
        {"ranking_year": 2024, "song_id": 2, "position": 2},
        {"ranking_year": 2024, "song_id": 4, "position": 3},
        {"ranking_year": 2025, "song_id": 1, "position": 1},
        {"ranking_year": 2025, "song_id": 3, "position": 2},
        {"ranking_year": 2025, "song_id": 4, "position": 3},
        {"ranking_year": 2025, "song_id": 5, "position": 4},
        {"ranking_year": 2026, "song_id": 1, "position": 1},
        {"ranking_year": 2026, "song_id": 3, "position": 2},
        {"ranking_year": 2026, "song_id": 5, "position": 3},
        {"ranking_year": 2026, "song_id": 6, "position": 4},
    ]

    with patch("src.analytics._read_ranking_data", return_value=(rankings, rows)):
        result = get_yearly_statistics()

    assert [
        (item.year, item.new_songs, item.repeated_songs, item.returning_songs)
        for item in result
    ] == [
        (2023, 3, 0, 0),
        (2024, 1, 2, 0),
        (2025, 1, 2, 1),
        (2026, 1, 3, 0),
    ]
    assert all(item.entries == item.new_songs + item.repeated_songs + item.returning_songs for item in result)
