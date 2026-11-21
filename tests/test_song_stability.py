from unittest.mock import patch

from src.analytics import SongStats, get_song_stability_statistics


def _song(song_id, positions):
    return SongStats(
        song_id=song_id,
        title=f"Song {song_id}",
        release_year=2020,
        artists=("Artist",),
        appearances=sum(position is not None for position in positions.values()),
        best_position=min(position for position in positions.values() if position is not None),
        average_position=1.0,
        average_position_with_absences=1.0,
        total_points=100,
        positions_by_year=positions,
        points_by_year={year: 0 for year in positions},
    )


def test_song_stability_calculates_presence_rate_and_streaks():
    songs = [
        _song(1, {2023: 5, 2024: 7, 2025: 9}),
        _song(2, {2023: 3, 2024: None, 2025: 4}),
        _song(3, {2023: None, 2024: 2, 2025: 6}),
    ]
    with patch("src.analytics.get_song_statistics", return_value=songs):
        result = get_song_stability_statistics()

    by_id = {item.song_id: item for item in result}
    assert by_id[1].years_present == 3
    assert by_id[1].presence_rate == 1.0
    assert by_id[1].longest_streak == 3
    assert by_id[1].current_streak == 3
    assert by_id[2].longest_streak == 1
    assert by_id[2].current_streak == 1
    assert by_id[3].longest_streak == 2
    assert by_id[3].current_streak == 2
    assert [item.song_id for item in result] == [1, 3, 2]


def test_song_stability_returns_empty_when_no_songs():
    with patch("src.analytics.get_song_statistics", return_value=[]):
        assert get_song_stability_statistics() == []
