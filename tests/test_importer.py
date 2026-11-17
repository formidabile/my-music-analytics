from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from openpyxl import Workbook, load_workbook

import src.database as database_module
import src.importer as importer_module
from src.database import get_connection, initialize_database
from src.analytics import get_artist_statistics, get_hall_of_fame, get_song_statistics
from src.importer import (
    ImportConflictError,
    WorkbookValidationError,
    import_excel,
    read_workbook,
    run_project_import,
)

SOURCE_WORKBOOK = Path(
    os.environ.get(
        "MUSIC_STATS_TEST_WORKBOOK",
        Path(__file__).resolve().parents[1] / "data" / "music_statistics.xlsx",
    )
)
HEADERS = ["Исполнители", "Песня", "Язык", "Страна", "Год"]


class ImporterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db_path = self.root / "music.db"
        self.workbook_path = self.root / "synthetic_music_statistics.xlsx"
        self._create_synthetic_workbook(self.workbook_path)

        # Даже если новый тест случайно забудет передать db_path явно,
        # он не сможет обратиться к постоянной базе проекта.
        self._database_path_patch = patch.object(database_module, "DATABASE_PATH", self.db_path)
        self._importer_database_path_patch = patch.object(importer_module, "DATABASE_PATH", self.db_path)
        self._workbook_path_patch = patch.object(importer_module, "WORKBOOK_PATH", self.workbook_path)
        self._database_path_patch.start()
        self._importer_database_path_patch.start()
        self._workbook_path_patch.start()

    def tearDown(self) -> None:
        self._workbook_path_patch.stop()
        self._importer_database_path_patch.stop()
        self._database_path_patch.stop()
        temporary_database = self.db_path
        self.temp_dir.cleanup()
        self.assertFalse(temporary_database.exists(), "Временная база теста должна удаляться")

    @staticmethod
    def _create_synthetic_workbook(path: Path) -> None:
        workbook = Workbook()
        first = workbook.active
        first.title = "2023"
        first.append(HEADERS)
        first.append(["Artist A; Artist B", "Shared track", "English; French", "USA; Canada", 2001])
        for number in range(2, 51):
            first.append([f"Artist {number}", f"Song {number}", "English", "USA", 1990 + number])

        second = workbook.create_sheet("2024")
        second.append(HEADERS)
        # Тот же набор исполнителей в обратном порядке: это должна быть одна песня.
        second.append(["Artist B; Artist A", "Shared track", "French; English", "Canada; USA", 2001])
        for number in range(2, 51):
            second.append(
                [f"Other Artist {number}", f"Other song {number}", "French", "France", 2000 + number]
            )
        workbook.save(path)
        workbook.close()

    def _counts(self) -> dict[str, int]:
        with get_connection(self.db_path) as connection:
            return {
                "songs": connection.execute("SELECT COUNT(*) FROM songs").fetchone()[0],
                "rankings": connection.execute("SELECT COUNT(*) FROM rankings").fetchone()[0],
                "entries": connection.execute("SELECT COUNT(*) FROM ranking_entries").fetchone()[0],
                "artists": connection.execute("SELECT COUNT(*) FROM artists").fetchone()[0],
                "countries": connection.execute("SELECT COUNT(*) FROM countries").fetchone()[0],
                "languages": connection.execute("SELECT COUNT(*) FROM languages").fetchone()[0],
            }

    def _snapshot(self) -> dict[str, list[tuple]]:
        """Снимок всех таблиц для проверки отсутствия частичных изменений."""
        table_order = {
            "countries": "country_id",
            "languages": "language_id",
            "artists": "artist_id",
            "songs": "song_id",
            "song_artists": "song_id, artist_id",
            "song_countries": "song_id, country_id",
            "song_languages": "song_id, language_id",
            "rankings": "ranking_id",
            "ranking_entries": "ranking_id, position",
        }
        with get_connection(self.db_path) as connection:
            return {
                table: [
                    tuple(row)
                    for row in connection.execute(
                        f"SELECT * FROM {table} ORDER BY {order_by}"
                    ).fetchall()
                ]
                for table, order_by in table_order.items()
            }

    def _mutate_workbook(self, sheet: str, cell: str, value: object, destination: Path) -> Path:
        workbook = load_workbook(self.workbook_path)
        workbook[sheet][cell] = value
        workbook.save(destination)
        workbook.close()
        return destination

    def test_synthetic_workbook_validates_and_artist_order_is_ignored(self) -> None:
        data = read_workbook(self.workbook_path)
        self.assertEqual(data.ranking_years, [2023, 2024])
        self.assertEqual(len(data.songs), 100)
        self.assertEqual(len(data.unique_songs), 99)

    def test_import_reuses_song_across_years_and_keeps_positions(self) -> None:
        result = import_excel(self.workbook_path, self.db_path)
        self.assertEqual(result.rankings_imported, 2)
        self.assertEqual(result.songs_added, 99)
        self.assertEqual(result.entries_added, 100)
        counts = self._counts()
        self.assertEqual(counts["songs"], 99)
        self.assertEqual(counts["rankings"], 2)
        self.assertEqual(counts["entries"], 100)

        with get_connection(self.db_path) as connection:
            shared_song_id = connection.execute(
                "SELECT song_id FROM songs WHERE title = 'Shared track'"
            ).fetchone()[0]
            positions = connection.execute(
                """
                SELECT r.ranking_year, re.position
                FROM ranking_entries re
                JOIN rankings r ON r.ranking_id = re.ranking_id
                WHERE re.song_id = ?
                ORDER BY r.ranking_year
                """,
                (shared_song_id,),
            ).fetchall()
            artist_links = connection.execute(
                "SELECT COUNT(*) FROM song_artists WHERE song_id = ?", (shared_song_id,)
            ).fetchone()[0]
            language_links = connection.execute(
                "SELECT COUNT(*) FROM song_languages WHERE song_id = ?", (shared_song_id,)
            ).fetchone()[0]
            country_links = connection.execute(
                "SELECT COUNT(*) FROM song_countries WHERE song_id = ?", (shared_song_id,)
            ).fetchone()[0]
        self.assertEqual([(row["ranking_year"], row["position"]) for row in positions], [(2023, 1), (2024, 1)])
        self.assertEqual((artist_links, language_links, country_links), (2, 2, 2))

    def test_repeat_import_is_blocked_and_explicit_replacement_is_safe(self) -> None:
        import_excel(self.workbook_path, self.db_path)
        before = self._counts()
        with self.assertRaises(ImportConflictError):
            import_excel(self.workbook_path, self.db_path)
        self.assertEqual(self._counts(), before)

        result = import_excel(self.workbook_path, self.db_path, replace_existing=True)
        self.assertEqual(result.songs_added, 0)
        self.assertEqual(result.songs_reused, 99)
        self.assertEqual(result.replaced_years, (2023, 2024))
        self.assertEqual(self._counts(), before)

    def test_any_existing_year_blocks_default_import_without_changes(self) -> None:
        first_year_book = self.root / "only_2023.xlsx"
        workbook = load_workbook(self.workbook_path)
        del workbook["2024"]
        workbook.save(first_year_book)
        workbook.close()

        import_excel(first_year_book, self.db_path)
        before = self._snapshot()

        # Даже если в книге есть новый 2024 год, наличие 2023 блокирует
        # весь импорт: никакого частичного или молчаливого пропуска.
        with self.assertRaises(ImportConflictError):
            import_excel(self.workbook_path, self.db_path)
        self.assertEqual(self._snapshot(), before)

        # Отдельная книга только с новым годом импортируется штатно.
        new_year_book = self.root / "only_2024.xlsx"
        workbook = load_workbook(self.workbook_path)
        del workbook["2023"]
        workbook.save(new_year_book)
        workbook.close()
        result = import_excel(new_year_book, self.db_path)
        self.assertEqual(result.rankings_imported, 1)
        self.assertEqual(result.entries_added, 50)
        self.assertEqual(self._counts()["rankings"], 2)
        self.assertEqual(self._counts()["entries"], 100)


    def test_default_project_import_uses_temporary_database(self) -> None:
        # Проверяем именно штатный путь запуска: без передачи путей вручную.
        result = run_project_import()
        self.assertEqual(result.rankings_imported, 2)
        self.assertEqual(result.songs_added, 99)
        self.assertTrue(self.db_path.exists())
        self.assertEqual(self._counts()["entries"], 100)

    def test_default_project_import_uses_temporary_database(self) -> None:
        # Проверяем штатный запуск без путей и гарантируем изоляцию от основной БД.
        result = run_project_import()
        self.assertEqual(result.rankings_imported, 2)
        self.assertEqual(result.songs_added, 99)
        self.assertTrue(self.db_path.exists())
        self.assertEqual(self._counts()["entries"], 100)

    def test_song_artist_and_hall_of_fame_analytics(self) -> None:
        import_excel(self.workbook_path, self.db_path)

        song_stats = get_song_statistics(self.db_path)
        shared_track = next(item for item in song_stats if item.title == "Shared track")
        self.assertEqual(shared_track.appearances, 2)
        self.assertEqual(shared_track.best_position, 1)
        self.assertEqual(shared_track.average_position, 1.0)
        self.assertEqual(shared_track.total_points, 100)
        self.assertEqual(shared_track.positions_by_year, {2023: 1, 2024: 1})

        one_year_track = next(item for item in song_stats if item.title == "Song 2")
        self.assertEqual(one_year_track.positions_by_year, {2023: 2, 2024: None})
        self.assertEqual(one_year_track.average_position_with_absences, 26.5)
        self.assertEqual(one_year_track.total_points, 49)

        artist_stats = get_artist_statistics(self.db_path)
        artist_a = next(item for item in artist_stats if item.artist == "Artist A")
        self.assertEqual(artist_a.appearances, 2)
        self.assertEqual(artist_a.total_points, 100)
        self.assertEqual(artist_a.appearances_by_year, {2023: 1, 2024: 1})

        hall_of_fame = get_hall_of_fame(threshold=2, db_path=self.db_path)
        self.assertEqual(len(hall_of_fame), 1)
        self.assertEqual(hall_of_fame[0].title, "Shared track")
        self.assertEqual(hall_of_fame[0].induction_year, 2024)

    def test_hall_of_fame_rejects_invalid_threshold(self) -> None:
        with self.assertRaises(ValueError):
            get_hall_of_fame(threshold=0, db_path=self.db_path)

    def test_dry_run_does_not_create_database(self) -> None:
        result = import_excel(self.workbook_path, self.db_path, dry_run=True)
        self.assertTrue(result.dry_run)
        self.assertEqual(result.songs_added, 99)
        self.assertFalse(self.db_path.exists())

    def test_existing_song_metadata_conflict_aborts_without_changes(self) -> None:
        import_excel(self.workbook_path, self.db_path)
        before = self._snapshot()
        # Big in Japan встречается только в рейтинге 2023, поэтому изменение
        # метаданных конфликтует именно с записью в уже существующей БД,
        # а не с другой строкой исходной книги.
        changed = self._mutate_workbook(
            "2023", "D7", "France", self.root / "metadata_conflict.xlsx"
        )
        with self.assertRaises(ImportConflictError) as context:
            import_excel(changed, self.db_path, replace_existing=True)
        self.assertTrue(
            any("метаданные в базе отличаются" in item for item in context.exception.conflicts)
        )
        self.assertEqual(self._snapshot(), before)

    def test_same_title_year_different_artist_is_reported(self) -> None:
        import_excel(self.workbook_path, self.db_path)
        before = self._snapshot()
        # Та же уникальная песня и год выпуска, но другой исполнитель:
        # проверяем конфликт с каталогом SQLite.
        changed = self._mutate_workbook(
            "2023", "A7", "Different artist", self.root / "artist_conflict.xlsx"
        )
        with self.assertRaises(ImportConflictError) as context:
            import_excel(changed, self.db_path, replace_existing=True)
        self.assertTrue(
            any("в базе исполнители" in item for item in context.exception.conflicts)
        )
        self.assertEqual(self._snapshot(), before)

    def test_cross_year_metadata_conflict_inside_workbook_is_reported(self) -> None:
        changed = self._mutate_workbook(
            "2024", "D2", "Germany", self.root / "internal_conflict.xlsx"
        )
        with self.assertRaises(ImportConflictError) as context:
            read_workbook(changed)
        self.assertTrue(any("внутри Excel" in item for item in context.exception.conflicts))
        self.assertFalse(self.db_path.exists())

    def test_missing_required_value_is_rejected(self) -> None:
        changed = self._mutate_workbook(
            "2023", "C3", None, self.root / "missing_value.xlsx"
        )
        with self.assertRaises(WorkbookValidationError):
            read_workbook(changed)

    def test_non_integer_or_out_of_range_release_year_is_rejected(self) -> None:
        for invalid_year in (1888, 2008.5):
            changed = self._mutate_workbook(
                "2023", "E3", invalid_year, self.root / f"bad_year_{invalid_year}.xlsx"
            )
            with self.subTest(invalid_year=invalid_year):
                with self.assertRaises(WorkbookValidationError):
                    read_workbook(changed)

    def test_duplicate_song_inside_one_ranking_is_rejected(self) -> None:
        workbook = load_workbook(self.workbook_path)
        ws = workbook["2023"]
        for column in range(1, 6):
            ws.cell(row=3, column=column).value = ws.cell(row=2, column=column).value
        destination = self.root / "duplicate_song.xlsx"
        workbook.save(destination)
        workbook.close()
        with self.assertRaises(WorkbookValidationError) as context:
            read_workbook(destination)
        self.assertTrue(any("дубликат песни" in item for item in context.exception.errors))

    def test_wrong_top_size_is_rejected(self) -> None:
        workbook = load_workbook(self.workbook_path)
        workbook["2023"].delete_rows(51)
        destination = self.root / "short_ranking.xlsx"
        workbook.save(destination)
        workbook.close()
        with self.assertRaises(WorkbookValidationError) as context:
            read_workbook(destination)
        self.assertTrue(any("ожидалось ровно 50" in item for item in context.exception.errors))

    def test_empty_element_in_semicolon_list_is_rejected(self) -> None:
        changed = self._mutate_workbook(
            "2023", "A3", "Artist 2;;Artist X", self.root / "bad_list.xlsx"
        )
        with self.assertRaises(WorkbookValidationError) as context:
            read_workbook(changed)
        self.assertTrue(any("пустой элемент" in item for item in context.exception.errors))

    def test_transaction_rolls_back_all_rows_after_sql_error(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws.title = "2026"
        ws.append(HEADERS)
        for number in range(1, 51):
            ws.append([f"Artist {number}", f"New song {number}", "English", "USA", 2020])
        source = self.root / "2026.xlsx"
        workbook.save(source)
        workbook.close()

        initialize_database(self.db_path)
        with get_connection(self.db_path) as connection:
            connection.execute(
                """
                CREATE TRIGGER fail_second_entry
                BEFORE INSERT ON ranking_entries
                WHEN NEW.position = 2
                BEGIN
                    SELECT RAISE(ABORT, 'forced test failure');
                END;
                """
            )

        with self.assertRaises(sqlite3.IntegrityError):
            import_excel(source, self.db_path)

        with get_connection(self.db_path) as connection:
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM songs").fetchone()[0], 0)
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM rankings").fetchone()[0], 0)
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM ranking_entries").fetchone()[0], 0)
            self.assertEqual(connection.execute("SELECT COUNT(*) FROM artists").fetchone()[0], 0)


    def test_schema_trigger_rejects_position_above_top_size(self) -> None:
        initialize_database(self.db_path)
        with get_connection(self.db_path) as connection:
            ranking_id = connection.execute(
                "INSERT INTO rankings (ranking_year, top_size) VALUES (2026, 1)"
            ).lastrowid
            song_id = connection.execute(
                "INSERT INTO songs (title, release_year) VALUES ('Trigger test', 2020)"
            ).lastrowid
            with self.assertRaises(sqlite3.IntegrityError):
                connection.execute(
                    "INSERT INTO ranking_entries (ranking_id, song_id, position) VALUES (?, ?, 2)",
                    (ranking_id, song_id),
                )

    def test_available_user_workbook_validates(self) -> None:
        if not SOURCE_WORKBOOK.is_file():
            self.skipTest("Пользовательская Excel-книга не подключена к этому окружению")
        data = read_workbook(SOURCE_WORKBOOK)
        self.assertEqual(data.ranking_years, [2023, 2024, 2025])
        self.assertEqual(len(data.songs), 150)
        self.assertEqual(len(data.unique_songs), 119)

    def test_available_user_workbook_import_and_replacement(self) -> None:
        if not SOURCE_WORKBOOK.is_file():
            self.skipTest("Пользовательская Excel-книга не подключена к этому окружению")
        result = import_excel(SOURCE_WORKBOOK, self.db_path)
        self.assertEqual(result.songs_added, 119)
        self.assertEqual(result.entries_added, 150)
        before = self._counts()

        with self.assertRaises(ImportConflictError):
            import_excel(SOURCE_WORKBOOK, self.db_path)
        self.assertEqual(self._counts(), before)

        replacement = import_excel(SOURCE_WORKBOOK, self.db_path, replace_existing=True)
        self.assertEqual(replacement.songs_added, 0)
        self.assertEqual(replacement.replaced_years, (2023, 2024, 2025))
        self.assertEqual(self._counts(), before)


if __name__ == "__main__":
    unittest.main()
