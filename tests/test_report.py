from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.database import get_connection
from src.importer import import_excel
from src.report import generate_excel_report
from src.analytics import get_composition_statistics


HEADERS = ["Исполнители", "Песня", "Язык", "Страна", "Год"]


class ExcelReportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db_path = self.root / "music.db"
        self.workbook_path = self.root / "source.xlsx"
        self.output_path = self.root / "nested" / "report.xlsx"
        self._create_workbook()
        import_excel(self.workbook_path, self.db_path)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()
        self.assertFalse(self.db_path.exists(), "Тестовая база должна удаляться")

    def _create_workbook(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "2023"
        sheet.append(HEADERS)
        sheet.append(["Artist A; Artist B", "Shared track", "English; Russian", "USA; Sweden", 2001])
        for number in range(2, 51):
            sheet.append([f"Artist {number}", f"Song {number}", "English", "USA", 1990 + number])

        sheet = workbook.create_sheet("2024")
        sheet.append(HEADERS)
        sheet.append(["Other Artist 1", "Other song 1", "English", "USA", 2001])
        sheet.append(["Artist B; Artist A", "Shared track", "English; Russian", "USA; Sweden", 2001])
        for number in range(2, 50):
            sheet.append([f"Other Artist {number}", f"Other song {number}", "English", "USA", 2000 + number])
        workbook.save(self.workbook_path)
        workbook.close()

    def _database_snapshot(self) -> dict[str, list[tuple]]:
        tables = [
            "countries", "languages", "artists", "songs", "song_artists",
            "song_countries", "song_languages", "rankings", "ranking_entries",
        ]
        with get_connection(self.db_path) as connection:
            return {
                table: [tuple(row) for row in connection.execute(f"SELECT * FROM {table} ORDER BY 1, 2").fetchall()]
                for table in tables
            }

    def test_report_has_four_sheets_and_expected_values(self) -> None:
        before = self._database_snapshot()
        result = generate_excel_report(self.db_path, self.output_path)
        self.assertEqual(result, self.output_path)
        self.assertTrue(result.is_file())
        self.assertEqual(before, self._database_snapshot(), "Генерация отчёта не должна менять БД")

        workbook = load_workbook(result, data_only=True)
        self.assertEqual(
            workbook.sheetnames,
            [
                "Статистика песен", "Статистика исполнителей", "Устойчивость исполнителей",
                "Зал славы", "Статистика по годам", "Состав по годам", "Общий состав",
                "_Данные графиков", "Графики по годам", "Общие графики", "Методика",
            ],
        )

        songs = workbook["Статистика песен"]
        self.assertEqual(songs.freeze_panes, "A2")
        self.assertEqual(songs.auto_filter.ref, songs.dimensions)
        self.assertEqual(songs.max_row, 100)  # заголовок + 99 уникальных песен
        self.assertEqual(songs.cell(2, 1).value, "Shared track")
        self.assertEqual(songs.cell(2, 4).value, 2)
        self.assertEqual(songs.cell(2, 8).value, 99)

        yearly = workbook["Статистика по годам"]
        self.assertEqual(yearly.max_row, 3)
        self.assertEqual([yearly.cell(2, col).value for col in range(1, 6)], [2023, 50, 50, 50, 0])
        self.assertEqual([yearly.cell(3, col).value for col in range(1, 6)], [2024, 50, 50, 49, 1])
        self.assertAlmostEqual(yearly.cell(3, 6).value, 0.02)

        composition = workbook["Состав по годам"]
        # Строки содержат только категории, реально представленные в рейтинге.
        self.assertEqual(composition.cell(1, 1).value, "Год рейтинга")
        self.assertEqual(composition.cell(1, 4).value, "Количество песен")
        self.assertTrue(any(composition.cell(row, 2).value == "Страна" for row in range(2, composition.max_row + 1)))
        self.assertEqual(workbook["_Данные графиков"].sheet_state, "hidden")
        self.assertEqual(len(workbook["Графики по годам"]._charts), 4)
        self.assertEqual(len(workbook["Общие графики"]._charts), 4)
        self.assertEqual(workbook["Методика"].max_row, 13)

        # В годовых диаграммах нет ошибочной серии «Категория»:
        # подписи категорий используются только как названия рядов.
        for chart in workbook["Графики по годам"]._charts:
            self.assertTrue(chart.series)
            for series in chart.series:
                self.assertRegex(series.tx.strRef.f, r"!\$?A\$?[0-9]+$")
                self.assertRegex(series.val.numRef.f, r"!\$?B\$?[0-9]+:\$?C\$?[0-9]+$")

        stability = workbook["Устойчивость исполнителей"]
        stability_headers = [cell.value for cell in stability[1]]
        artist_a_stability = next(
            row for row in range(2, stability.max_row + 1)
            if stability.cell(row, 1).value == "Artist A"
        )
        self.assertEqual(stability.cell(artist_a_stability, stability_headers.index("Рейтингов с присутствием") + 1).value, 2)
        self.assertEqual(stability.cell(artist_a_stability, stability_headers.index("Самая длинная серия рейтингов подряд") + 1).value, 2)

        # Для песни, присутствующей в обоих годах, дельта позиции и баллов рассчитана.
        headers = [cell.value for cell in songs[1]]
        delta_position_col = headers.index("2024: Δ позиции*") + 1
        delta_points_col = headers.index("2024: Δ баллов") + 1
        self.assertEqual(songs.cell(2, delta_position_col).value, -1)
        self.assertEqual(songs.cell(2, delta_points_col).value, -1)

        artists = workbook["Статистика исполнителей"]
        artist_headers = [cell.value for cell in artists[1]]
        artist_row = next(
            row for row in range(2, artists.max_row + 1)
            if artists.cell(row, 1).value == "Artist A"
        )
        self.assertEqual(artists.cell(artist_row, artist_headers.index("2024: Δ лучшей позиции*") + 1).value, -1)
        self.assertEqual(artists.cell(artist_row, artist_headers.index("2024: Δ средней позиции*") + 1).value, -1)
        self.assertEqual(artists.cell(artist_row, artist_headers.index("2024: Δ баллов") + 1).value, -1)
        workbook.close()

    def test_composition_statistics_support_multiple_countries_and_languages(self) -> None:
        items = get_composition_statistics(self.db_path, interval_count=7)
        self.assertTrue(items)

        yearly_country = next(
            item for item in items
            if item.year == 2023 and item.dimension == "Страна" and item.category == "Sweden"
        )
        yearly_language = next(
            item for item in items
            if item.year == 2023 and item.dimension == "Язык" and item.category == "Russian"
        )
        overall_country = next(
            item for item in items
            if item.year is None and item.dimension == "Страна" and item.category == "Sweden"
        )
        self.assertEqual(yearly_country.songs_count, 1)
        self.assertEqual(yearly_language.songs_count, 1)
        self.assertEqual(overall_country.songs_count, 1)
        self.assertAlmostEqual(overall_country.share, 1 / 99)

    def test_hall_of_fame_threshold_is_configurable(self) -> None:
        generate_excel_report(self.db_path, self.output_path, hall_of_fame_threshold=2)
        workbook = load_workbook(self.output_path, data_only=True)
        hall = workbook["Зал славы"]
        self.assertEqual(hall.max_row, 2)
        self.assertEqual(hall.cell(2, 1).value, "Shared track")
        self.assertEqual(hall.cell(2, 5).value, 2024)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
