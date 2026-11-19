from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.database import get_connection, initialize_database
from src.streaming import StreamingImportError, import_streaming_workbook, read_streaming_workbook, get_streaming_statistics
from src.streaming_report import generate_streaming_report
from src.importer import import_excel


class StreamingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = self.root / 'music.db'
        self.xlsx = self.root / 'streaming.xlsx'
        initialize_database(self.db)
        with get_connection(self.db) as conn:
            conn.executemany('INSERT INTO artists(artist_name) VALUES (?)', [('Lady Gaga',), ('Катя Лель',), ('Винтаж',)])
            conn.commit()

    def tearDown(self):
        self.temp.cleanup()

    def create_book(self, rows):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Стриминги'
        ws.append(['Год', 'Исполнитель', 'Время прослушивания'])
        for row in rows:
            ws.append(row)
        wb.save(self.xlsx)
        wb.close()

    def test_positions_are_derived_and_missing_minutes_are_null(self):
        self.create_book([(2024, 'Lady Gaga', 897), (2024, 'Катя Лель', None), (2024, 'Винтаж', None)])
        self.assertEqual(import_streaming_workbook(self.xlsx, self.db), 3)
        rows = get_streaming_statistics(self.db)
        self.assertEqual([row['position'] for row in rows], [1, 2, 3])
        self.assertEqual(rows[0]['listening_minutes'], 897)
        self.assertIsNone(rows[1]['listening_minutes'])

    def test_unknown_artist_fails_without_partial_write(self):
        self.create_book([(2024, 'Lady Gaga', 897), (2024, 'Unknown Artist', 100)])
        with self.assertRaises(StreamingImportError):
            import_streaming_workbook(self.xlsx, self.db)
        self.assertEqual(get_streaming_statistics(self.db), [])

    def test_streaming_report_compares_top5_with_yearly_top50(self):
        actual = self.root / 'actual.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '2024'
        ws.append(['Исполнители', 'Песня', 'Язык', 'Страна', 'Год'])
        ws.append(['Lady Gaga', 'Track A', 'English', 'USA', 2020])
        ws.append(['Катя Лель', 'Track B', 'Russian', 'Russia', 2000])
        while ws.max_row < 51:
            n = ws.max_row
            ws.append([f'Artist {n}', f'Track {n}', 'English', 'USA', 1980])
        wb.save(actual)
        wb.close()
        import_excel(actual, self.db)
        self.create_book([(2024, 'Lady Gaga', 897), (2024, 'Катя Лель', None), (2024, 'Винтаж', None)])
        import_streaming_workbook(self.xlsx, self.db)
        output = self.root / 'streaming_report.xlsx'
        generate_streaming_report(self.db, output)
        from openpyxl import load_workbook
        report = load_workbook(output, data_only=True)
        ws = report['Стриминги и Top-50']
        self.assertEqual(ws.max_row, 4)
        self.assertEqual(ws.cell(2, 3).value, 'Lady Gaga')
        self.assertEqual(ws.cell(2, 4).value, 897)
        self.assertEqual(ws.cell(2, 7).value, 1)
        self.assertEqual(ws.cell(2, 10).value, 'Да')
        self.assertIsNone(ws.cell(3, 4).value)
        self.assertEqual(ws.cell(4, 10).value, 'Нет')
        all_time = report['За всё время']
        self.assertEqual(all_time.max_row, 4)
        self.assertEqual(all_time.cell(2, 1).value, 'Lady Gaga')
        self.assertEqual(all_time.cell(2, 2).value, 1)
        self.assertEqual(all_time.cell(2, 5).value, 1)
        self.assertEqual(all_time.cell(2, 10).value, 897)
        self.assertEqual(all_time.cell(2, 12).value, 1)
        self.assertIsNone(all_time.cell(3, 10).value)
        report.close()

    def test_all_time_streaming_sheet_aggregates_years_and_known_minutes(self):
        actual = self.root / 'actual.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '2024'
        ws.append(['Исполнители', 'Песня', 'Язык', 'Страна', 'Год'])
        ws.append(['Lady Gaga', 'Track A', 'English', 'USA', 2020])
        while ws.max_row < 51:
            n = ws.max_row
            ws.append([f'Artist {n}', f'Track {n}', 'English', 'USA', 1980])
        wb.save(actual)
        wb.close()
        import_excel(actual, self.db)

        self.create_book([
            (2024, 'Lady Gaga', 897), (2024, 'Катя Лель', None),
            (2025, 'Lady Gaga', 1292), (2025, 'Катя Лель', 500),
        ])
        import_streaming_workbook(self.xlsx, self.db)
        output = self.root / 'streaming_report.xlsx'
        generate_streaming_report(self.db, output)
        from openpyxl import load_workbook
        report = load_workbook(output, data_only=True)
        ws = report['За всё время']
        rows = {ws.cell(row, 1).value: row for row in range(2, ws.max_row + 1)}
        lady_row = rows['Lady Gaga']
        self.assertEqual(ws.cell(lady_row, 2).value, 2)
        self.assertEqual(ws.cell(lady_row, 4).value, 2025)
        self.assertEqual(ws.cell(lady_row, 5).value, 1)
        self.assertEqual(ws.cell(lady_row, 10).value, 2189)
        self.assertEqual(ws.cell(lady_row, 11).value, 1094.5)
        self.assertEqual(ws.cell(lady_row, 12).value, 2)
        katya_row = rows['Катя Лель']
        self.assertEqual(ws.cell(katya_row, 10).value, 500)
        self.assertEqual(ws.cell(katya_row, 12).value, 1)
        # Сортировка листа «За всё время» определяется общим количеством минут.
        self.assertEqual(ws.cell(2, 1).value, 'Lady Gaga')
        self.assertEqual(ws.cell(2, 10).value, 2189)
        self.assertEqual(ws.cell(3, 1).value, 'Катя Лель')
        self.assertEqual(ws.cell(3, 10).value, 500)
        report.close()

    def test_more_than_five_artists_is_rejected(self):
        self.create_book([(2024, name, 10) for name in ['Lady Gaga', 'Катя Лель', 'Винтаж', 'A', 'B', 'C']])
        with self.assertRaises(StreamingImportError):
            read_streaming_workbook(self.xlsx)


if __name__ == '__main__':
    unittest.main()
