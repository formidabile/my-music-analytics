from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.forecasts import import_forecasts, get_forecast_analysis, get_delayed_forecasts
from src.forecast_report import generate_forecast_report
from src.importer import import_excel

HEADERS = ["Исполнители", "Песня", "Язык", "Страна", "Год"]


class ForecastTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = self.root / "music.db"
        self.actual = self.root / "actual.xlsx"
        self.forecasts = self.root / "forecasts.xlsx"
        self._make_actual()
        self._make_forecasts()
        import_excel(self.actual, self.db)
        import_forecasts(self.forecasts, self.db)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _make_actual(self) -> None:
        wb = Workbook()
        for year in (2024, 2025):
            ws = wb.active if year == 2024 else wb.create_sheet()
            ws.title = str(year)
            ws.append(HEADERS)
            if year == 2024:
                ws.append(["Artist A", "Present 2024", "English", "USA", 2000])
            else:
                ws.append(["Artist B", "Delayed", "English", "USA", 2001])
                ws.append(["Artist C", "Repeat", "English", "USA", 2002])
            while ws.max_row < 51:
                n = ws.max_row
                ws.append([f"Artist {year}-{n}", f"Track {year}-{n}", "English", "USA", 1980+n])
        wb.save(self.actual)
        wb.close()

    def _make_forecasts(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "2024"
        ws.append(HEADERS)
        ws.append(["Artist A", "Present 2024", "English", "USA", 2000])
        ws.append(["Artist B", "Delayed", "English", "USA", 2001])
        ws.append(["Artist C", "Repeat", "English", "USA", 2002])
        ws.append(["Artist D", "Never realized", "English", "USA", 2003])
        ws = wb.create_sheet("2025")
        ws.append(HEADERS)
        ws.append(["Artist C", "Repeat", "English", "USA", 2002])
        wb.save(self.forecasts)
        wb.close()

    def test_forecast_report_is_separate_and_contains_expected_sheets(self) -> None:
        output = self.root / "forecast_report.xlsx"
        generate_forecast_report(self.db, output)
        workbook = load_workbook(output, data_only=True)
        self.assertEqual(workbook.sheetnames, [
            "Сводка прогнозов", "Сопоставление прогнозов", "Неожиданные попадания",
            "Отложенные попадания", "Повторные прогнозы", "Нереализованные прогнозы", "Методика",
        ])
        summary = workbook["Сводка прогнозов"]
        self.assertEqual(summary.cell(2, 1).value, 2024)
        self.assertEqual(summary.cell(2, 4).value, 1)
        self.assertEqual(workbook["Отложенные попадания"].cell(2, 1).value, 2024)
        self.assertEqual(workbook["Повторные прогнозы"].cell(2, 1).value, "Repeat")
        self.assertTrue(any(
            workbook["Нереализованные прогнозы"].cell(row, 3).value == "Never realized"
            for row in range(2, workbook["Нереализованные прогнозы"].max_row + 1)
        ))
        workbook.close()

    def test_metrics_and_delayed_repeated_forecast(self) -> None:
        metrics, comparisons, unexpected = get_forecast_analysis(self.db)
        by_year = {m.year: m for m in metrics}
        self.assertEqual(by_year[2024].predicted_count, 4)
        self.assertEqual(by_year[2024].hits, 1)
        self.assertEqual(by_year[2025].hits, 1)
        self.assertEqual(by_year[2024].unexpected, 49)
        delayed = get_delayed_forecasts(self.db)
        item = next(x for x in delayed if x.title == "Delayed")
        self.assertEqual(item.forecast_year, 2024)
        self.assertEqual(item.first_actual_year, 2025)
        self.assertEqual(item.delay_years, 1)
        self.assertEqual(item.first_actual_position, 1)
        repeated = [x for x in comparisons if x.title == "Repeat"]
        self.assertEqual([x.status for x in repeated], ["DELAYED", "HIT"])


if __name__ == "__main__":
    unittest.main()
