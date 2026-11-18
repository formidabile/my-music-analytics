"""Формирование отдельной Excel-книги с аналитикой прогнозов."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.database import BASE_DIR, DATABASE_PATH
from src.forecasts import ForecastComparison, get_delayed_forecasts, get_forecast_analysis

DEFAULT_FORECAST_REPORT_PATH = BASE_DIR / "reports" / "music_forecast_analytics.xlsx"
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10, color="202124")
ALT_FILL = PatternFill(fill_type="solid", fgColor="F3F6F9")


def _artists(items: Sequence[str]) -> str:
    return "; ".join(items)


def _write_sheet(workbook: Workbook, title: str, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    ws = workbook.create_sheet(title)
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 34
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_idx in range(2, ws.max_row + 1):
        for cell in ws[row_idx]:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = ALT_FILL
    for col in range(1, ws.max_column + 1):
        values = [str(ws.cell(r, col).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[get_column_letter(col)].width = min(max(max(map(len, values), default=10) + 2, 12), 38)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _comparison_row(item: ForecastComparison) -> tuple[object, ...]:
    return (
        item.forecast_year, item.predicted_position, item.title, _artists(item.artists),
        item.release_year, item.actual_position, item.position_error,
        abs(item.position_error) if item.position_error is not None else None,
        item.status, item.first_actual_year, item.first_actual_position, item.delay_years,
    )


def generate_forecast_report(
    db_path: str | Path | None = None,
    output_path: str | Path | None = None,
) -> Path:
    """Создаёт книгу аналитики прогнозов, не изменяя базу данных.

    Если прогнозов в базе нет, создаётся книга с листом-пояснением.
    """
    database_path = Path(db_path) if db_path is not None else DATABASE_PATH
    destination = Path(output_path) if output_path is not None else DEFAULT_FORECAST_REPORT_PATH
    metrics, comparisons, unexpected = get_forecast_analysis(database_path)
    delayed = get_delayed_forecasts(database_path)

    wb = Workbook()
    wb.remove(wb.active)
    if not metrics:
        _write_sheet(wb, "Информация", ["Статус"], [["В базе данных пока нет прогнозов."]])
    else:
        metric_rows = [(
            m.year, m.predicted_count, m.actual_count, m.hits, m.misses, m.unexpected,
            m.precision, m.recall, m.mae, m.exact_matches, m.within_5, m.within_10,
        ) for m in metrics]
        _write_sheet(wb, "Сводка прогнозов", [
            "Год прогноза", "Песен в прогнозе", "Песен в фактическом Top-50", "Попаданий",
            "Не попали", "Неожиданных песен", "Precision", "Recall", "MAE",
            "Точная позиция", "Ошибка ≤ 5", "Ошибка ≤ 10",
        ], metric_rows)
        for col in (7, 8):
            for cell in wb["Сводка прогнозов"][get_column_letter(col)][1:]:
                cell.number_format = "0.0%"

        comparison_rows = [_comparison_row(x) for x in comparisons]
        _write_sheet(wb, "Сопоставление прогнозов", [
            "Год прогноза", "Прогнозная позиция", "Песня", "Исполнители", "Год выпуска",
            "Фактическая позиция в этом году", "Ошибка позиции*", "Абсолютная ошибка",
            "Статус", "Первый год в Top-50", "Позиция первого попадания", "Задержка, лет",
        ], comparison_rows)
        _write_sheet(wb, "Неожиданные попадания", [
            "Год рейтинга", "Фактическая позиция", "Песня", "Исполнители", "Год выпуска",
        ], [(
            x.forecast_year, x.actual_position, x.title, _artists(x.artists), x.release_year,
        ) for x in unexpected])
        _write_sheet(wb, "Отложенные попадания", [
            "Первый год прогноза", "Первая прогнозная позиция", "Песня", "Исполнители",
            "Год выпуска", "Первый год в Top-50", "Позиция первого попадания", "Задержка, лет",
        ], [(
            x.forecast_year, x.predicted_position, x.title, _artists(x.artists), x.release_year,
            x.first_actual_year, x.first_actual_position, x.delay_years,
        ) for x in delayed])

        by_song: dict[tuple, list[ForecastComparison]] = defaultdict(list)
        for item in comparisons:
            key = (item.title.casefold(), item.release_year, tuple(sorted(a.casefold() for a in item.artists)))
            by_song[key].append(item)
        repeated_rows = []
        for events in by_song.values():
            if len(events) < 2:
                continue
            first = min(events, key=lambda x: x.forecast_year)
            realized = next((x for x in events if x.first_actual_year is not None), first)
            repeated_rows.append((
                first.title, _artists(first.artists), first.release_year,
                ", ".join(str(x.forecast_year) for x in sorted(events, key=lambda x: x.forecast_year)),
                len(events), realized.first_actual_year, realized.first_actual_position,
                realized.delay_years, ", ".join(x.status for x in sorted(events, key=lambda x: x.forecast_year)),
            ))
        repeated_rows.sort(key=lambda r: (str(r[0]).casefold(), str(r[1]).casefold()))
        _write_sheet(wb, "Повторные прогнозы", [
            "Песня", "Исполнители", "Год выпуска", "Годы прогноза", "Число прогнозов",
            "Первый год в Top-50", "Позиция первого попадания", "Задержка, лет", "Статусы прогнозов",
        ], repeated_rows)

        pending_rows = [(
            x.forecast_year, x.predicted_position, x.title, _artists(x.artists), x.release_year,
        ) for x in comparisons if x.status == "PENDING"]
        _write_sheet(wb, "Нереализованные прогнозы", [
            "Год прогноза", "Прогнозная позиция", "Песня", "Исполнители", "Год выпуска",
        ], pending_rows)

        notes = [
            ("Precision", "Попадания / число песен в прогнозе. Непредсказанные позиции неполного прогноза не считаются ошибками."),
            ("Recall", "Попадания / число песен фактического Top-50."),
            ("MAE", "Среднее абсолютное отклонение позиций только для песен, попавших в Top-50 в год прогноза."),
            ("Ошибка позиции", "Фактическая позиция минус прогнозная. Положительное значение означает, что песня оказалась ниже прогноза."),
            ("Отложенное попадание", "Песня не попала в Top-50 в год прогноза, но впервые появилась в одном из последующих фактических рейтингов. Задержка считается от первого года прогноза."),
            ("Повторные прогнозы", "Каждый прогноз одной песни в отдельном году оценивается самостоятельно; в листе показана история таких прогнозов."),
            ("Нереализованный прогноз", "На момент формирования отчёта песня не появилась ни в одном фактическом рейтинге начиная с года первого прогноза."),
        ]
        _write_sheet(wb, "Методика", ["Показатель", "Правило расчёта"], notes)

    destination.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destination)
    wb.close()
    return destination


if __name__ == "__main__":
    path = generate_forecast_report()
    print(f"Отчёт по прогнозам сохранён: {path}")
