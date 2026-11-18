"""Формирование расширенного Excel-отчёта по музыкальным рейтингам."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.analytics import (
    ArtistStats,
    ArtistStabilityStats,
    CompositionStats,
    HallOfFameEntry,
    SongStats,
    YearStats,
    get_artist_statistics,
    get_artist_stability_statistics,
    get_composition_statistics,
    get_hall_of_fame,
    get_song_statistics,
    get_yearly_statistics,
)
from src.database import BASE_DIR, DATABASE_PATH

DEFAULT_REPORT_PATH = BASE_DIR / "reports" / "music_analytics.xlsx"

HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10, color="202124")
ALT_FILL = PatternFill(fill_type="solid", fgColor="F3F6F9")
DIMENSIONS = ["Страна", "Язык", "Десятилетие", "Интервал года выпуска"]
DIMENSION_KEYS = {
    "Страна": "Страна",
    "Язык": "Язык",
    "Десятилетие": "Десятилетие",
    "Интервал года выпуска": "Интервал года выпуска",
}


def _artists_text(artists: Sequence[str]) -> str:
    return "; ".join(artists)


def _write_sheet(
    workbook: Workbook,
    title: str,
    headers: Sequence[str],
    rows: Iterable[Sequence[object]],
    *,
    widths: dict[int, float] | None = None,
    numeric_columns: set[int] | None = None,
) -> None:
    worksheet = workbook.create_sheet(title)
    worksheet.append(list(headers))
    for row in rows:
        worksheet.append(list(row))

    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.row_dimensions[1].height = 34

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index in range(2, worksheet.max_row + 1):
        for cell in worksheet[row_index]:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center")
            if row_index % 2 == 0:
                cell.fill = ALT_FILL
            if numeric_columns and cell.column in numeric_columns and isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(cell.value, float):
                    cell.number_format = "0.00"

    for column_index, _header in enumerate(headers, start=1):
        if widths and column_index in widths:
            width = widths[column_index]
        else:
            values = [str(worksheet.cell(row, column_index).value or "") for row in range(1, worksheet.max_row + 1)]
            width = min(max(max((len(value) for value in values), default=10) + 2, 12), 34)
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.outlinePr.summaryBelow = True


def _song_summary_rows(songs: list[SongStats]) -> list[tuple[object, ...]]:
    return [
        (
            song.title, _artists_text(song.artists), song.release_year, song.appearances,
            song.best_position, song.average_position,
            song.average_position_with_absences, song.total_points,
        )
        for song in songs
    ]


def _artist_summary_rows(artists: list[ArtistStats]) -> list[tuple[object, ...]]:
    return [
        (artist.artist, artist.appearances, artist.best_position,
         artist.average_position, artist.total_points)
        for artist in artists
    ]


def _song_year_rows(songs: list[SongStats], years: list[int], year: int) -> list[tuple[object, ...]]:
    """Возвращает накопительный список песен, появлявшихся не позднее указанного года."""
    index = years.index(year)
    previous_year = years[index - 1] if index > 0 else None
    rows: list[tuple[object, ...]] = []
    for song in songs:
        history = years[:index + 1]
        if not any(song.positions_by_year.get(item) is not None for item in history):
            continue

        position = song.positions_by_year.get(year)
        previous_position = song.positions_by_year.get(previous_year) if previous_year is not None else None
        previous_cumulative = sum(song.points_by_year.get(item, 0) for item in years[:index])
        current_year_points = song.points_by_year.get(year, 0)
        cumulative = previous_cumulative + current_year_points
        points_delta = cumulative - previous_cumulative if previous_year is not None else None
        position_delta = previous_position - position if previous_position is not None and position is not None else None

        earlier_appearances = [item for item in years[:index] if song.positions_by_year.get(item) is not None]
        if position is None:
            status = "OUT"
        elif not earlier_appearances:
            status = "NEW"
        elif previous_position is None:
            status = "RE"
        else:
            status = ""

        rows.append((
            song.title, _artists_text(song.artists), song.release_year, position,
            current_year_points, cumulative, points_delta, position_delta, status,
        ))

    # Лидеры каждого исторического среза определяются накопленными баллами на конец года.
    rows.sort(key=lambda row: (-int(row[5]), str(row[0]).casefold(), str(row[1]).casefold()))
    return rows


def _artist_year_rows(artists: list[ArtistStats], years: list[int], year: int) -> list[tuple[object, ...]]:
    """Возвращает накопительный список исполнителей, появлявшихся не позднее указанного года."""
    index = years.index(year)
    previous_year = years[index - 1] if index > 0 else None
    rows: list[tuple[object, ...]] = []
    for artist in artists:
        history = years[:index + 1]
        if not any(artist.appearances_by_year.get(item, 0) > 0 for item in history):
            continue

        appearances = artist.appearances_by_year.get(year, 0)
        best_position = artist.best_positions_by_year.get(year)
        average_position = artist.average_positions_by_year.get(year)
        previous_best = artist.best_positions_by_year.get(previous_year) if previous_year is not None else None
        previous_average = artist.average_positions_by_year.get(previous_year) if previous_year is not None else None
        previous_cumulative = sum(artist.points_by_year.get(item, 0) for item in years[:index])
        current_year_points = artist.points_by_year.get(year, 0)
        cumulative = previous_cumulative + current_year_points
        points_delta = cumulative - previous_cumulative if previous_year is not None else None
        best_delta = previous_best - best_position if previous_best is not None and best_position is not None else None
        average_delta = previous_average - average_position if previous_average is not None and average_position is not None else None

        earlier_appearances = [item for item in years[:index] if artist.appearances_by_year.get(item, 0) > 0]
        if not appearances:
            status = "OUT"
        elif not earlier_appearances:
            status = "NEW"
        elif not artist.appearances_by_year.get(previous_year, 0):
            status = "RE"
        else:
            status = ""

        rows.append((
            artist.artist, appearances, best_position, average_position,
            current_year_points, cumulative, points_delta,
            best_delta, average_delta, status,
        ))

    rows.sort(key=lambda row: (-int(row[5]), str(row[0]).casefold()))
    return rows


def _write_yearly_report(
    songs: list[SongStats], artists: list[ArtistStats], years: list[int], destination: Path,
) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for year in years:
        song_headers = [
            "Песня", "Исполнители", "Год выпуска", "Позиция", "Баллы за год",
            "Общие баллы на конец года", "Δ общих баллов", "Δ позиции*", "Статус",
        ]
        _write_sheet(
            workbook, f"{year} Песни", song_headers,
            _song_year_rows(songs, years, year),
            widths={1: 30, 2: 34, 6: 25}, numeric_columns=set(range(3, 9)),
        )
        artist_headers = [
            "Исполнитель", "Песен в рейтинге", "Лучшая позиция", "Средняя позиция",
            "Баллы за год", "Общие баллы на конец года", "Δ общих баллов",
            "Δ лучшей позиции*", "Δ средней позиции*", "Статус",
        ]
        _write_sheet(
            workbook, f"{year} Исполнители", artist_headers,
            _artist_year_rows(artists, years, year),
            widths={1: 32, 6: 25}, numeric_columns=set(range(2, 10)),
        )
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    return destination


def _artist_stability_rows(items: list[ArtistStabilityStats]) -> list[tuple[object, ...]]:
    return [
        (
            item.artist, item.years_present, item.total_ranking_years, item.presence_rate,
            item.longest_streak, item.current_streak, item.max_tracks_in_year,
            item.average_tracks_per_present_year, item.top_track_title,
            item.top_track_points, item.top_track_points_share,
        )
        for item in items
    ]


def _hall_of_fame_rows(entries: list[HallOfFameEntry]) -> list[tuple[object, ...]]:
    return [
        (entry.title, _artists_text(entry.artists), entry.release_year, entry.appearances,
         entry.induction_year, entry.best_position, entry.average_position, entry.total_points)
        for entry in entries
    ]


def _year_rows(years: list[YearStats]) -> list[tuple[object, ...]]:
    return [
        (item.year, item.top_size, item.entries, item.new_songs, item.returning_songs,
         item.returning_songs / item.entries if item.entries else 0.0, item.average_position)
        for item in years
    ]


def _composition_rows(items: list[CompositionStats], year: int | None = None) -> list[tuple[object, ...]]:
    selected = [item for item in items if item.year == year]
    return [
        (item.dimension, item.category, item.songs_count, item.share)
        for item in sorted(selected, key=lambda x: (DIMENSIONS.index(x.dimension), -x.songs_count, x.category.casefold()))
    ]


def _add_composition_sheets(workbook: Workbook, items: list[CompositionStats], years: list[int]) -> None:
    yearly_rows = [
        (item.year, item.dimension, item.category, item.songs_count, item.share)
        for item in items if item.year is not None
    ]
    yearly_rows.sort(key=lambda x: (x[0], DIMENSIONS.index(x[1]), -x[3], x[2].casefold()))
    _write_sheet(
        workbook,
        "Состав по годам",
        ["Год рейтинга", "Показатель", "Категория", "Количество песен", "Доля в рейтинге"],
        yearly_rows,
        widths={1: 16, 2: 25, 3: 28, 5: 18},
        numeric_columns={1, 4, 5},
    )
    ws = workbook["Состав по годам"]
    for cell in ws["E"][1:]:
        cell.number_format = "0.0%"

    overall_rows = _composition_rows(items, None)
    _write_sheet(
        workbook,
        "Общий состав",
        ["Показатель", "Категория", "Уникальных песен", "Доля от всех песен"],
        overall_rows,
        widths={1: 25, 2: 30, 4: 20},
        numeric_columns={3, 4},
    )
    ws = workbook["Общий состав"]
    for cell in ws["D"][1:]:
        cell.number_format = "0.0%"

    _add_composition_charts(workbook, items, years)


def _add_composition_charts(workbook: Workbook, items: list[CompositionStats], years: list[int]) -> None:
    """Строит четыре графика динамики и четыре графика общей структуры."""
    data_ws = workbook.create_sheet("_Данные графиков")
    data_ws.sheet_state = "hidden"
    yearly_chart_ws = workbook.create_sheet("Графики по годам")
    yearly_chart_ws.sheet_view.showGridLines = False
    overall_chart_ws = workbook.create_sheet("Общие графики")
    overall_chart_ws.sheet_view.showGridLines = False

    yearly_chart_ws["A1"] = "Динамика состава годовых рейтингов"
    yearly_chart_ws["A1"].font = Font(size=16, bold=True, color="1F4E78")
    overall_chart_ws["A1"] = "Структура всех уникальных песен за историю рейтингов"
    overall_chart_ws["A1"].font = Font(size=16, bold=True, color="1F4E78")

    chart_positions = ["A3", "J3", "A21", "J21"]
    overall_positions = ["A3", "J3", "A21", "J21"]
    data_row = 1

    for index, dimension in enumerate(DIMENSIONS):
        yearly_dimension = [item for item in items if item.year is not None and item.dimension == dimension]
        totals: dict[str, int] = defaultdict(int)
        for item in yearly_dimension:
            totals[item.category] += item.songs_count
        categories = [name for name, _ in sorted(totals.items(), key=lambda x: (-x[1], x[0].casefold()))]
        # Графики не перегружаем: полная детализация остаётся в таблице.
        chart_categories = categories[:8]
        if len(categories) > 8:
            chart_categories.append("Другие")

        data_ws.cell(data_row, 1, f"{dimension} — по годам")
        header_row = data_row + 1
        data_ws.cell(header_row, 1, "Категория")
        for col, year in enumerate(years, start=2):
            data_ws.cell(header_row, col, year)
        for offset, category in enumerate(chart_categories, start=1):
            row_num = header_row + offset
            data_ws.cell(row_num, 1, category)
            for col, year in enumerate(years, start=2):
                if category == "Другие":
                    count = sum(
                        item.songs_count for item in yearly_dimension
                        if item.year == year and item.category not in chart_categories[:-1]
                    )
                else:
                    count = sum(
                        item.songs_count for item in yearly_dimension
                        if item.year == year and item.category == category
                    )
                data_ws.cell(row_num, col, count)
        end_row = header_row + len(chart_categories)

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.grouping = "stacked"
        chart.overlap = 100
        chart.title = f"{dimension}: изменение по годам"
        chart.y_axis.title = "Количество песен"
        chart.x_axis.title = "Год рейтинга"
        chart.height = 9
        chart.width = 17
        chart.legend.position = "b"
        chart.add_data(
            Reference(data_ws, min_col=1, max_col=1 + len(years), min_row=header_row + 1, max_row=end_row),
            titles_from_data=True,
            from_rows=True,
        )
        chart.set_categories(Reference(data_ws, min_col=2, max_col=1 + len(years), min_row=header_row, max_row=header_row))
        yearly_chart_ws.add_chart(chart, chart_positions[index])

        overall_dimension = [item for item in items if item.year is None and item.dimension == dimension]
        overall_dimension.sort(key=lambda x: (-x.songs_count, x.category.casefold()))
        # Для общей структуры оставляем до 12 категорий; прочие объединяем.
        overall_top = overall_dimension[:12]
        overall_categories = [(item.category, item.songs_count) for item in overall_top]
        if len(overall_dimension) > 12:
            overall_categories.append(("Другие", sum(x.songs_count for x in overall_dimension[12:])))
        overall_header = data_row + 2 + len(chart_categories)
        data_ws.cell(overall_header, 1, f"{dimension} — общая структура")
        data_ws.cell(overall_header + 1, 1, "Категория")
        data_ws.cell(overall_header + 1, 2, "Количество песен")
        for offset, (category, count) in enumerate(overall_categories, start=1):
            data_ws.cell(overall_header + 1 + offset, 1, category)
            data_ws.cell(overall_header + 1 + offset, 2, count)
        overall_end = overall_header + 1 + len(overall_categories)

        overall_chart = BarChart()
        overall_chart.type = "bar"
        overall_chart.style = 11
        overall_chart.title = f"{dimension}: общая структура"
        overall_chart.x_axis.title = "Количество песен"
        overall_chart.y_axis.title = "Категория"
        overall_chart.height = 9
        overall_chart.width = 17
        overall_chart.legend = None
        overall_chart.add_data(
            Reference(data_ws, min_col=2, max_col=2, min_row=overall_header + 1, max_row=overall_end),
            titles_from_data=True,
        )
        overall_chart.set_categories(
            Reference(data_ws, min_col=1, max_col=1, min_row=overall_header + 2, max_row=overall_end)
        )
        overall_chart_ws.add_chart(overall_chart, overall_positions[index])
        data_row = overall_end + 2


def generate_excel_report(
    db_path: str | Path | None = None,
    output_path: str | Path | None = None,
    hall_of_fame_threshold: int = 3,
    interval_count: int = 7,
    yearly_output_path: str | Path | None = None,
) -> Path:
    """Создаёт общую и годовую книги. Функция только читает SQLite-базу."""
    database_path = Path(db_path) if db_path is not None else DATABASE_PATH
    destination = Path(output_path) if output_path is not None else DEFAULT_REPORT_PATH
    yearly_destination = (
        Path(yearly_output_path) if yearly_output_path is not None
        else destination.with_name("music_yearly_analytics.xlsx")
    )

    songs = get_song_statistics(database_path)
    artists = get_artist_statistics(database_path)
    artist_stability = get_artist_stability_statistics(database_path)
    hall_of_fame = get_hall_of_fame(hall_of_fame_threshold, database_path)
    yearly = get_yearly_statistics(database_path)
    composition = get_composition_statistics(database_path, interval_count=interval_count)
    years = [item.year for item in yearly]

    workbook = Workbook()
    workbook.remove(workbook.active)

    song_headers = [
        "Песня", "Исполнители", "Год выпуска", "Попаданий", "Лучшая позиция",
        "Средняя позиция", "Средняя позиция с учётом отсутствий", "Баллы",
    ]
    _write_sheet(
        workbook, "Статистика песен", song_headers, _song_summary_rows(songs),
        widths={1: 30, 2: 34, 7: 30}, numeric_columns=set(range(3, len(song_headers) + 1)),
    )

    artist_headers = ["Исполнитель", "Попаданий", "Лучшая позиция", "Средняя позиция", "Баллы"]
    _write_sheet(
        workbook, "Статистика исполнителей", artist_headers, _artist_summary_rows(artists),
        widths={1: 32}, numeric_columns=set(range(2, len(artist_headers) + 1)),
    )

    stability_headers = [
        "Исполнитель", "Рейтингов с присутствием", "Всего рейтингов", "Доля присутствия",
        "Самая длинная серия рейтингов подряд", "Текущая серия рейтингов подряд",
        "Максимум песен за один рейтинг", "Среднее песен за рейтинг с присутствием",
        "Самый результативный трек", "Баллы лучшего трека", "Доля баллов лучшего трека",
    ]
    _write_sheet(
        workbook, "Устойчивость исполнителей", stability_headers,
        _artist_stability_rows(artist_stability), widths={1: 30, 8: 28, 9: 34},
        numeric_columns={2, 3, 4, 5, 6, 7, 8, 10, 11},
    )
    stability_ws = workbook["Устойчивость исполнителей"]
    for cell in stability_ws["D"][1:]:
        cell.number_format = "0.0%"
    for cell in stability_ws["K"][1:]:
        cell.number_format = "0.0%"

    _write_sheet(
        workbook, "Зал славы",
        ["Песня", "Исполнители", "Год выпуска", "Попаданий", "Год вступления", "Лучшая позиция", "Средняя позиция", "Баллы"],
        _hall_of_fame_rows(hall_of_fame), widths={1: 30, 2: 34}, numeric_columns=set(range(3, 9)),
    )

    _write_sheet(
        workbook, "Статистика по годам",
        ["Год рейтинга", "Размер рейтинга", "Записей", "Новых песен", "Повторных песен", "Доля повторов", "Средняя позиция"],
        _year_rows(yearly), widths={1: 16, 6: 18}, numeric_columns={1, 2, 3, 4, 5, 6, 7},
    )
    for cell in workbook["Статистика по годам"]["F"][1:]:
        cell.number_format = "0.0%"

    _add_composition_sheets(workbook, composition, years)

    # Пояснения к трактовке дельт и многозначных категорий.
    notes = workbook.create_sheet("Методика")
    notes.append(["Показатель", "Правило расчёта"])
    notes.append(["Доля категории", "Количество уникальных песен в категории / количество уникальных песен в соответствующем рейтинге или всей истории."])
    notes.append(["Страны и языки", "Песня учитывается во всех указанных для неё странах и языках. Поэтому суммы категорий и долей могут превышать размер рейтинга и 100%."])
    notes.append(["Десятилетия", "Год выпуска округляется вниз до начала десятилетия: 1987 год относится к 1980-м."])
    notes.append(["Равные интервалы", f"Используется {interval_count} общих интервалов по диапазону лет выпуска всех песен. Границы одинаковы для всех годовых рейтингов."])
    notes.append(["Δ позиции песни", "Позиция предыдущего рейтинга минус позиция текущего. Положительное значение — песня поднялась. Сравнение есть только если песня присутствовала в обоих соседних рейтингах."])
    notes.append(["Годовая аналитика", "Подробная аналитика вынесена в отдельную книгу: для каждого года есть листы песен и исполнителей. Общие баллы на конец года — накопленная сумма с начала истории; Δ общих баллов — разница накопленных итогов текущего и предыдущего рейтингового года."])
    notes.append(["Позиция исполнителя", "Для каждого года показываются лучшая и средняя позиции среди песен исполнителя; совместные песни засчитываются каждому исполнителю."])
    notes.append(["Δ позиции исполнителя", "Для лучшей и средней позиции: значение прошлого рейтинга минус значение текущего. Положительное значение — улучшение. Считается только если исполнитель присутствовал в обоих соседних рейтингах."])
    notes.append(["Статусы NEW, RE и OUT", "NEW — первое появление песни или исполнителя в истории рейтингов. RE — возвращение после отсутствия в предыдущем рейтинге. OUT — ранее появлялся, но отсутствует в рейтинге текущего года. Пустой статус означает присутствие и в предыдущем, и в текущем рейтингах."])
    notes.append(["Накопительная годовая аналитика", "Каждый годовой лист включает все песни и исполнителей, появлявшихся в любом рейтинге с начала истории по выбранный год, даже если они отсутствуют в текущем рейтинге. Строки сортируются по общим накопленным баллам на конец выбранного года по убыванию; при равенстве используется алфавитная сортировка."])
    notes.append(["Устойчивость исполнителя", "Доля присутствия — число рейтингов с исполнителем / общее число рейтингов. Серия подряд считается по последовательности имеющихся рейтингов, а не по календарным годам. Максимум песен и среднее число песен показывают ширину каталога в рейтингах."])
    notes.append(["Концентрация баллов", "Доля баллов самого результативного трека — баллы этого трека / все баллы исполнителя. Чем выше доля, тем сильнее общий результат зависит от одной песни."])
    notes.append(["Сортировка лет", "Сравнение выполняется между соседними годами, для которых существуют рейтинги, даже если календарный год между ними пропущен."])
    for cell in notes[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in notes.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    notes.column_dimensions["A"].width = 30
    notes.column_dimensions["B"].width = 110
    notes.freeze_panes = "A2"
    notes.sheet_view.showGridLines = False
    for row in range(2, notes.max_row + 1):
        notes.row_dimensions[row].height = 42

    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    _write_yearly_report(songs, artists, years, yearly_destination)
    return destination


if __name__ == "__main__":
    report_path = generate_excel_report()
    print(f"Excel-отчёт сохранён: {report_path}")
