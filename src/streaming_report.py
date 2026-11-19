"""Сравнение ежегодного стримингового Top-5 с музыкальным Top-50."""
from __future__ import annotations

from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .analytics import get_artist_statistics
from .database import BASE_DIR, DATABASE_PATH, get_connection

DEFAULT_REPORT_PATH = BASE_DIR / "reports" / "music_streaming_analytics.xlsx"
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10, color="202124")
ALT_FILL = PatternFill(fill_type="solid", fgColor="F3F6F9")


def _write_sheet(workbook: Workbook, title: str, headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    ws = workbook.create_sheet(title)
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 38
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


def generate_streaming_report(
    db_path: str | Path | None = None,
    output_path: str | Path | None = None,
) -> Path:
    """Формирует годовое сопоставление стримингового Top-5 и артистов Top-50."""
    database_path = Path(db_path) if db_path is not None else DATABASE_PATH
    destination = Path(output_path) if output_path is not None else DEFAULT_REPORT_PATH
    with get_connection(database_path) as connection:
        streaming_rows = connection.execute("""
            SELECT se.year, se.position, a.artist_name, se.listening_minutes
            FROM streaming_entries se JOIN artists a USING (artist_id)
            ORDER BY se.year, se.position
        """).fetchall()
        ranking_years = {row[0] for row in connection.execute("SELECT ranking_year FROM rankings")}

    artist_stats = get_artist_statistics(database_path)
    stats_by_year: dict[int, list] = {year: [] for year in ranking_years}
    for artist in artist_stats:
        for year in ranking_years:
            if artist.appearances_by_year.get(year, 0):
                stats_by_year[year].append(artist)

    ranks: dict[tuple[int, str], int] = {}
    for year, artists in stats_by_year.items():
        artists.sort(key=lambda item: (
            -item.points_by_year.get(year, 0),
            -item.appearances_by_year.get(year, 0),
            item.best_positions_by_year.get(year) or 999,
            item.artist.casefold(),
        ))
        for rank, artist in enumerate(artists, start=1):
            ranks[(year, artist.artist.casefold())] = rank

    rows = []
    cumulative: dict[str, dict[str, object]] = {}
    for item in streaming_rows:
        year = item["year"]
        artist_name = item["artist_name"]
        artist = next((a for a in artist_stats if a.artist.casefold() == artist_name.casefold()), None)
        appearances = artist.appearances_by_year.get(year, 0) if artist else 0
        best_position = artist.best_positions_by_year.get(year) if artist else None
        points = artist.points_by_year.get(year, 0) if artist else 0
        artist_rank = ranks.get((year, artist_name.casefold()))
        rank_difference = artist_rank - item["position"] if artist_rank is not None else None
        rows.append((
            year, item["position"], artist_name, item["listening_minutes"],
            artist_rank, rank_difference, appearances or 0, best_position, points,
            "Да" if appearances else "Нет",
        ))

        key = artist_name.casefold()
        aggregate = cumulative.setdefault(key, {
            "artist": artist_name, "years": [], "positions": [],
            "known_minutes": [], "latest_year": year,
            "latest_position": item["position"],
            "latest_minutes": item["listening_minutes"],
        })
        aggregate["years"].append(year)
        aggregate["positions"].append(item["position"])
        if item["listening_minutes"] is not None:
            aggregate["known_minutes"].append(item["listening_minutes"])
        if year >= aggregate["latest_year"]:
            aggregate["latest_year"] = year
            aggregate["latest_position"] = item["position"]
            aggregate["latest_minutes"] = item["listening_minutes"]

    all_time_rows = []
    artist_by_name = {artist.artist.casefold(): artist for artist in artist_stats}
    for aggregate in cumulative.values():
        artist = artist_by_name.get(aggregate["artist"].casefold())
        known_minutes = aggregate["known_minutes"]
        positions = aggregate["positions"]
        all_time_rows.append((
            aggregate["artist"], len(aggregate["years"]),
            min(aggregate["years"]), max(aggregate["years"]),
            sum(positions) / len(positions), min(positions), max(positions),
            aggregate["latest_position"], aggregate["latest_minutes"],
            sum(known_minutes) if known_minutes else None,
            sum(known_minutes) / len(known_minutes) if known_minutes else None,
            len(known_minutes),
            artist.appearances if artist else 0,
            artist.best_position if artist else None,
            artist.total_points if artist else 0,
        ))
    # Главный критерий итогового рейтинга — суммарное известное время прослушивания.
    # Исполнители без единого значения минут идут в конец; при равенстве минут
    # выше располагается исполнитель с большим числом лет в Top-5.
    all_time_rows.sort(key=lambda row: (
        row[9] is None,
        -(row[9] or 0),
        -row[1],
        row[0].casefold(),
    ))

    wb = Workbook()
    wb.remove(wb.active)
    _write_sheet(wb, "Стриминги и Top-50", [
        "Год", "Позиция в стримингах", "Исполнитель", "Время прослушивания, мин",
        "Позиция исполнителя в Top-50*", "Разница позиций**", "Песен в Top-50",
        "Лучшая позиция песни", "Баллы исполнителя", "Есть в Top-50",
    ], rows)
    _write_sheet(wb, "За всё время", [
        "Исполнитель", "Лет в стриминговом Top-5", "Первый год", "Последний год",
        "Средняя позиция в стримингах", "Лучшая позиция", "Худшая позиция",
        "Позиция в последнем году", "Минуты в последнем году", "Всего минут (известные значения)",
        "Среднее минут за год с данными", "Лет с данными о минутах",
        "Попаданий в музыкальный Top-50 за всё время", "Лучшая позиция песни за всё время",
        "Баллы в музыкальном Top-50 за всё время",
    ], all_time_rows)
    _write_sheet(wb, "Методика", ["Показатель", "Правило расчёта"], [
        ("Стриминговая позиция", "Позиция исполнителя в исходном Top-5; определяется порядком строк в Excel."),
        ("Позиция исполнителя в Top-50", "Отдельный годовой рейтинг исполнителей по сумме баллов их песен в Top-50. При равенстве учитываются число песен, лучшая позиция песни и название исполнителя."),
        ("Разница позиций", "Позиция исполнителя в Top-50 минус позиция в стримингах. Положительное значение означает, что в Top-50 исполнитель расположился ниже, чем в стримингах."),
        ("Песен в Top-50", "Количество песен исполнителя в годовом рейтинге. Совместные песни засчитываются каждому исполнителю."),
        ("Баллы исполнителя", "Сумма баллов всех его песен в годовом Top-50; баллы песни равны 51 минус её позиция для Top-50."),
        ("Пустое время прослушивания", "Остаётся пустым и не трактуется как ноль."),
        ("Ограничение сравнения", "В годовом отчёте включены исполнители из стримингового Top-5. Позиция исполнителя в Top-50 — производный рейтинг по баллам, а не исходная позиция отдельной песни."),
        ("Статистика за всё время", "Отдельный лист агрегирует всех исполнителей, когда-либо попадавших в стриминговый Top-5, по всем годам, загруженным в базу. При добавлении нового года лист пересчитывается автоматически."),
        ("Лет в стриминговом Top-5", "Количество разных лет, в которых исполнитель входил в Top-5. В каждом году исполнитель учитывается не более одного раза."),
        ("Средняя / лучшая / худшая позиция", "Среднее арифметическое, минимум и максимум годовых позиций исполнителя в стриминговом Top-5."),
        ("Всего минут", "Сумма только известных значений времени прослушивания. Пустые ячейки исключаются, а не считаются нулём. По этому показателю лист «За всё время» отсортирован по убыванию; исполнители без известных значений находятся в конце."),
        ("Среднее минут за год с данными", "Среднее только по годам, для которых время прослушивания указано. Число таких лет вынесено в отдельный столбец."),
        ("Последний год", "Последний по хронологии год, присутствующий в загруженных стриминговых данных; при добавлении 2026 года он станет актуальным автоматически."),
        ("Музыкальный Top-50 за всё время", "Число песен, лучшая позиция и сумма баллов исполнителя по всем загруженным фактическим рейтингам. Совместные песни засчитываются каждому исполнителю."),
    ])
    # Mark methodology footnotes in the header labels explicitly.
    ws = wb["Стриминги и Top-50"]
    ws["E1"] = "Позиция исполнителя в Top-50*"
    ws["F1"] = "Разница позиций**"
    destination.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destination)
    wb.close()
    return destination


if __name__ == "__main__":
    path = generate_streaming_report()
    print(f"Отчёт по стримингам сохранён: {path}")
