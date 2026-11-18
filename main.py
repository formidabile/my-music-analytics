"""Точка входа приложения: импорт фактических рейтингов и прогнозов."""
from src.importer import main as import_project_data
from src.forecasts import import_forecasts


def main() -> int:
    result = import_project_data()
    if result != 0:
        return result
    try:
        count = import_forecasts()
    except FileNotFoundError:
        print("Книга прогнозов не найдена; импорт прогнозов пропущен.")
    except (OSError, ValueError) as exc:
        print(f"Ошибка импорта прогнозов: {exc}")
        return 5
    else:
        print(f"Импортировано прогнозных записей: {count}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
