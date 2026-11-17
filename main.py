"""Точка входа приложения: подготовка SQLite и импорт новых рейтингов."""
from src.importer import main as import_project_data


def main() -> int:
    return import_project_data()


if __name__ == "__main__":
    raise SystemExit(main())
