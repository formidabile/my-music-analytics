"""Запуск локального веб-интерфейса Music Statistics."""

import uvicorn


def main() -> None:
    uvicorn.run("web.main:app", host="127.0.0.1", port=8000, reload=False)


if __name__ == "__main__":
    main()
