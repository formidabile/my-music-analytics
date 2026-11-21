from __future__ import annotations

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pathlib import Path

from web.api import data

WEB_DIR = Path(__file__).resolve().parent
app = FastAPI(title="Music Statistics Dashboard", version="1.0.0")
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")
templates = Jinja2Templates(directory=WEB_DIR / "templates")

PAGES = {
    "overview": ("Главная", "Общий обзор музыкальной статистики"),
    "rankings": ("Рейтинги", "Годовые рейтинги Top-50"),
    "analytics": ("Аналитика", "Статистика песен и исполнителей"),
    "streaming": ("Стриминги", "Сравнение рейтингов и времени прослушивания"),
    "forecasts": ("Прогнозы", "Сопоставление прогнозов с результатами"),
}


@app.get("/", response_class=HTMLResponse)
@app.get("/{page}", response_class=HTMLResponse)
def render_page(request: Request, page: str = "overview"):
    if page not in PAGES:
        raise HTTPException(status_code=404, detail="Страница не найдена")
    title, description = PAGES[page]
    return templates.TemplateResponse(
        request=request,
        name="page.html",
        context={"page": page, "title": title, "description": description, "pages": PAGES},
    )


@app.get("/songs/{song_id}", response_class=HTMLResponse)
def song_page(request: Request, song_id: int):
    item = data.get_song_detail(song_id)
    if item is None:
        raise HTTPException(status_code=404, detail="Песня не найдена")
    return templates.TemplateResponse(
        request=request, name="detail.html",
        context={"entity_type": "song", "item": item, "title": item["title"]},
    )


@app.get("/artists/{artist_id}", response_class=HTMLResponse)
def artist_page(request: Request, artist_id: int):
    item = data.get_artist_detail(artist_id)
    if item is None:
        raise HTTPException(status_code=404, detail="Исполнитель не найден")
    return templates.TemplateResponse(
        request=request, name="detail.html",
        context={"entity_type": "artist", "item": item, "title": item["artist"]},
    )


@app.get("/api/overview")
def overview():
    return data.get_overview()


@app.get("/api/years")
def years():
    return {"years": data.get_years()}


@app.get("/api/rankings/{year}")
def ranking(year: int):
    result = data.get_ranking(year)
    if not result:
        raise HTTPException(status_code=404, detail=f"Рейтинг за {year} год не найден")
    return {"year": year, "entries": result, "composition": data.get_ranking_composition(year)}


@app.get("/api/songs/{song_id}")
def song(song_id: int):
    result = data.get_song_detail(song_id)
    if result is None:
        raise HTTPException(status_code=404, detail="Песня не найдена")
    return result


@app.get("/api/artists/{artist_id}")
def artist(artist_id: int):
    result = data.get_artist_detail(artist_id)
    if result is None:
        raise HTTPException(status_code=404, detail="Исполнитель не найден")
    return result


@app.get("/api/analytics")
def analytics():
    return data.get_analytics_payload()


@app.get("/api/streaming")
def streaming():
    return data.get_streaming_payload()


@app.get("/api/forecasts")
def forecasts():
    return data.get_forecasts_payload()
