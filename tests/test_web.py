from fastapi.testclient import TestClient

from web.main import app

client = TestClient(app)


def test_all_main_pages_render():
    for path, heading in [
        ("/", "Главная"),
        ("/rankings", "Рейтинги"),
        ("/analytics", "Аналитика"),
        ("/streaming", "Стриминги"),
        ("/forecasts", "Прогнозы"),
    ]:
        response = client.get(path)
        assert response.status_code == 200
        assert heading in response.text
        assert 'data-page=' in response.text


def test_unknown_page_returns_404():
    response = client.get("/unknown-page")
    assert response.status_code == 404


def test_static_stylesheet_is_served():
    response = client.get("/static/css/style.css")
    assert response.status_code == 200
    assert "--bg" in response.text


def test_overview_endpoint_is_available_without_database():
    response = client.get("/api/overview")
    assert response.status_code == 200
    assert "database_ready" in response.json()
