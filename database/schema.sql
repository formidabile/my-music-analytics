PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS countries (
    country_id INTEGER PRIMARY KEY,
    country_name TEXT NOT NULL UNIQUE CHECK (length(trim(country_name)) > 0)
);

CREATE TABLE IF NOT EXISTS languages (
    language_id INTEGER PRIMARY KEY,
    language_name TEXT NOT NULL UNIQUE CHECK (length(trim(language_name)) > 0)
);

CREATE TABLE IF NOT EXISTS artists (
    artist_id INTEGER PRIMARY KEY,
    artist_name TEXT NOT NULL UNIQUE CHECK (length(trim(artist_name)) > 0)
);

CREATE TABLE IF NOT EXISTS songs (
    song_id INTEGER PRIMARY KEY,
    title TEXT NOT NULL CHECK (length(trim(title)) > 0),
    release_year INTEGER NOT NULL CHECK (release_year BETWEEN 1900 AND 2100)
);

CREATE TABLE IF NOT EXISTS song_artists (
    song_id INTEGER NOT NULL,
    artist_id INTEGER NOT NULL,
    PRIMARY KEY (song_id, artist_id),
    FOREIGN KEY (song_id) REFERENCES songs(song_id) ON DELETE CASCADE,
    FOREIGN KEY (artist_id) REFERENCES artists(artist_id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS song_countries (
    song_id INTEGER NOT NULL,
    country_id INTEGER NOT NULL,
    PRIMARY KEY (song_id, country_id),
    FOREIGN KEY (song_id) REFERENCES songs(song_id) ON DELETE CASCADE,
    FOREIGN KEY (country_id) REFERENCES countries(country_id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS song_languages (
    song_id INTEGER NOT NULL,
    language_id INTEGER NOT NULL,
    PRIMARY KEY (song_id, language_id),
    FOREIGN KEY (song_id) REFERENCES songs(song_id) ON DELETE CASCADE,
    FOREIGN KEY (language_id) REFERENCES languages(language_id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS rankings (
    ranking_id INTEGER PRIMARY KEY,
    ranking_year INTEGER NOT NULL UNIQUE CHECK (ranking_year BETWEEN 1900 AND 2100),
    top_size INTEGER NOT NULL DEFAULT 50 CHECK (top_size > 0)
);

CREATE TABLE IF NOT EXISTS ranking_entries (
    ranking_id INTEGER NOT NULL,
    song_id INTEGER NOT NULL,
    position INTEGER NOT NULL CHECK (position > 0),
    PRIMARY KEY (ranking_id, song_id),
    UNIQUE (ranking_id, position),
    FOREIGN KEY (ranking_id) REFERENCES rankings(ranking_id) ON DELETE CASCADE,
    FOREIGN KEY (song_id) REFERENCES songs(song_id) ON DELETE RESTRICT
);

CREATE INDEX IF NOT EXISTS idx_songs_title_year
    ON songs(title, release_year);
CREATE INDEX IF NOT EXISTS idx_ranking_entries_song
    ON ranking_entries(song_id);

CREATE TRIGGER IF NOT EXISTS trg_ranking_entries_position_insert
BEFORE INSERT ON ranking_entries
FOR EACH ROW
WHEN NEW.position > (
    SELECT top_size FROM rankings WHERE ranking_id = NEW.ranking_id
)
BEGIN
    SELECT RAISE(ABORT, 'position exceeds ranking top_size');
END;

CREATE TRIGGER IF NOT EXISTS trg_ranking_entries_position_update
BEFORE UPDATE OF position, ranking_id ON ranking_entries
FOR EACH ROW
WHEN NEW.position > (
    SELECT top_size FROM rankings WHERE ranking_id = NEW.ranking_id
)
BEGIN
    SELECT RAISE(ABORT, 'position exceeds ranking top_size');
END;

-- Прогнозы хранятся отдельно от фактических рейтингов.
CREATE TABLE IF NOT EXISTS forecast_years (
    forecast_year INTEGER PRIMARY KEY CHECK (forecast_year BETWEEN 1900 AND 2100)
);

CREATE TABLE IF NOT EXISTS forecast_entries (
    forecast_entry_id INTEGER PRIMARY KEY,
    forecast_year INTEGER NOT NULL,
    predicted_position INTEGER NOT NULL CHECK (predicted_position > 0),
    title TEXT NOT NULL,
    release_year INTEGER NOT NULL CHECK (release_year BETWEEN 1900 AND 2100),
    artists_json TEXT NOT NULL,
    languages_json TEXT NOT NULL,
    countries_json TEXT NOT NULL,
    identity_key TEXT NOT NULL,
    UNIQUE (forecast_year, predicted_position),
    UNIQUE (forecast_year, identity_key),
    FOREIGN KEY (forecast_year) REFERENCES forecast_years(forecast_year) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_forecast_entries_identity
    ON forecast_entries(identity_key);


-- Стриминговый Top-5 использует общий справочник исполнителей.
CREATE TABLE IF NOT EXISTS streaming_entries (
    year INTEGER NOT NULL CHECK (year BETWEEN 1900 AND 2100),
    artist_id INTEGER NOT NULL,
    position INTEGER NOT NULL CHECK (position BETWEEN 1 AND 5),
    listening_minutes INTEGER CHECK (listening_minutes IS NULL OR listening_minutes >= 0),
    PRIMARY KEY (year, position),
    UNIQUE (year, artist_id),
    FOREIGN KEY (artist_id) REFERENCES artists(artist_id) ON DELETE RESTRICT
);
CREATE INDEX IF NOT EXISTS idx_streaming_entries_artist
    ON streaming_entries(artist_id);
