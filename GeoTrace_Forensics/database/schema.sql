CREATE TABLE IF NOT EXISTS cases (
    case_id INTEGER PRIMARY KEY AUTOINCREMENT,
    case_name TEXT NOT NULL,
    investigator_name TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS images (
    image_id INTEGER PRIMARY KEY AUTOINCREMENT,
    case_id INTEGER NOT NULL,
    file_name TEXT NOT NULL,
    file_path TEXT NOT NULL,
    file_size TEXT,
    sha256_hash TEXT,
    uploaded_at TEXT NOT NULL,
    FOREIGN KEY(case_id) REFERENCES cases(case_id)
);

CREATE TABLE IF NOT EXISTS metadata (
    metadata_id INTEGER PRIMARY KEY AUTOINCREMENT,
    image_id INTEGER NOT NULL,
    camera_make TEXT,
    camera_model TEXT,
    date_taken TEXT,
    latitude REAL,
    longitude REAL,
    software TEXT,
    image_width TEXT,
    image_height TEXT,
    raw_exif TEXT,
    FOREIGN KEY(image_id) REFERENCES images(image_id)
);

CREATE TABLE IF NOT EXISTS anomalies (
    anomaly_id INTEGER PRIMARY KEY AUTOINCREMENT,
    image_id INTEGER NOT NULL,
    anomaly_type TEXT,
    severity TEXT,
    description TEXT,
    FOREIGN KEY(image_id) REFERENCES images(image_id)
);