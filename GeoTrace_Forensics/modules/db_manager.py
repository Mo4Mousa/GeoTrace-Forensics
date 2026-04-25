import sqlite3
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent


class DatabaseManager:
    def __init__(self, db_path=None, schema_path=None):
        self.db_path = self._resolve_path(db_path or BASE_DIR / "database/geotrace.db")
        self.schema_path = self._resolve_path(schema_path or BASE_DIR / "database/schema.sql")

        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize_database()

    def _resolve_path(self, path_value):
        path = Path(path_value)
        return path if path.is_absolute() else (BASE_DIR / path).resolve()

    def connect(self):
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def initialize_database(self):
        if not self.schema_path.exists():
            raise FileNotFoundError(f"Schema file not found: {self.schema_path}")

        with self.connect() as connection:
            connection.executescript(self.schema_path.read_text(encoding="utf-8"))

    def create_case(self, case_name, investigator_name):
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO cases (case_name, investigator_name, created_at)
                VALUES (?, ?, ?)
                """,
                (case_name, investigator_name, created_at),
            )
            connection.commit()
            return cursor.lastrowid

    def get_case(self, case_id):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT case_id, case_name, investigator_name, created_at
                FROM cases
                WHERE case_id = ?
                """,
                (case_id,),
            )
            row = cursor.fetchone()
            return dict(row) if row else None

    def get_all_cases(self):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT case_id, case_name, investigator_name, created_at
                FROM cases
                ORDER BY case_id DESC
                """
            )
            return [dict(row) for row in cursor.fetchall()]

    def insert_image(self, case_id, file_name, file_path, file_size, sha256_hash):
        uploaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO images
                (case_id, file_name, file_path, file_size, sha256_hash, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (case_id, file_name, file_path, file_size, sha256_hash, uploaded_at),
            )
            connection.commit()
            return cursor.lastrowid

    def insert_metadata(
        self,
        image_id,
        camera_make=None,
        camera_model=None,
        date_taken=None,
        latitude=None,
        longitude=None,
        software=None,
        image_width=None,
        image_height=None,
        raw_exif=None,
    ):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO metadata
                (image_id, camera_make, camera_model, date_taken, latitude, longitude,
                 software, image_width, image_height, raw_exif)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    image_id,
                    camera_make,
                    camera_model,
                    date_taken,
                    latitude,
                    longitude,
                    software,
                    image_width,
                    image_height,
                    raw_exif,
                ),
            )
            connection.commit()
            return cursor.lastrowid

    def insert_anomaly(self, image_id, anomaly_type, severity, description):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO anomalies
                (image_id, anomaly_type, severity, description)
                VALUES (?, ?, ?, ?)
                """,
                (image_id, anomaly_type, severity, description),
            )
            connection.commit()
            return cursor.lastrowid

    def get_anomalies_for_image(self, image_id):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT anomaly_type, severity, description
                FROM anomalies
                WHERE image_id = ?
                ORDER BY anomaly_id ASC
                """,
                (image_id,),
            )
            return [dict(row) for row in cursor.fetchall()]

    def find_images_by_hash(self, sha256_hash, exclude_image_id=None):
        if not sha256_hash:
            return []

        query = [
            """
            SELECT image_id, case_id, file_name, file_path, sha256_hash
            FROM images
            WHERE sha256_hash = ?
            """
        ]
        parameters = [sha256_hash]

        if exclude_image_id is not None:
            query.append("AND image_id != ?")
            parameters.append(exclude_image_id)

        query.append("ORDER BY image_id ASC")

        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute("\n".join(query), tuple(parameters))
            return [dict(row) for row in cursor.fetchall()]

    def get_case_results(self, case_id):
        with self.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                SELECT
                    i.image_id,
                    i.file_name,
                    i.file_path,
                    i.file_size,
                    i.sha256_hash,
                    i.uploaded_at,
                    m.camera_make,
                    m.camera_model,
                    m.date_taken,
                    m.latitude,
                    m.longitude,
                    m.software,
                    m.image_width,
                    m.image_height,
                    m.raw_exif
                FROM images i
                LEFT JOIN metadata m ON i.image_id = m.image_id
                WHERE i.case_id = ?
                ORDER BY
                    CASE WHEN m.date_taken IS NULL THEN 1 ELSE 0 END,
                    m.date_taken ASC,
                    i.uploaded_at ASC
                """,
                (case_id,),
            )
            results = [dict(row) for row in cursor.fetchall()]

        for row in results:
            row["anomalies"] = self.get_anomalies_for_image(row["image_id"])

        return results

    def get_case_artifact_dir(self, case_id):
        artifact_dir = BASE_DIR / "output" / f"case_{case_id}"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        return artifact_dir
