# GeoTrace Forensics

GeoTrace Forensics is a desktop forensic analysis tool for digital image metadata extraction and geolocation investigation. It extracts EXIF metadata, decodes and validates GPS coordinates, builds a geolocation timeline, flags metadata anomalies, and produces map and report artifacts for case work.

## Implemented Features

- EXIF extraction using `exifread` with Pillow fallback for supported image formats
- GPS decoding, validation, and investigator-friendly coordinate formatting
- Interactive folium map with marker popups, movement path, and timestamp slider
- Timeline and proximity correlation analysis across multiple images
- Metadata anomaly detection for missing/invalid EXIF, suspicious timestamps, editing software, and post-capture modification
- Evidence chain storage in SQLite with SHA-256 file hashing
- PDF and Excel/JSON forensic exports
- PyQt5 desktop GUI for case creation, image intake, review, and artifact generation

## Run

1. Create or activate a Python environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Launch the application from the `GeoTrace_Forensics` directory:

```bash
python main.py
```

## Outputs

Generated artifacts are stored in `output/case_<id>/` and include:

- Interactive HTML map
- PDF forensic report
- Excel export
- JSON export

## Database

The SQLite database is stored at `database/geotrace.db` and keeps:

- Cases
- Imported image evidence
- Extracted metadata
- Detected anomalies
