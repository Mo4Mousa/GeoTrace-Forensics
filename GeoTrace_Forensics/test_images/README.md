# Test Images

These sample files are ready to use with the GeoTrace Forensics desktop app.

## Included Files

- `01_cairo_gps.jpg`: contains EXIF camera info, timestamp, and GPS coordinates near Cairo.
- `02_alexandria_gps.jpg`: contains EXIF camera info, timestamp, and GPS coordinates near Alexandria.
- `03_duplicate_of_cairo.jpg`: exact duplicate of `01_cairo_gps.jpg` to test duplicate-hash detection.
- `04_no_exif.png`: image without EXIF metadata to test missing metadata behavior.
- `05_edited_software.jpg`: image with GPS plus `Software=Adobe Photoshop 2025` to trigger editing-software anomaly detection.

## Usage

Open the app, create/select a case, then click `Add Images` and import files from this folder.

## Regeneration

If you ever delete the files, regenerate them with:

```bash
GeoTrace_Forensics/venv/bin/python GeoTrace_Forensics/test_images/generate_test_images.py
```
