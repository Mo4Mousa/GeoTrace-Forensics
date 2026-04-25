from pathlib import Path
import shutil
from datetime import datetime
import os

from PIL import Image, ImageDraw, ImageFont, ExifTags, TiffImagePlugin


ROOT = Path(__file__).resolve().parent
TAGS = {value: key for key, value in ExifTags.TAGS.items()}
GPS_TAGS = {value: key for key, value in ExifTags.GPSTAGS.items()}
RATIONAL = TiffImagePlugin.IFDRational


def _load_font(size):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
    ]
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def _to_dms(decimal_value):
    absolute = abs(decimal_value)
    degrees = int(absolute)
    minutes_full = (absolute - degrees) * 60
    minutes = int(minutes_full)
    seconds = round((minutes_full - minutes) * 60, 2)
    return (
        RATIONAL(degrees, 1),
        RATIONAL(minutes, 1),
        RATIONAL(int(seconds * 100), 100),
    )


def _build_exif(make, model, taken_at, software=None, latitude=None, longitude=None):
    exif = Image.Exif()
    exif[TAGS["Make"]] = make
    exif[TAGS["Model"]] = model
    exif[TAGS["DateTimeOriginal"]] = taken_at

    if software:
        exif[TAGS["Software"]] = software

    if latitude is not None and longitude is not None:
        exif[TAGS["GPSInfo"]] = {
            GPS_TAGS["GPSLatitudeRef"]: "N" if latitude >= 0 else "S",
            GPS_TAGS["GPSLatitude"]: _to_dms(latitude),
            GPS_TAGS["GPSLongitudeRef"]: "E" if longitude >= 0 else "W",
            GPS_TAGS["GPSLongitude"]: _to_dms(longitude),
        }

    return exif


def _make_card(filename, title, subtitle, color, exif=None):
    image = Image.new("RGB", (1600, 1000), color)
    draw = ImageDraw.Draw(image)
    title_font = _load_font(74)
    subtitle_font = _load_font(38)

    draw.rounded_rectangle((70, 70, 1530, 930), radius=30, outline="white", width=6)
    draw.text((120, 180), title, fill="white", font=title_font)
    draw.text((120, 320), subtitle, fill="white", font=subtitle_font)
    draw.text((120, 390), filename, fill="white", font=subtitle_font)

    save_path = ROOT / filename
    if exif is not None:
        image.save(save_path, quality=95, exif=exif)
    else:
        image.save(save_path)
    return save_path


def _apply_capture_mtime(file_path, taken_at):
    if not taken_at:
        return

    timestamp = datetime.strptime(taken_at, "%Y:%m:%d %H:%M:%S").timestamp()
    os.utime(file_path, (timestamp, timestamp))


def generate():
    ROOT.mkdir(parents=True, exist_ok=True)

    cairo_taken_at = "2026:04:20 10:15:00"
    cairo = _make_card(
        "01_cairo_gps.jpg",
        "GeoTrace Test Image 01",
        "GPS-enabled image near Cairo",
        "#1f5f74",
        _build_exif(
            make="Canon",
            model="EOS Test 90D",
            taken_at=cairo_taken_at,
            latitude=30.0444,
            longitude=31.2357,
        ),
    )
    _apply_capture_mtime(cairo, cairo_taken_at)

    alex_taken_at = "2026:04:20 12:05:00"
    alex = _make_card(
        "02_alexandria_gps.jpg",
        "GeoTrace Test Image 02",
        "GPS-enabled image near Alexandria",
        "#2d6a4f",
        _build_exif(
            make="Nikon",
            model="D750 Test",
            taken_at=alex_taken_at,
            latitude=31.2001,
            longitude=29.9187,
        ),
    )
    _apply_capture_mtime(alex, alex_taken_at)

    shutil.copy2(cairo, ROOT / "03_duplicate_of_cairo.jpg")

    _make_card(
        "04_no_exif.png",
        "GeoTrace Test Image 04",
        "No EXIF metadata embedded",
        "#6d597a",
        exif=None,
    )

    edited_taken_at = "2026:04:20 14:30:00"
    edited = _make_card(
        "05_edited_software.jpg",
        "GeoTrace Test Image 05",
        "Contains editing software metadata",
        "#9c6644",
        _build_exif(
            make="Sony",
            model="Alpha Test A7",
            taken_at=edited_taken_at,
            software="Adobe Photoshop 2025",
            latitude=29.9765,
            longitude=31.1313,
        ),
    )
    _apply_capture_mtime(edited, edited_taken_at)


if __name__ == "__main__":
    generate()
