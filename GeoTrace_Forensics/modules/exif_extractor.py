import json
from pathlib import Path

import exifread
from PIL import ExifTags, Image

from modules.gps_decoder import decode_gps


EXIF_DATE_KEYS = [
    "EXIF DateTimeOriginal",
    "EXIF DateTimeDigitized",
    "Image DateTime",
    "DateTimeOriginal",
    "DateTimeDigitized",
    "DateTime",
]


def _clean_text(value):
    if value is None:
        return None

    text = str(value).strip()
    return text or None


def _pillow_tag_name(tag_id):
    return ExifTags.TAGS.get(tag_id, str(tag_id))


def _normalize_pillow_exif(image):
    normalized = {}

    try:
        raw_exif = image.getexif()
    except Exception:
        raw_exif = {}

    for tag_id, value in raw_exif.items():
        tag_name = _pillow_tag_name(tag_id)
        if tag_name == "GPSInfo" and isinstance(value, dict):
            for gps_key, gps_value in value.items():
                gps_name = ExifTags.GPSTAGS.get(gps_key, str(gps_key))
                normalized[gps_name] = gps_value
        else:
            normalized[tag_name] = value

    return normalized


def _extract_with_exifread(file_path):
    with open(file_path, "rb") as image_file:
        return exifread.process_file(image_file, details=False)


def extract_exif_data(file_path):
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    metadata = {
        "camera_make": None,
        "camera_model": None,
        "date_taken": None,
        "latitude": None,
        "longitude": None,
        "software": None,
        "image_width": None,
        "image_height": None,
        "file_format": None,
        "raw_exif": "{}",
        "has_exif": False,
    }

    exifread_tags = {}
    pillow_tags = {}

    try:
        with Image.open(file_path) as image:
            metadata["image_width"], metadata["image_height"] = image.size
            metadata["file_format"] = image.format
            pillow_tags = _normalize_pillow_exif(image)
    except Exception:
        pillow_tags = {}

    try:
        exifread_tags = _extract_with_exifread(file_path)
    except Exception:
        exifread_tags = {}

    merged_tags = {}
    merged_tags.update(pillow_tags)
    merged_tags.update({str(key): value for key, value in exifread_tags.items()})

    if not merged_tags:
        return metadata

    metadata["has_exif"] = True
    metadata["camera_make"] = _clean_text(
        merged_tags.get("Image Make") or merged_tags.get("Make")
    )
    metadata["camera_model"] = _clean_text(
        merged_tags.get("Image Model") or merged_tags.get("Model")
    )
    metadata["software"] = _clean_text(
        merged_tags.get("Image Software") or merged_tags.get("Software")
    )

    for key in EXIF_DATE_KEYS:
        if _clean_text(merged_tags.get(key)):
            metadata["date_taken"] = _clean_text(merged_tags.get(key))
            break

    latitude, longitude = decode_gps(merged_tags)
    metadata["latitude"] = latitude
    metadata["longitude"] = longitude

    serializable_tags = {}
    for key, value in merged_tags.items():
        serializable_tags[str(key)] = str(value)

    metadata["raw_exif"] = json.dumps(serializable_tags, ensure_ascii=False, indent=2)
    return metadata
