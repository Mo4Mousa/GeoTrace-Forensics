from datetime import datetime
from pathlib import Path

from modules.gps_decoder import verify_coordinates


EDITING_SOFTWARE_KEYWORDS = [
    "photoshop",
    "lightroom",
    "gimp",
    "snapseed",
    "canva",
    "picsart",
    "editor",
    "adobe",
]

DATE_FORMATS = [
    "%Y:%m:%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
]


def _parse_datetime(value):
    if not value:
        return None

    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(str(value), date_format)
        except ValueError:
            continue
    return None


def detect_anomalies(file_path, metadata):
    anomalies = []
    file_path = Path(file_path)

    has_exif = metadata.get("has_exif", False)
    latitude = metadata.get("latitude")
    longitude = metadata.get("longitude")
    date_taken = metadata.get("date_taken")
    software = metadata.get("software")
    camera_make = metadata.get("camera_make")
    camera_model = metadata.get("camera_model")

    if not has_exif:
        anomalies.append(
            {
                "type": "Missing EXIF",
                "severity": "High",
                "description": "No EXIF metadata was found. Metadata may have been stripped or never embedded.",
            }
        )

    if latitude is None or longitude is None:
        anomalies.append(
            {
                "type": "Missing GPS",
                "severity": "Medium",
                "description": "GPS coordinates are missing, so the image cannot be geolocated directly.",
            }
        )
    else:
        gps_validation = verify_coordinates(latitude, longitude)
        if not gps_validation["is_valid"]:
            anomalies.append(
                {
                    "type": "Invalid GPS",
                    "severity": "High",
                    "description": " ".join(gps_validation["issues"]),
                }
            )

    parsed_date = _parse_datetime(date_taken)
    if not date_taken:
        anomalies.append(
            {
                "type": "Missing Date Taken",
                "severity": "Medium",
                "description": "Original capture time is missing, weakening timeline analysis.",
            }
        )
    elif parsed_date is None:
        anomalies.append(
            {
                "type": "Unparseable Date",
                "severity": "Medium",
                "description": f"Capture timestamp '{date_taken}' does not match supported EXIF date formats.",
            }
        )
    elif parsed_date > datetime.now():
        anomalies.append(
            {
                "type": "Future Capture Time",
                "severity": "Medium",
                "description": "Capture timestamp appears to be in the future relative to the workstation clock.",
            }
        )

    if software:
        software_lower = software.lower()
        for keyword in EDITING_SOFTWARE_KEYWORDS:
            if keyword in software_lower:
                anomalies.append(
                    {
                        "type": "Editing Software Detected",
                        "severity": "High",
                        "description": f"Metadata references possible editing software: {software}",
                    }
                )
                break

    if not camera_make and not camera_model and has_exif:
        anomalies.append(
            {
                "type": "Missing Camera Identity",
                "severity": "Low",
                "description": "EXIF exists but camera make/model fields are empty.",
            }
        )

    try:
        file_stats = file_path.stat()

        if file_stats.st_size == 0:
            anomalies.append(
                {
                    "type": "Zero-byte File",
                    "severity": "High",
                    "description": "The file size is 0 bytes, indicating corruption or an invalid evidence file.",
                }
            )

        modified_time = datetime.fromtimestamp(file_stats.st_mtime)
        if parsed_date and modified_time > parsed_date:
            anomalies.append(
                {
                    "type": "File Modified After Capture",
                    "severity": "Low",
                    "description": "Filesystem modification time is later than the recorded capture time.",
                }
            )
    except OSError:
        anomalies.append(
            {
                "type": "File Access Issue",
                "severity": "Medium",
                "description": "The evidence file could not be fully inspected on disk.",
            }
        )

    return anomalies
