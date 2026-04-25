def _ratio_to_float(value):
    """Convert EXIF ratio-like values to ``float``."""
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, tuple) and len(value) == 2:
        numerator, denominator = value
        denominator = float(denominator) if denominator else 0.0
        return float(numerator) / denominator if denominator else None

    if isinstance(value, str):
        if "/" in value:
            numerator, denominator = value.split("/", 1)
            denominator = float(denominator) if denominator else 0.0
            return float(numerator) / denominator if denominator else None
        try:
            return float(value)
        except ValueError:
            return None

    for numerator_attr, denominator_attr in (("num", "den"), ("numerator", "denominator")):
        if hasattr(value, numerator_attr) and hasattr(value, denominator_attr):
            denominator = float(getattr(value, denominator_attr))
            return float(getattr(value, numerator_attr)) / denominator if denominator else None

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _extract_sequence(value):
    if value is None:
        return None
    if hasattr(value, "values"):
        return list(value.values)
    if isinstance(value, (list, tuple)):
        return list(value)
    return None


def convert_to_degrees(value):
    """
    Convert EXIF DMS coordinates to decimal degrees.

    Supports both ``exifread`` tag objects and Pillow GPS tuples/lists.
    """
    values = _extract_sequence(value)
    if not values or len(values) < 3:
        return None

    degrees = _ratio_to_float(values[0])
    minutes = _ratio_to_float(values[1])
    seconds = _ratio_to_float(values[2])

    if degrees is None or minutes is None or seconds is None:
        return None

    return degrees + (minutes / 60.0) + (seconds / 3600.0)


def _normalize_ref(value):
    if value is None:
        return None
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="ignore").strip().upper()
    return str(value).strip().upper()


def decode_gps(tags):
    """
    Extract and decode GPS latitude/longitude from EXIF tags.

    Accepts either ``exifread`` tag names or Pillow-style dictionaries.
    """
    gps_latitude = tags.get("GPS GPSLatitude") or tags.get("GPSLatitude")
    gps_latitude_ref = tags.get("GPS GPSLatitudeRef") or tags.get("GPSLatitudeRef")
    gps_longitude = tags.get("GPS GPSLongitude") or tags.get("GPSLongitude")
    gps_longitude_ref = tags.get("GPS GPSLongitudeRef") or tags.get("GPSLongitudeRef")

    if not gps_latitude or not gps_latitude_ref or not gps_longitude or not gps_longitude_ref:
        return None, None

    latitude = convert_to_degrees(gps_latitude)
    longitude = convert_to_degrees(gps_longitude)

    if latitude is None or longitude is None:
        return None, None

    if _normalize_ref(gps_latitude_ref) == "S":
        latitude = -latitude

    if _normalize_ref(gps_longitude_ref) == "W":
        longitude = -longitude

    return latitude, longitude


def verify_coordinates(latitude, longitude):
    """Validate decoded GPS coordinates for forensic use."""
    issues = []

    if latitude is None or longitude is None:
        issues.append("GPS coordinates are incomplete.")
        return {"is_valid": False, "issues": issues}

    if not (-90 <= latitude <= 90):
        issues.append("Latitude is outside the valid range (-90 to 90).")

    if not (-180 <= longitude <= 180):
        issues.append("Longitude is outside the valid range (-180 to 180).")

    if latitude == 0 and longitude == 0:
        issues.append("Coordinates resolve to 0,0, which is commonly a placeholder value.")

    return {"is_valid": not issues, "issues": issues}


def format_coordinates(latitude, longitude):
    if latitude is None or longitude is None:
        return "Unavailable"
    return f"{latitude:.6f}, {longitude:.6f}"
