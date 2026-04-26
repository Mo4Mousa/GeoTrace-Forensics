from datetime import datetime
from math import asin, cos, radians, sin, sqrt


DATE_FORMATS = [
    "%Y:%m:%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
]


def parse_datetime(value):
    if not value:
        return None

    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(str(value), date_format)
        except ValueError:
            continue
    return None


def haversine_km(lat1, lon1, lat2, lon2):
    if None in (lat1, lon1, lat2, lon2):
        return None

    earth_radius_km = 6371.0
    delta_lat = radians(lat2 - lat1)
    delta_lon = radians(lon2 - lon1)
    lat1 = radians(lat1)
    lat2 = radians(lat2)

    a = (
        sin(delta_lat / 2) ** 2
        + cos(lat1) * cos(lat2) * sin(delta_lon / 2) ** 2
    )
    return 2 * earth_radius_km * asin(sqrt(a))


def generate_timeline(case_results, time_window_minutes=60, distance_window_km=2.0):
    events = []

    for row in case_results:
        taken_at = parse_datetime(row.get("date_taken"))
        uploaded_at = parse_datetime(row.get("uploaded_at"))
        anomalies = row.get("anomalies", [])

        events.append(
            {
                "image_id": row.get("image_id"),
                "file_name": row.get("file_name"),
                "file_path": row.get("file_path"),
                "date_taken": row.get("date_taken"),
                "taken_at": taken_at,
                "uploaded_at": uploaded_at,
                "latitude": row.get("latitude"),
                "longitude": row.get("longitude"),
                "camera_make": row.get("camera_make"),
                "camera_model": row.get("camera_model"),
                "sha256_hash": row.get("sha256_hash"),
                "anomaly_count": len(anomalies),
                "anomalies": anomalies,
            }
        )

    events.sort(
        key=lambda event: (
            event["taken_at"] is None,
            event["taken_at"] or event["uploaded_at"] or datetime.max,
        )
    )

    correlations = []
    previous_event = None

    for index, event in enumerate(events, start=1):
        event["sequence"] = index
        event["time_delta_minutes"] = None
        event["distance_from_previous_km"] = None

        if previous_event:
            if event["taken_at"] and previous_event["taken_at"]:
                event["time_delta_minutes"] = round(
                    (event["taken_at"] - previous_event["taken_at"]).total_seconds() / 60,
                    2,
                )

            event["distance_from_previous_km"] = (
                round(
                    haversine_km(
                        previous_event["latitude"],
                        previous_event["longitude"],
                        event["latitude"],
                        event["longitude"],
                    ),
                    3,
                )
                if None
                not in (
                    previous_event["latitude"],
                    previous_event["longitude"],
                    event["latitude"],
                    event["longitude"],
                )
                else None
            )

            #Speed calculation (km/h)
            speed_kmh = None
            if (
                event["time_delta_minutes"] is not None
                and event["distance_from_previous_km"] is not None
                and event["time_delta_minutes"] > 0
            ):
                hours = event["time_delta_minutes"] / 60
                speed_kmh = round(event["distance_from_previous_km"] / hours, 2)

            event["speed_kmh"] = speed_kmh

            #Suspicious movement detection
            event["suspicious_movement"] = False
            if speed_kmh is not None and speed_kmh > 300:
                event["suspicious_movement"] = True

            
            if (
                event["time_delta_minutes"] is not None
                and event["distance_from_previous_km"] is not None
                and event["time_delta_minutes"] <= time_window_minutes
                and event["distance_from_previous_km"] <= distance_window_km
            ):
                correlations.append(
                    {
                        "from_image": previous_event["file_name"],
                        "to_image": event["file_name"],
                        "time_delta_minutes": event["time_delta_minutes"],
                        "distance_km": event["distance_from_previous_km"],
                        
                        "summary": (
                            f"Movement detected from {previous_event['file_name']} to {event['file_name']}: "
                            f"{event['distance_from_previous_km']} km in {event['time_delta_minutes']} minutes."
                        ),
                    }
                )

        previous_event = event

    #Total distance calculation
    total_distance = sum(
        event["distance_from_previous_km"] or 0 for event in events
    )

    summary = {
        "total_images": len(events),
        "images_with_gps": sum(
            1 for event in events if event["latitude"] is not None and event["longitude"] is not None
        ),
        "images_with_timestamps": sum(1 for event in events if event["taken_at"] is not None),
        "total_anomalies": sum(event["anomaly_count"] for event in events),
        "earliest_capture": next(
            (event["date_taken"] for event in events if event["taken_at"] is not None),
            None,
        ),
        "latest_capture": next(
            (
                event["date_taken"]
                for event in reversed(events)
                if event["taken_at"] is not None
            ),
            None,
        ),
        #NEW field
        "total_distance_km": round(total_distance, 2),
    }

    return {
        "events": events,
        "summary": summary,
        "correlations": correlations,
    }