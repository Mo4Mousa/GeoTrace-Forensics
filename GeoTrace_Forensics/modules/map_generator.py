from pathlib import Path

import folium
from folium.plugins import MarkerCluster, TimestampedGeoJson


def _safe_slug(value):
    return "".join(char.lower() if char.isalnum() else "_" for char in value).strip("_") or "case"


def generate_interactive_map(case_name, timeline_data, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    events = timeline_data.get("events", [])
    geo_events = [
        event
        for event in events
        if event.get("latitude") is not None and event.get("longitude") is not None
    ]

    if geo_events:
        center = [
            sum(event["latitude"] for event in geo_events) / len(geo_events),
            sum(event["longitude"] for event in geo_events) / len(geo_events),
        ]
    else:
        center = [20.0, 0.0]

    map_object = folium.Map(location=center, zoom_start=3, control_scale=True)
    marker_cluster = MarkerCluster(name="Evidence Markers").add_to(map_object)
    marker_points = []
    timestamp_features = []

    for event in geo_events:
        integrity_status = event.get("integrity_status", "Unknown")
        duplicate_count = event.get("duplicate_count", 0)
        anomaly_count = event.get("anomaly_count", 0)

        speed = event.get("speed_kmh")
        suspicious = event.get("suspicious_movement", False)

        popup_html = (
            f"<b>{event['file_name']}</b><br>"
            f"Captured: {event.get('date_taken') or 'Unavailable'}<br>"
            f"Camera: {event.get('camera_make') or 'Unknown'} {event.get('camera_model') or ''}<br>"
            f"Integrity: {integrity_status}<br>"
            f"Duplicates: {duplicate_count}<br>"
            f"Anomalies: {anomaly_count}<br>"
            f"Speed: {speed if speed is not None else 'N/A'} km/h<br>"
            f"Suspicious Movement: {'Yes' if suspicious else 'No'}"
        )

        
        icon_type = "camera"
        if duplicate_count > 0:
            icon_type = "copy"

        marker_color = "green"
        if suspicious:
            marker_color = "darkred"
        elif integrity_status != "Verified":
            marker_color = "red"
        elif anomaly_count > 0:
            marker_color = "orange"

        folium.Marker(
            location=[event["latitude"], event["longitude"]],
            popup=popup_html,
            tooltip=f"{event['sequence']}. {event['file_name']}",
            icon=folium.Icon(color=marker_color, icon=icon_type, prefix="fa"),
        ).add_to(marker_cluster)

        #BONUS: highlight suspicious locations
        if suspicious:
            folium.Circle(
                location=[event["latitude"], event["longitude"]],
                radius=200,
                color="darkred",
                fill=True,
                fill_opacity=0.3,
                tooltip="Suspicious movement detected",
            ).add_to(map_object)

        marker_points.append([event["latitude"], event["longitude"]])

        if event.get("taken_at"):
            timestamp_features.append(
                {
                    "type": "Feature",
                    "geometry": {
                        "type": "Point",
                        "coordinates": [event["longitude"], event["latitude"]],
                    },
                    "properties": {
                        "time": event["taken_at"].isoformat(),
                        "popup": popup_html,
                        "icon": "circle",
                        "iconstyle": {
                            "fillColor": "#c44536",
                            "fillOpacity": 0.85,
                            "stroke": "true",
                            "radius": 7,
                        },
                    },
                }
            )

    # movement path
    if len(marker_points) >= 2:
        folium.PolyLine(
            marker_points,
            color="#264653",
            weight=3,
            opacity=0.8,
            dash_array="5, 10",
            tooltip="Chronological movement path",
        ).add_to(map_object)

    if timestamp_features:
        TimestampedGeoJson(
            {
                "type": "FeatureCollection",
                "features": timestamp_features,
            },
            period="PT1M",
            add_last_point=True,
            auto_play=False,
            loop=False,
            max_speed=10,
            loop_button=True,
            date_options="YYYY-MM-DD HH:mm:ss",
            time_slider_drag_update=True,
        ).add_to(map_object)

    #BONUS: legend
    legend_html = """
     <div style="
     position: fixed;
     bottom: 50px; left: 50px; width: 220px;
     background-color: white;
     border:2px solid grey;
     z-index:9999;
     font-size:14px;
     padding: 10px;">
     <b>Legend</b><br>
     <i style="color:green;">●</i> Normal<br>
     <i style="color:orange;">●</i> Anomaly<br>
     <i style="color:red;">●</i> Integrity Issue<br>
     <i style="color:darkred;">●</i> Suspicious Movement<br>
     </div>
     """
    map_object.get_root().html.add_child(folium.Element(legend_html))

    #BONUS: summary box
    summary = timeline_data.get("summary", {})
    summary_html = f"""
    <div style="
    position: fixed;
    top: 50px; right: 50px;
    background-color: white;
    border:2px solid grey;
    padding: 10px;
    z-index:9999;
    font-size:14px;">
    <b>Case Summary</b><br>
    Images: {summary.get('total_images', 0)}<br>
    With GPS: {summary.get('images_with_gps', 0)}<br>
    Anomalies: {summary.get('total_anomalies', 0)}<br>
    Distance: {summary.get('total_distance_km', 0)} km
    </div>
    """
    map_object.get_root().html.add_child(folium.Element(summary_html))

    if not geo_events:
        folium.Marker(
            location=center,
            popup="No GPS-enabled images were found for this case.",
            tooltip="No geolocation data",
        ).add_to(map_object)

    folium.LayerControl().add_to(map_object)

    output_path = output_dir / f"{_safe_slug(case_name)}_map.html"
    map_object.save(str(output_path))
    return output_path