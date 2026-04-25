import json
from pathlib import Path

import pandas as pd


def export_case_json(case_info, case_results, timeline_data, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"case_{case_info['case_id']}_export.json"

    payload = {
        "case": case_info,
        "images": case_results,
        "timeline": {
            "summary": timeline_data.get("summary", {}),
            "events": [
                {
                    **event,
                    "taken_at": event["taken_at"].isoformat() if event.get("taken_at") else None,
                    "uploaded_at": event["uploaded_at"].isoformat() if event.get("uploaded_at") else None,
                }
                for event in timeline_data.get("events", [])
            ],
            "correlations": timeline_data.get("correlations", []),
        },
    }

    output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return output_path


def export_case_excel(case_info, case_results, timeline_data, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"case_{case_info['case_id']}_export.xlsx"

    evidence_rows = []
    anomaly_rows = []
    for row in case_results:
        evidence_rows.append(
            {
                "image_id": row["image_id"],
                "file_name": row["file_name"],
                "file_path": row["file_path"],
                "file_size": row["file_size"],
                "sha256_hash": row["sha256_hash"],
                "uploaded_at": row["uploaded_at"],
                "camera_make": row["camera_make"],
                "camera_model": row["camera_model"],
                "date_taken": row["date_taken"],
                "latitude": row["latitude"],
                "longitude": row["longitude"],
                "software": row["software"],
                "image_width": row["image_width"],
                "image_height": row["image_height"],
                "integrity_status": row.get("integrity_status"),
                "integrity_details": row.get("integrity_details"),
                "duplicate_count": row.get("duplicate_count", 0),
            }
        )

        for anomaly in row.get("anomalies", []):
            anomaly_rows.append(
                {
                    "image_id": row["image_id"],
                    "file_name": row["file_name"],
                    "anomaly_type": anomaly["anomaly_type"],
                    "severity": anomaly["severity"],
                    "description": anomaly["description"],
                }
            )

    timeline_rows = [
        {
            "sequence": event["sequence"],
            "file_name": event["file_name"],
            "date_taken": event["date_taken"],
            "latitude": event["latitude"],
            "longitude": event["longitude"],
            "time_delta_minutes": event["time_delta_minutes"],
            "distance_from_previous_km": event["distance_from_previous_km"],
            "anomaly_count": event["anomaly_count"],
        }
        for event in timeline_data.get("events", [])
    ]

    summary_rows = [{"metric": key, "value": value} for key, value in timeline_data.get("summary", {}).items()]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame([case_info]).to_excel(writer, sheet_name="Case", index=False)
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame(evidence_rows).to_excel(writer, sheet_name="Evidence", index=False)
        pd.DataFrame(timeline_rows).to_excel(writer, sheet_name="Timeline", index=False)
        pd.DataFrame(anomaly_rows).to_excel(writer, sheet_name="Anomalies", index=False)

    return output_path
