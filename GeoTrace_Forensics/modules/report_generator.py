from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def _paragraphs_from_anomalies(anomalies):
    if not anomalies:
        return ["No anomalies detected."]

    return [
        f"{item['severity']} | {item['anomaly_type']}: {item['description']}"
        for item in anomalies
    ]


def generate_forensic_report(case_info, case_results, timeline_data, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"case_{case_info['case_id']}_forensic_report.pdf"
    logo_path = Path(__file__).resolve().parent.parent / "assets" / "project_logo.png"

    styles = getSampleStyleSheet()
    story = []

    if logo_path.exists():
        story.append(RLImage(str(logo_path), width=15.5 * cm, height=5.75 * cm))
        story.append(Spacer(1, 0.25 * cm))

    story.append(Paragraph("GeoTrace Forensics Report", styles["Title"]))
    story.append(Spacer(1, 0.4 * cm))
    story.append(
        Paragraph(
            (
                f"<b>Case:</b> {case_info['case_name']}<br/>"
                f"<b>Investigator:</b> {case_info['investigator_name']}<br/>"
                f"<b>Created:</b> {case_info['created_at']}"
            ),
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 0.4 * cm))

    summary = timeline_data.get("summary", {})
    summary_rows = [
        ["Metric", "Value"],
        ["Total Images", str(summary.get("total_images", 0))],
        ["Images With GPS", str(summary.get("images_with_gps", 0))],
        ["Images With Timestamps", str(summary.get("images_with_timestamps", 0))],
        ["Total Anomalies", str(summary.get("total_anomalies", 0))],
        ["Earliest Capture", summary.get("earliest_capture") or "Unavailable"],
        ["Latest Capture", summary.get("latest_capture") or "Unavailable"],
    ]

    summary_table = Table(summary_rows, colWidths=[6 * cm, 9 * cm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#264653")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(Paragraph("Case Summary", styles["Heading2"]))
    story.append(summary_table)
    story.append(Spacer(1, 0.5 * cm))

    evidence_rows = [["Image", "Path", "SHA-256", "Integrity", "Duplicates"]]
    for row in case_results:
        evidence_rows.append(
            [
                row["file_name"],
                row["file_path"],
                row["sha256_hash"],
                row.get("integrity_status") or "Unknown",
                str(row.get("duplicate_count", 0)),
            ]
        )

    evidence_table = Table(
        evidence_rows,
        colWidths=[2.8 * cm, 4.9 * cm, 5.2 * cm, 2.2 * cm, 1.7 * cm],
        repeatRows=1,
    )
    evidence_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2a9d8f")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("PADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(Paragraph("Evidence Chain", styles["Heading2"]))
    story.append(evidence_table)
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph("Timeline Findings", styles["Heading2"]))
    for event in timeline_data.get("events", []):
        event_text = (
            f"<b>{event['sequence']}. {event['file_name']}</b><br/>"
            f"Capture Time: {event.get('date_taken') or 'Unavailable'}<br/>"
            f"Coordinates: {event.get('latitude')}, {event.get('longitude')}<br/>"
            f"Camera: {event.get('camera_make') or 'Unknown'} {event.get('camera_model') or ''}<br/>"
            f"Integrity: {event.get('integrity_status') or 'Unknown'}<br/>"
            f"Duplicates: {event.get('duplicate_count', 0)}<br/>"
            f"Anomaly Count: {event.get('anomaly_count', 0)}"
        )
        story.append(Paragraph(event_text, styles["BodyText"]))
        story.append(Spacer(1, 0.18 * cm))

    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Correlated Images", styles["Heading2"]))
    correlations = timeline_data.get("correlations", [])
    if correlations:
        for correlation in correlations:
            story.append(Paragraph(correlation["summary"], styles["BodyText"]))
            story.append(Spacer(1, 0.15 * cm))
    else:
        story.append(Paragraph("No strong time/location correlations were detected.", styles["BodyText"]))

    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Anomaly Register", styles["Heading2"]))
    for row in case_results:
        story.append(Paragraph(f"<b>{row['file_name']}</b>", styles["BodyText"]))
        if row.get("integrity_details"):
            story.append(Paragraph(f"Integrity Details: {row['integrity_details']}", styles["BodyText"]))
        for anomaly_text in _paragraphs_from_anomalies(row.get("anomalies", [])):
            story.append(Paragraph(anomaly_text, styles["BodyText"]))
        story.append(Spacer(1, 0.18 * cm))

    document = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
    )
    document.build(story)
    return output_path
