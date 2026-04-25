from pathlib import Path

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices, QPixmap
from PyQt5.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QHeaderView,
)

from modules.anomaly_detector import detect_anomalies
from modules.db_manager import DatabaseManager
from modules.exif_extractor import extract_exif_data
from modules.export_manager import export_case_excel, export_case_json
from modules.gps_decoder import format_coordinates
from modules.hash_calculator import calculate_sha256, get_file_size
from modules.map_generator import generate_interactive_map
from modules.report_generator import generate_forensic_report
from modules.timeline_generator import generate_timeline


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.current_case_id = None
        self.current_case = None
        self.current_results = []
        self.current_timeline = {"events": [], "summary": {}, "correlations": []}
        self.preview_path_label = None

        self.setWindowTitle("GeoTrace Forensics")
        self.resize(1320, 820)
        self._build_ui()
        self.refresh_cases()

    def _build_ui(self):
        central_widget = QWidget()
        root_layout = QVBoxLayout(central_widget)
        root_layout.setSpacing(14)

        root_layout.addWidget(self._build_brand_header())
        root_layout.addWidget(self._build_case_controls())
        root_layout.addWidget(self._build_action_bar())

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_results_table())
        splitter.addWidget(self._build_detail_tabs())
        splitter.setSizes([760, 520])
        root_layout.addWidget(splitter)

        self.setCentralWidget(central_widget)

    def _build_brand_header(self):
        container = QGroupBox()
        layout = QVBoxLayout(container)

        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)

        logo_path = Path(__file__).resolve().parent.parent / "assets" / "project_logo.png"
        pixmap = QPixmap(str(logo_path))
        if not pixmap.isNull():
            logo_label.setPixmap(
                pixmap.scaled(
                    680,
                    220,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation,
                )
            )
        else:
            logo_label.setText("GeoTrace Forensics")

        layout.addWidget(logo_label)
        return container

    def _build_case_controls(self):
        group = QGroupBox("Case Management")
        layout = QHBoxLayout(group)

        form_layout = QFormLayout()
        self.case_name_input = QLineEdit()
        self.investigator_input = QLineEdit()
        form_layout.addRow("Case Name", self.case_name_input)
        form_layout.addRow("Investigator", self.investigator_input)
        layout.addLayout(form_layout, stretch=2)

        right_panel = QVBoxLayout()
        self.case_selector = QComboBox()
        self.case_selector.currentIndexChanged.connect(self._load_selected_case)
        self.create_case_button = QPushButton("Create Case")
        self.create_case_button.clicked.connect(self.create_case)
        right_panel.addWidget(QLabel("Open Existing Case"))
        right_panel.addWidget(self.case_selector)
        right_panel.addWidget(self.create_case_button)
        right_panel.addStretch(1)
        layout.addLayout(right_panel, stretch=1)

        return group

    def _build_action_bar(self):
        group = QGroupBox("Investigation Actions")
        layout = QHBoxLayout(group)

        self.add_images_button = QPushButton("Add Images")
        self.add_images_button.clicked.connect(self.add_images)
        self.generate_map_button = QPushButton("Generate Map")
        self.generate_map_button.clicked.connect(self.generate_map)
        self.generate_report_button = QPushButton("Generate PDF Report")
        self.generate_report_button.clicked.connect(self.generate_report)
        self.verify_integrity_button = QPushButton("Verify Integrity")
        self.verify_integrity_button.clicked.connect(self.verify_integrity)
        self.export_excel_button = QPushButton("Export Excel")
        self.export_excel_button.clicked.connect(self.export_excel)
        self.export_json_button = QPushButton("Export JSON")
        self.export_json_button.clicked.connect(self.export_json)

        for button in (
            self.add_images_button,
            self.generate_map_button,
            self.generate_report_button,
            self.verify_integrity_button,
            self.export_excel_button,
            self.export_json_button,
        ):
            layout.addWidget(button)

        layout.addStretch(1)
        return group

    def _build_results_table(self):
        self.results_table = QTableWidget(0, 8)
        self.results_table.setHorizontalHeaderLabels(
            ["Image", "Capture Time", "Coordinates", "Camera", "Integrity", "Duplicates", "Anomalies", "SHA-256"]
        )
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.results_table.itemSelectionChanged.connect(self.display_selected_image_details)
        return self.results_table

    def _build_detail_tabs(self):
        self.details_tabs = QTabWidget()

        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.preview_label = QLabel("No image selected")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumHeight(260)
        self.preview_label.setStyleSheet("border: 1px dashed #d7c9b3; background: #fffdf8;")
        self.preview_path_label = QLabel("")
        self.preview_path_label.setWordWrap(True)
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.preview_path_label)

        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.anomalies_text = QTextEdit()
        self.anomalies_text.setReadOnly(True)
        self.raw_exif_text = QTextEdit()
        self.raw_exif_text.setReadOnly(True)
        self.timeline_text = QTextEdit()
        self.timeline_text.setReadOnly(True)

        self.details_tabs.addTab(preview_widget, "Preview")
        self.details_tabs.addTab(self.summary_text, "Metadata")
        self.details_tabs.addTab(self.anomalies_text, "Anomalies")
        self.details_tabs.addTab(self.raw_exif_text, "Raw EXIF")
        self.details_tabs.addTab(self.timeline_text, "Timeline")
        return self.details_tabs

    def _require_case(self):
        if self.current_case_id is None:
            QMessageBox.warning(self, "No Active Case", "Create or select a case first.")
            return False
        return True

    def refresh_cases(self):
        cases = self.db.get_all_cases()
        self.case_selector.blockSignals(True)
        self.case_selector.clear()
        self.case_selector.addItem("Select a case...", None)
        for case in cases:
            label = f"#{case['case_id']} | {case['case_name']} | {case['investigator_name']}"
            self.case_selector.addItem(label, case["case_id"])
        self.case_selector.blockSignals(False)

        if self.current_case_id:
            index = self.case_selector.findData(self.current_case_id)
            if index >= 0:
                self.case_selector.setCurrentIndex(index)

    def create_case(self):
        case_name = self.case_name_input.text().strip()
        investigator_name = self.investigator_input.text().strip()

        if not case_name or not investigator_name:
            QMessageBox.warning(self, "Missing Data", "Enter both case name and investigator name.")
            return

        case_id = self.db.create_case(case_name, investigator_name)
        self.current_case_id = case_id
        self.current_case = self.db.get_case(case_id)
        self.case_name_input.clear()
        self.investigator_input.clear()
        self.refresh_cases()
        self.load_case(case_id)
        QMessageBox.information(self, "Case Created", f"Case #{case_id} is ready for evidence intake.")

    def _load_selected_case(self, _index=None):
        case_id = self.case_selector.currentData()
        if case_id:
            self.load_case(case_id)

    def load_case(self, case_id):
        self.current_case_id = case_id
        self.current_case = self.db.get_case(case_id)
        self.current_results = self._decorate_case_results(self.db.get_case_results(case_id))
        self.current_timeline = generate_timeline(self.current_results)
        self.populate_results_table()
        self.populate_timeline_tab()

    def _verify_integrity_for_row(self, row):
        file_path = Path(row["file_path"])

        if not file_path.exists():
            return "Missing", "Evidence file is missing from the recorded path."

        try:
            current_hash = calculate_sha256(file_path)
        except Exception as error:
            return "Error", f"Failed to compute SHA-256: {error}"

        if current_hash != row.get("sha256_hash"):
            return "Tampered", "Current SHA-256 hash does not match the stored evidence hash."

        return "Verified", "Current SHA-256 hash matches the stored evidence hash."

    def _decorate_case_results(self, results):
        for row in results:
            integrity_status, integrity_details = self._verify_integrity_for_row(row)
            duplicates = self.db.find_images_by_hash(
                row.get("sha256_hash"),
                exclude_image_id=row.get("image_id"),
            )

            anomalies = list(row.get("anomalies", []))
            existing_signatures = {
                (item.get("anomaly_type"), item.get("description"))
                for item in anomalies
            }

            if duplicates:
                duplicate_names = ", ".join(match["file_name"] for match in duplicates[:5])
                duplicate_description = (
                    f"Matching SHA-256 hash found in {len(duplicates)} other image(s): {duplicate_names}"
                )
                signature = ("Duplicate Evidence", duplicate_description)
                if signature not in existing_signatures:
                    anomalies.append(
                        {
                            "anomaly_type": "Duplicate Evidence",
                            "severity": "High",
                            "description": duplicate_description,
                        }
                    )

            if integrity_status != "Verified":
                integrity_description = integrity_details
                signature = ("Integrity Verification Failed", integrity_description)
                if signature not in existing_signatures:
                    anomalies.append(
                        {
                            "anomaly_type": "Integrity Verification Failed",
                            "severity": "High",
                            "description": integrity_description,
                        }
                    )

            row["integrity_status"] = integrity_status
            row["integrity_details"] = integrity_details
            row["duplicate_count"] = len(duplicates)
            row["duplicate_matches"] = duplicates
            row["anomalies"] = anomalies

        return results

    def add_images(self):
        if not self._require_case():
            return

        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Evidence Images",
            "",
            "Images (*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp *.heic *.heif *.gif);;All Files (*)",
        )
        if not file_paths:
            return

        imported = 0
        failures = []

        for file_path in file_paths:
            try:
                path = Path(file_path)
                sha256_hash = calculate_sha256(path)
                file_size = get_file_size(path)
                image_id = self.db.insert_image(
                    self.current_case_id,
                    path.name,
                    str(path.resolve()),
                    file_size,
                    sha256_hash,
                )

                metadata = extract_exif_data(path)
                self.db.insert_metadata(
                    image_id=image_id,
                    camera_make=metadata.get("camera_make"),
                    camera_model=metadata.get("camera_model"),
                    date_taken=metadata.get("date_taken"),
                    latitude=metadata.get("latitude"),
                    longitude=metadata.get("longitude"),
                    software=metadata.get("software"),
                    image_width=metadata.get("image_width"),
                    image_height=metadata.get("image_height"),
                    raw_exif=metadata.get("raw_exif"),
                )

                anomalies = detect_anomalies(path, metadata)
                for anomaly in anomalies:
                    self.db.insert_anomaly(
                        image_id=image_id,
                        anomaly_type=anomaly["type"],
                        severity=anomaly["severity"],
                        description=anomaly["description"],
                    )

                duplicate_matches = self.db.find_images_by_hash(sha256_hash, exclude_image_id=image_id)
                if duplicate_matches:
                    duplicate_names = ", ".join(match["file_name"] for match in duplicate_matches[:5])
                    self.db.insert_anomaly(
                        image_id=image_id,
                        anomaly_type="Duplicate Evidence",
                        severity="High",
                        description=(
                            f"Matching SHA-256 hash found in {len(duplicate_matches)} other image(s): "
                            f"{duplicate_names}"
                        ),
                    )

                imported += 1
            except Exception as error:
                failures.append(f"{file_path}: {error}")

        self.load_case(self.current_case_id)

        if failures:
            QMessageBox.warning(
                self,
                "Import Completed With Issues",
                f"Imported {imported} image(s).\n\nIssues:\n" + "\n".join(failures[:10]),
            )
        else:
            QMessageBox.information(self, "Import Complete", f"Imported {imported} image(s) successfully.")

    def populate_results_table(self):
        self.results_table.setRowCount(len(self.current_results))

        for row_index, row in enumerate(self.current_results):
            camera = " ".join(filter(None, [row.get("camera_make"), row.get("camera_model")])) or "Unknown"
            values = [
                row.get("file_name") or "",
                row.get("date_taken") or "Unavailable",
                format_coordinates(row.get("latitude"), row.get("longitude")),
                camera,
                row.get("integrity_status") or "Unknown",
                str(row.get("duplicate_count", 0)),
                str(len(row.get("anomalies", []))),
                row.get("sha256_hash") or "",
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column_index == 4 and row.get("integrity_status") != "Verified":
                    item.setForeground(Qt.darkRed)
                if column_index == 6 and row.get("anomalies"):
                    item.setForeground(Qt.darkRed)
                self.results_table.setItem(row_index, column_index, item)

        if self.current_results:
            self.results_table.selectRow(0)
        else:
            self.preview_label.setText("No image selected")
            self.preview_label.setPixmap(QPixmap())
            self.preview_path_label.clear()
            self.summary_text.clear()
            self.anomalies_text.clear()
            self.raw_exif_text.clear()

    def display_selected_image_details(self):
        row_index = self.results_table.currentRow()
        if row_index < 0 or row_index >= len(self.current_results):
            return

        row = self.current_results[row_index]
        anomaly_lines = [
            f"[{item['severity']}] {item['anomaly_type']}: {item['description']}"
            for item in row.get("anomalies", [])
        ]

        summary_lines = [
            f"File: {row.get('file_name')}",
            f"Path: {row.get('file_path')}",
            f"File Size: {row.get('file_size')}",
            f"SHA-256: {row.get('sha256_hash')}",
            f"Integrity Status: {row.get('integrity_status') or 'Unknown'}",
            f"Integrity Details: {row.get('integrity_details') or 'Unavailable'}",
            f"Duplicate Matches: {row.get('duplicate_count', 0)}",
            f"Uploaded: {row.get('uploaded_at')}",
            f"Camera Make: {row.get('camera_make') or 'Unavailable'}",
            f"Camera Model: {row.get('camera_model') or 'Unavailable'}",
            f"Date Taken: {row.get('date_taken') or 'Unavailable'}",
            f"Software: {row.get('software') or 'Unavailable'}",
            f"Coordinates: {format_coordinates(row.get('latitude'), row.get('longitude'))}",
            f"Dimensions: {row.get('image_width')} x {row.get('image_height')}",
        ]

        pixmap = QPixmap(row.get("file_path", ""))
        if not pixmap.isNull():
            scaled = pixmap.scaled(
                420,
                260,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation,
            )
            self.preview_label.setPixmap(scaled)
            self.preview_label.setText("")
        else:
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("Preview unavailable")
        self.preview_path_label.setText(row.get("file_path") or "")

        self.summary_text.setPlainText("\n".join(summary_lines))
        self.anomalies_text.setPlainText("\n".join(anomaly_lines) if anomaly_lines else "No anomalies detected.")
        self.raw_exif_text.setPlainText(row.get("raw_exif") or "{}")

    def populate_timeline_tab(self):
        if not self.current_case:
            self.timeline_text.clear()
            return

        summary = self.current_timeline.get("summary", {})
        lines = [
            f"Case: {self.current_case['case_name']}",
            f"Investigator: {self.current_case['investigator_name']}",
            f"Total Images: {summary.get('total_images', 0)}",
            f"Images With GPS: {summary.get('images_with_gps', 0)}",
            f"Images With Timestamps: {summary.get('images_with_timestamps', 0)}",
            f"Total Anomalies: {summary.get('total_anomalies', 0)}",
            f"Integrity Failures: {sum(1 for row in self.current_results if row.get('integrity_status') != 'Verified')}",
            f"Duplicate Evidence Items: {sum(1 for row in self.current_results if row.get('duplicate_count', 0) > 0)}",
            "",
            "Timeline Events:",
        ]

        for event in self.current_timeline.get("events", []):
            lines.append(
                (
                    f"{event['sequence']}. {event['file_name']} | "
                    f"{event.get('date_taken') or 'Unavailable'} | "
                    f"{format_coordinates(event.get('latitude'), event.get('longitude'))} | "
                    f"Integrity: {event.get('integrity_status') or 'Unknown'} | "
                    f"Anomalies: {event.get('anomaly_count', 0)}"
                )
            )

        lines.append("")
        lines.append("Correlations:")
        correlations = self.current_timeline.get("correlations", [])
        if correlations:
            for correlation in correlations:
                lines.append(f"- {correlation['summary']}")
        else:
            lines.append("- No strong time/location correlations detected.")

        self.timeline_text.setPlainText("\n".join(lines))

    def _artifact_dir(self):
        return self.db.get_case_artifact_dir(self.current_case_id)

    def _open_file(self, file_path):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(file_path)))

    def generate_map(self):
        if not self._require_case():
            return
        self.current_results = self._decorate_case_results(self.db.get_case_results(self.current_case_id))
        self.current_timeline = generate_timeline(self.current_results)
        self.populate_results_table()
        self.populate_timeline_tab()
        output_path = generate_interactive_map(
            self.current_case["case_name"],
            self.current_timeline,
            self._artifact_dir(),
        )
        self._open_file(output_path)

    def generate_report(self):
        if not self._require_case():
            return
        self.current_results = self._decorate_case_results(self.db.get_case_results(self.current_case_id))
        self.current_timeline = generate_timeline(self.current_results)
        output_path = generate_forensic_report(
            self.current_case,
            self.current_results,
            self.current_timeline,
            self._artifact_dir(),
        )
        self._open_file(output_path)

    def export_excel(self):
        if not self._require_case():
            return
        self.current_results = self._decorate_case_results(self.db.get_case_results(self.current_case_id))
        self.current_timeline = generate_timeline(self.current_results)
        output_path = export_case_excel(
            self.current_case,
            self.current_results,
            self.current_timeline,
            self._artifact_dir(),
        )
        self._open_file(output_path)

    def export_json(self):
        if not self._require_case():
            return
        self.current_results = self._decorate_case_results(self.db.get_case_results(self.current_case_id))
        self.current_timeline = generate_timeline(self.current_results)
        output_path = export_case_json(
            self.current_case,
            self.current_results,
            self.current_timeline,
            self._artifact_dir(),
        )
        self._open_file(output_path)

    def verify_integrity(self):
        if not self._require_case():
            return

        self.current_results = self._decorate_case_results(self.db.get_case_results(self.current_case_id))
        self.current_timeline = generate_timeline(self.current_results)
        self.populate_results_table()
        self.populate_timeline_tab()

        failed_count = sum(1 for row in self.current_results if row.get("integrity_status") != "Verified")
        QMessageBox.information(
            self,
            "Integrity Verification",
            (
                f"Verified {len(self.current_results)} image(s).\n"
                f"Integrity issues detected in {failed_count} image(s)."
            ),
        )
