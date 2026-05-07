# 🔍 GeoTrace Forensics

> **A digital image forensic investigation tool that extracts EXIF metadata, decodes GPS coordinates, detects image manipulation, and maintains a tamper-proof chain of custody.**

Built as a final-year forensics project for **El Sewedy University of Technology — Polytechnic of Egypt**.
Supervised by **Dr. Mariam Adel**.

---

## 🎯 What It Does

GeoTrace Forensics is a desktop application that helps digital investigators:

- 📷 Extract hidden EXIF metadata from any image
- 📍 Decode GPS coordinates and plot them on an interactive map
- ⏱️ Build a chronological timeline of where and when photos were taken
- 🚨 Detect signs of image manipulation, editing, and metadata stripping
- 🤖 Run AI-powered authenticity analysis using Error Level Analysis (ELA)
- 🔗 Maintain a blockchain-based tamper-proof evidence log
- 📑 Generate professional forensic PDF reports

---

## ✨ Key Features

### Core Forensic Capabilities
- **Multi-format EXIF extraction** — supports JPG, PNG, TIFF, BMP, WEBP, HEIC, GIF
- **GPS decoder** — converts raw EXIF GPS rationals to decimal degrees
- **Interactive folium maps** — with timestamped animation, marker clustering, and movement paths
- **Anomaly detection** — flags missing camera info, software traces, date mismatches, stripped GPS
- **SHA-256 file hashing** — for evidence integrity verification

### 🤖 AI Manipulation Detector (Bonus)
Three forensic techniques combined:
1. **Error Level Analysis (ELA)** — visualizes JPEG compression inconsistencies in edited regions
2. **Noise consistency analysis** — measures sensor noise variance across image quadrants
3. **Metadata cross-validation** — detects Photoshop traces, date mismatches, missing thumbnails

Outputs a confidence score and verdict: ✅ Authentic / ⚠️ Possibly Modified / 🚨 Manipulated.

### 🔗 Blockchain Chain of Custody (Bonus)
Each evidence file is recorded as a block with:
- File name and SHA-256 hash
- Investigator name
- Timestamp
- Link to previous block (cryptographic chaining)

Any modification to the evidence log breaks the chain — instantly detectable. This proves in court that evidence has not been altered post-submission.

### Additional Features
- Multi-case management with persistent SQLite database
- Excel and JSON export for inter-tool compatibility
- Image preview and raw EXIF viewer
- Live integrity verification on demand

---

## 📸 Screenshots

> **Add screenshots here before submission:**
> - `screenshots/main_window.png` — main investigation panel
> - `screenshots/map.png` — interactive geolocation map
> - `screenshots/ai_analysis.png` — AI manipulation detection result
> - `screenshots/blockchain.png` — chain of custody verification

---

## 🛠️ Tech Stack

| Component | Library |
|---|---|
| GUI | PyQt5 |
| Image processing | Pillow + NumPy |
| EXIF extraction | exifread |
| Mapping | Folium (Leaflet.js under the hood) |
| PDF generation | ReportLab |
| Database | SQLite3 (built-in) |
| Excel export | openpyxl |
| Cryptography | hashlib (SHA-256) |

---

## 🚀 Installation

### Prerequisites
- Python **3.10+**
- Windows, macOS, or Linux

### Setup
```bash
# 1. Clone the repository
git clone https://github.com/YOUR-USERNAME/GeoTrace-Forensics.git
cd GeoTrace-Forensics

# 2. (Recommended) Create a virtual environment
python -m venv venv

# Activate on Windows:
venv\Scripts\activate

# Activate on macOS/Linux:
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the application
python main.py
```

---

## 📖 How to Use

### 1. Create a New Case
- Enter a **Case Name** and **Investigator Name**
- Click **Create Case**

### 2. Add Evidence Images
- Click **Add Images**
- Select one or multiple images (JPG, PNG, TIFF, etc.)
- Each image is automatically:
  - Hashed with SHA-256
  - Logged to the blockchain
  - Analyzed for EXIF and anomalies

### 3. Investigate

| Action | Result |
|---|---|
| **Generate Map** | Opens an interactive map showing all GPS-tagged images with timeline |
| **Generate PDF Report** | Creates a complete forensic report with evidence chain |
| **Verify Integrity** | Re-hashes all files and detects tampering |
| **🤖 AI Analysis tab** | Runs manipulation detection on the selected image |
| **🔗 Blockchain tab** | Verifies the evidence log integrity |
| **Export Excel/JSON** | Exports data for use in other tools |

### 4. Read the Results
- **Anomalies tab** → list of suspicious findings per image
- **Raw EXIF tab** → complete metadata dump
- **Timeline tab** → chronological event list with locations

---

## 🧪 Testing the Tool

### Test Images Used
- **GPS-tagged authentic images** from [ianare/exif-samples](https://github.com/ianare/exif-samples)
- **Tampered images** from the [CASIA Image Tampering Dataset](https://www.kaggle.com/datasets/sophatvathana/casia-dataset)

### Validation Results

| Test Image | Expected | AI Verdict | Anomaly Count |
|---|---|---|---|
| DSCN0010.jpg (authentic, GPS) | Authentic | ✅ Likely Authentic | 1 |
| tampered_1.jpg (Photoshop edit) | Manipulated | 🚨 Probably Manipulated | 4 |
| tampered_2.jpg (Photoshop edit) | Manipulated | 🚨 Probably Manipulated | 4 |

### Demo Tampering Test
1. Add images to a case (creates blockchain blocks)
2. Click **Verify Chain** → ✅ Chain Intact
3. Manually edit `output/case_X/blockchain.json` (change a hash)
4. Click **Verify Chain** again → 🚨 **Chain BROKEN — Tampering Detected!**

---

## 📁 Project Structure

```
GeoTrace-Forensics/
├── main.py                        # Application entry point
├── requirements.txt               # Python dependencies
├── README.md                      # This file
├── assets/
│   └── project_logo.png           # Application logo
├── ui/
│   ├── main_window.py             # PyQt5 main window
│   └── style.qss                  # GUI stylesheet
├── modules/
│   ├── db_manager.py              # SQLite case database
│   ├── exif_extractor.py          # EXIF parsing
│   ├── gps_decoder.py             # GPS coordinate conversion
│   ├── hash_calculator.py         # SHA-256 file hashing
│   ├── anomaly_detector.py        # Metadata anomaly checks
│   ├── timeline_generator.py      # Chronological event builder
│   ├── map_generator.py           # Folium map renderer
│   ├── report_generator.py        # PDF report builder
│   ├── export_manager.py          # Excel + JSON exports
│   ├── ai_detector.py             # 🤖 AI manipulation detector (bonus)
│   └── blockchain_log.py          # 🔗 Blockchain evidence chain (bonus)
├── output/                        # Per-case artifacts (auto-created)
└── test_images/                   # Sample evidence for testing
```

---

## 🎓 Forensic Workflow

```
Add Image
   │
   ├─► SHA-256 Hash Calculation (file integrity)
   ├─► EXIF Extraction (camera, date, GPS, software)
   ├─► Anomaly Detection (missing fields, software traces)
   ├─► Blockchain Logging (tamper-proof chain)
   └─► Database Storage
         │
         ▼
   Optional Analyses:
   ├─► Map Generation (geolocation visualization)
   ├─► AI Manipulation Detection (ELA + noise + metadata)
   ├─► Timeline Building (chronological events)
   └─► PDF Report Generation
```

---

## ⚠️ Limitations

- ELA works best on JPEG images — accuracy may decrease on PNG
- AI verdict is heuristic-based, not deep learning — may flag heavily compressed images as manipulated
- GPS data is only available if the original camera/phone recorded it
- Some social media platforms strip EXIF on upload — this tool will detect that as an anomaly

---

## 🔮 Future Improvements

- Deep learning-based manipulation detection (CNN model)
- Reverse image search integration
- Face recognition with privacy controls
- Multi-language support for international cases
- Cloud-based collaborative case management
- Mobile companion app for field investigators

---

## 👥 Team

This project was developed by a team of 3 students for the Digital Forensics course.

- **[Selim ElSaadany]** 
- **[Lojain Emad]** 
- **[Mohamed Mousa]** 

---

## 🙏 Acknowledgments

- **Dr. Mariam Adel** — Project supervisor
- **El Sewedy University of Technology** — Polytechnic of Egypt
- **ianare/exif-samples** — Public test image dataset
- **CASIA** — Image tampering detection dataset
