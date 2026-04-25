import sys
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from ui.main_window import MainWindow


def _load_stylesheet():
    style_path = Path(__file__).resolve().parent / "ui" / "style.qss"
    if not style_path.exists():
        return ""
    return style_path.read_text(encoding="utf-8")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("GeoTrace Forensics")

    stylesheet = _load_stylesheet()
    if stylesheet:
        app.setStyleSheet(stylesheet)

    window = MainWindow()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
