import sys
from PyQt6.QtWidgets import QApplication
from gui import PDFToolboxApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToolboxApp()
    window.show()
    sys.exit(app.exec())

