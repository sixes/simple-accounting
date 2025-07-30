import sys
from PySide6.QtWidgets import QApplication
from excel_like import ExcelLike

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelLike()
    win.resize(1000, 600)
    win.show()
    sys.exit(app.exec())