import sys
import logging
from PySide6.QtWidgets import QApplication
from excel_like import ExcelLike

# Setup logging to file
logging.basicConfig(
    filename='banknote.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s %(message)s'
)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelLike()
    win.resize(1000, 600)
    win.show()
    sys.exit(app.exec())