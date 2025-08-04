import sys
import logging
from PySide6.QtWidgets import QApplication
from excel_like import ExcelLike
import faulthandler
faulthandler.enable()

with open("traceback.log", "w") as f:
    faulthandler.enable(file=f)

# Setup logging to file
logging.basicConfig(
    filename='banknote.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s %(message)s',
    filemode='w'
)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelLike()
    win.resize(1000, 600)
    win.show()
    sys.exit(app.exec())
