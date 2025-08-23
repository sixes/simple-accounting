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
    # Set window size to 80% of the screen size
    screen = app.primaryScreen()
    size = screen.availableGeometry()
    width = int(size.width() * 0.8)
    height = int(size.height() * 0.8)
    win.resize(width, height)
    # Center the window on the screen
    x = size.x() + (size.width() - width) // 2
    y = size.y() + (size.height() - height) // 2
    win.move(x, y)
    win.show()
    sys.exit(app.exec())
