import sys
from swap_ui import SwapUi
from PyQt5 import QtWidgets, uic
import os

if __name__ == '__main__':
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    sys.argv += ['--style', 'fusion']
    app = QtWidgets.QApplication(sys.argv)
    window = SwapUi()
    window.show()
    app.exec()
