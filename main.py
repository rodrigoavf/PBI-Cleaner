import sys
from PyQt6.QtWidgets import QApplication
from ui_main_window import MainWindow
from PyQt6.QtGui import QIcon

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.png"))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())