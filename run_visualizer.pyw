import sys

from PyQt5.QtWidgets import QApplication

from main import VisualizerWindow


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VisualizerWindow()
    window.show()
    sys.exit(app.exec_())
