import sys
from PyQt5.QtWidgets import QApplication
from ui.style.theme import STYLESHEET
from ui.app.config_editor import ConfigEditor

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    win = ConfigEditor()
    win.show()
    sys.exit(app.exec_())
