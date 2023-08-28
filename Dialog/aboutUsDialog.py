import os
import sys

from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QDialog, QPlainTextEdit, QApplication, QMessageBox

path = os.path.realpath(__file__)
DIR = os.path.dirname(os.path.dirname(path))

HTML_BODY = f"""
            <html>
                <h1>ABOUT US<h1>
            </html>
            """


class aboutUs(QDialog):
    def __init__(self):
        super(aboutUs, self).__init__()
        self.load_Ui()

    def load_Ui(self):
        try:
            uic.loadUi(DIR + "/Dialog/UI/aboutUs.ui", self)
            with open(DIR + '/Stylesheet/Combinear.qss') as f:
                qss = f.read()
            # self.setStyleSheet(qss)
            self.setWindowTitle("About Us")
            self.aboutustextedit = self.findChild(QPlainTextEdit, "aboutustextedit")
            self.aboutustextedit.appendHtml(HTML_BODY)
        except Exception as e:
            self.warningBox(e)

    def warningBox(self,txt):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(txt)
        msg.setWindowTitle("Warning")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()

if __name__ == "__main__":
    app = QApplication([])
    widget = aboutUs()
    widget.show()
    sys.exit(app.exec_())
