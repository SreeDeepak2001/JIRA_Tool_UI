import sys
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QDialog, QSpinBox, QRadioButton, QCheckBox, QApplication, QMessageBox
from PyQt5 import uic
from configparser import ConfigParser
import os

path = os.path.realpath(__file__)
DIR = os.path.dirname(os.path.dirname(path))


class configDialog(QDialog):
    def __init__(self):
        super(configDialog, self).__init__()
        self.load_ui()
        self.rbtstate = 'All'
        self.config = ConfigParser()
        try:
            self.config.read(DIR+'/Config/config.ini')
        except Exception as e:
            self.warningBox(e)
        if self.config["Priority"]["critical"] == 'True':self.critical.setChecked(True)
        else:self.critical.setChecked(False)
        if self.config["Priority"]["severe"] == 'True':self.severe.setChecked(True)
        else:self.severe.setChecked(False)
        if self.config["Priority"]["moderate"] == 'True':self.moderate.setChecked(True)
        else:self.moderate.setChecked(False)
        if self.config["Priority"]["minor"] == 'True':self.minor.setChecked(True)
        else:self.minor.setChecked(False)
        if self.config["COMMENT_TYPE"]["internal"] == 'True':self.internal.setChecked(True)
        else:self.internal.setChecked(False)
        if self.config["COMMENT_TYPE"]["external"] == 'True':self.external.setChecked(True)
        else:self.external.setChecked(False)
        if self.config["COMMENT_TYPE"]["all"] == 'True':self.allcomments.setChecked(True)
        else:self.allcomments.setChecked(False)
        if self.config["Status"]["status"] == 'True':self.status.setChecked(True)
        else:self.status.setChecked(False)
        if self.config["Days_Config"]["exclude"] == 'True':self.excludelastcomment.setChecked(True)
        else:self.excludelastcomment.setChecked(False)
        self.criticaldays.setValue(int(self.config["Days_Config"]['critical']))
        self.severedays.setValue(int(self.config["Days_Config"]['severe']))
        self.moderatedays.setValue(int(self.config["Days_Config"]['moderate']))
        self.minordays.setValue(int(self.config["Days_Config"]['minor']))

    def load_ui(self):
        try:
            uic.loadUi(DIR+"/Dialog/UI/configDialog.ui",self)
            with open(DIR+'/Stylesheet/Combinear.qss') as f:
                qss = f.read()
        except Exception as e:
            self.warningBox(e)
        # self.setStyleSheet(qss)
        self.setWindowTitle("Configuration")
        self.setWindowIcon(QIcon(DIR + '/Icon/config.png'))
        self.critical = self.findChild(QCheckBox,'critical')
        self.severe = self.findChild(QCheckBox,'severe')
        self.moderate = self.findChild(QCheckBox,'moderate')
        self.minor =  self.findChild(QCheckBox,'minor')
        self.status =  self.findChild(QCheckBox,'includeClosed')
        self.internal = self.findChild(QRadioButton,'internal')
        self.internal.toggled.connect(lambda: self.updaterbstate())
        self.external = self.findChild(QRadioButton, 'external')
        self.external.toggled.connect(lambda: self.updaterbstate())
        self.allcomments = self.findChild(QRadioButton, 'allcomments')
        self.allcomments.toggled.connect(lambda: self.updaterbstate())
        self.criticaldays = self.findChild(QSpinBox,'criticaldays')
        self.severedays = self.findChild(QSpinBox,'severedays')
        self.moderatedays = self.findChild(QSpinBox,'moderatedays')
        self.minordays = self.findChild(QSpinBox,'minordays')
        self.excludelastcomment = self.findChild(QCheckBox, 'excludelastcomment')


    def accept(self):
        self.updateConfig()
        self.close()



    def updaterbstate(self):
        rbtn = self.sender()
        if rbtn.isChecked() == True:
            self.rbtstate = rbtn.text()

    def warningBox(self,txt):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(txt)
        msg.setWindowTitle("Warning")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()

    def updateConfig(self):
        if self.critical.checkState():
            self.config.set("Priority", "Critical", 'True')
        else:
            self.config.set("Priority", "Critical", 'False')
        if self.severe.checkState():
            self.config.set("Priority", "Severe", 'True')
        else:
            self.config.set("Priority", "Severe", 'False')
        if self.moderate.checkState():
            self.config.set("Priority", "Moderate", 'True')
        else:
            self.config.set("Priority", "Moderate", 'False')
        if self.minor.checkState():
            self.config.set("Priority", "Minor", 'True')
        else:
            self.config.set("Priority", "Minor", 'False')
        if self.internal.isChecked():
            self.config.set("COMMENT_TYPE", "Internal", 'True')
        else:
            self.config.set("COMMENT_TYPE", "Internal", 'False')
        if self.external.isChecked():
            self.config.set("COMMENT_TYPE", "External", 'True')
        else:
            self.config.set("COMMENT_TYPE", "External", 'False')
        if self.allcomments.isChecked():
            self.config.set("COMMENT_TYPE", "All", 'True')
        else:
            self.config.set("COMMENT_TYPE", "All", 'False')
        if self.status.checkState():
            self.config.set("Status", "status", "True")
        else:
            self.config.set("Status", "status", "False")
        if self.excludelastcomment.checkState():
            self.config.set("Days_Config", "exclude", "True")
        else:
            self.config.set("Days_Config", "exclude", "False")
        criticaldays = self.criticaldays.value()
        self.config.set('Days_Config', 'critical', str(criticaldays))
        severedays = self.severedays.value()
        self.config.set('Days_Config', 'severe', str(severedays))
        moderatedays = self.moderatedays.value()
        self.config.set('Days_Config', 'moderate', str(moderatedays))
        minordays = self.minordays.value()
        self.config.set('Days_Config', 'minor', str(minordays))
        try:
            with open(DIR+'/Config/config.ini','w') as configfile:
                self.config.write(configfile)
        except Exception as e:
            print(e)





if __name__ == "__main__":
    app = QApplication([])
    widget = configDialog()
    widget.show()
    sys.exit(app.exec_())