# ver 1.2.2
# internal libraries
import sys
import subprocess
import pkg_resources
from ctypes import windll
import webbrowser
from os import getcwd
windll.shcore.SetProcessDpiAwareness(1)

# check for required packages - install if missing
required = {'pandas', 'openpyxl', 'pywin32', 'requests', 'python-certifi-win32', 'pyodbc', 'pyqt6', 'psutil'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed

if missing:
    python = sys.executable
    subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)
    windll.user32.MessageBoxW(0, "SETUP COMPLETE: RE-OPEN THE PROGRAM", "FIRST TIME SETUP", 0x0 | 0x40)
    sys.exit()

# external libraries
from SMTP_backend import *
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox

class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        window_length = 900
        window_width = 630

        MainWindow.resize(window_length, window_width)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("DEVELOPER_FILES/icon_email.png"), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        font = QtGui.QFont()
        font.setPointSize(10)
        font_bold = QtGui.QFont()
        font_bold.setPointSize(10)
        font_bold.setBold(True)

        self.sendMail = QtWidgets.QPushButton(parent=self.centralwidget)
        self.sendMail.setGeometry(QtCore.QRect(10, 12, 62, 80))
        self.sendMail.setFont(font_bold)
        self.sendMail.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("DEVELOPER_FILES/icon_send.png"), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        self.sendMail.setIcon(icon)
        self.sendMail.setObjectName("sendMail")
        self.sendMail.clicked.connect(self.sendLogic)

        self.lblTo = QtWidgets.QLabel(parent=self.centralwidget)
        self.lblTo.setGeometry(QtCore.QRect(90, 12, 20, 16))
        self.lblTo.setFont(font_bold)
        self.lblTo.setObjectName("lblTo")

        self.lblCC = QtWidgets.QLabel(parent=self.centralwidget)
        self.lblCC.setGeometry(QtCore.QRect(90, 42, 20, 16))
        self.lblCC.setFont(font_bold)
        self.lblCC.setObjectName("lblCC")

        self.lblSubject = QtWidgets.QLabel(parent=self.centralwidget)
        self.lblSubject.setGeometry(QtCore.QRect(90, 72, 60, 16))
        self.lblSubject.setFont(font_bold)
        self.lblSubject.setObjectName("lblSubject")

        self.lineSubject = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.lineSubject.setGeometry(QtCore.QRect(160, 70, window_length - 175, 22))
        self.lineSubject.setFont(font)
        self.lineSubject.setObjectName("lineSubject")

        self.pteMessage = QtWidgets.QPlainTextEdit(parent=self.centralwidget)
        self.pteMessage.setGeometry(QtCore.QRect(0, 107, window_length+1, window_width - 153))
        self.pteMessage.setFont(font)
        self.pteMessage.setFrameShape(QtWidgets.QFrame.Shape.Box)
        self.pteMessage.setPlainText("")
        self.pteMessage.setTabStopDistance(QtGui.QFontMetricsF(font).horizontalAdvance(' ') * 4)
        self.pteMessage.setObjectName("pteMessage")

        self.cbPrimary = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbPrimary.setGeometry(QtCore.QRect(160, 12, 75, 20))
        self.cbPrimary.setFont(font)
        self.cbPrimary.setChecked(True)
        self.cbPrimary.setObjectName("cbPrimary")

        self.cbOperation = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbOperation.setGeometry(QtCore.QRect(255, 12, 90, 20))
        self.cbOperation.setFont(font)
        self.cbOperation.setObjectName("cbOperation")

        self.cbCustomer = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbCustomer.setGeometry(QtCore.QRect(365, 12, 85, 20))
        self.cbCustomer.setFont(font)
        self.cbCustomer.setObjectName("cbCustomer")

        self.cbScorecard = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbScorecard.setGeometry(QtCore.QRect(470, 12, 90, 20))
        self.cbScorecard.setFont(font)
        self.cbScorecard.setObjectName("cbScorecard")

        self.cbShip = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbShip.setGeometry(QtCore.QRect(580, 12, 75, 20))
        self.cbShip.setFont(font)
        self.cbShip.setObjectName("cbShip")

        self.cbNPI = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbNPI.setGeometry(QtCore.QRect(675, 12, 40, 20))
        self.cbNPI.setFont(font)
        self.cbNPI.setObjectName("cbNPI")

        self.cbSPA = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbSPA.setGeometry(QtCore.QRect(160, 42, 40, 20))
        self.cbSPA.setFont(font)
        self.cbSPA.setObjectName("cbSPA")

        self.cbUser = QtWidgets.QCheckBox(parent=self.centralwidget)
        self.cbUser.setGeometry(QtCore.QRect(220, 42, 50, 20))
        self.cbUser.setFont(font)
        self.cbUser.setObjectName("cbUser")
        self.cbUser.setObjectName("cbUser")

        self.lblMode = QtWidgets.QLabel(parent=self.centralwidget)
        self.lblMode.setFont(font_bold)
        self.lblMode.setStyleSheet("QLabel {color:white;}"); #background-color : red;
        self.lblMode.setText("TEST MODE")

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setStyleSheet("QStatusBar{padding-left:8px;background:rgba(62,65,73,1);color:white;}") #font-weight:bold;
        self.statusbar.setObjectName("statusbar")
        self.statusbar.addPermanentWidget(self.lblMode)
        self.statusbar.setFont(font)
        MainWindow.setStatusBar(self.statusbar)

        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 24))
        self.menubar.setFont(font)
        self.menubar.setObjectName("menubar")
        self.menubar.setStyleSheet("""
        QMenuBar {
            background-color: rgba(63,66,75,1);
            color: white;
        }
        QMenuBar::item {
            background-color: rgba(63,66,75,1);
            color: white;
        }
        QMenuBar::item::selected {
            background-color: rgb(0,0,0);
        }
        QMenu {
            background-color: rgba(63,66,75,1);
            color: white;           
        }
        QMenu::item::selected {
            background-color: rgb(0,0,0);
        }""")

        self.menuSETTING = QtWidgets.QMenu(parent=self.menubar)
        self.menuSETTING.setFont(font)
        self.menuSETTING.setObjectName("menuSETTING")

        self.menuHELP = QtWidgets.QMenu(parent=self.menubar)
        self.menuHELP.setFont(font)
        self.menuHELP.setObjectName("menuHELP")

        self.menuFILE = QtWidgets.QMenu(parent=self.menubar)
        self.menuFILE.setFont(font)
        self.menuFILE.setObjectName("menuFILE")

        MainWindow.setMenuBar(self.menubar)
        self.actionTESTMODE = QtGui.QAction(parent=MainWindow)
        self.actionTESTMODE.setCheckable(True)
        self.actionTESTMODE.setChecked(True)
        self.actionTESTMODE.setFont(font)
        self.actionTESTMODE.setObjectName("actionTESTMODE")
        self.actionTESTMODE.triggered.connect(self.stateChange)

        self.actionSMTP = QtGui.QAction(parent=MainWindow)
        self.actionSMTP.setCheckable(True)
        self.actionSMTP.setChecked(True)
        self.actionSMTP.setFont(font)
        self.actionSMTP.setObjectName("actionSMTP")
        self.actionSMTP.triggered.connect(self.stateChange)

        self.actionMANUAL = QtGui.QAction(parent=MainWindow)
        self.actionMANUAL.setFont(font)
        self.actionMANUAL.setObjectName("actionMANUAL")
        self.actionMANUAL.triggered.connect(self.readMe)

        self.actionSave = QtGui.QAction(parent=MainWindow)
        self.actionSave.setFont(font)
        self.actionSave.setShortcutContext(QtCore.Qt.ShortcutContext.WidgetShortcut)
        self.actionSave.setObjectName("actionSave")

        self.actionLoad = QtGui.QAction(parent=MainWindow)
        self.actionLoad.setFont(font)
        self.actionLoad.setObjectName("actionLoad")

        self.menuSETTING.addAction(self.actionTESTMODE)
        self.menuSETTING.addAction(self.actionSMTP)
        self.menuHELP.addAction(self.actionMANUAL)
        self.menuFILE.addAction(self.actionSave)
        self.menuFILE.addAction(self.actionLoad)
        self.menubar.addAction(self.menuFILE.menuAction())
        self.menubar.addAction(self.menuSETTING.menuAction())
        self.menubar.addAction(self.menuHELP.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    
    def sendLogic(self):
        roles_selected = []
        if self.cbPrimary.isChecked(): roles_selected.append('Primary')
        if self.cbOperation.isChecked(): roles_selected.append('Operations')
        if self.cbCustomer.isChecked(): roles_selected.append('Customer Service')
        if self.cbScorecard.isChecked(): roles_selected.append('Scorecards')
        if self.cbShip.isChecked(): roles_selected.append('Shipping')
        if self.cbNPI.isChecked(): roles_selected.append('NPI')

        cc_selected = []
        if self.cbSPA.isChecked(): cc_selected.append('SPA')
        if self.cbUser.isChecked(): cc_selected.append('USER')

        em_message = self.pteMessage.toPlainText()
        em_subject = self.lineSubject.text()
        test_mode = True if self.actionTESTMODE.isChecked() else False
        SMTP_mode = True if self.actionSMTP.isChecked() else False

        current_time = strftime("%H:%M:%S", localtime())
        if not roles_selected:
            self.statusbar.showMessage(f'[{current_time}] ERROR: ATLEAST ONE ROLE MUST BE SELECTED')
            return

        if hasHandle('supplier_send_list.xlsx'):
            self.statusbar.showMessage(f'[{current_time}] ERROR: SUPPLIER SEND FILE IN USE')
            return

        if not test_mode:
            sendDialog = QMessageBox(MainWindow)
            sendDialog.setWindowTitle("E-MAIL IS LIVE")
            sendDialog.setText("DO YOU WANT TO SEND EMAIL?")
            sendDialog.setIcon(QMessageBox.Icon.Warning)
            sendDialog.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            reply = sendDialog.exec()

            if reply == QMessageBox.StandardButton.No:
                return

        error_message = sendSupplier(test_mode=test_mode,
            SMTP_mode=SMTP_mode,
            roles_selected=roles_selected,
            cc_selected=cc_selected,
            message=em_message, 
            subject=em_subject)
        self.statusbar.showMessage(error_message)


    def readMe(self):
        webbrowser.open_new(f'{getcwd()}/DOCUMENTATION/READ_ME.pdf')


    def stateChange(self):
        # use to change the test mode label on bottom right corner of UI
        if self.actionTESTMODE.isChecked():
            self.lblMode.setText("TEST MODE")
            self.actionTESTMODE.setChecked(True)
        else:
            self.lblMode.setText("LIVE MODE")
            self.actionTESTMODE.setChecked(False)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SUPPLIER E-MAIL TOOL v1.2.2"))
        self.sendMail.setText(_translate("MainWindow", "SEND"))
        self.lblTo.setText(_translate("MainWindow", "TO:"))
        self.lblCC.setText(_translate("MainWindow", "CC:"))
        self.lblSubject.setText(_translate("MainWindow", "SUBJECT:"))
        self.cbPrimary.setText(_translate("MainWindow", "PRIMARY"))
        self.cbOperation.setText(_translate("MainWindow", "OPERATION"))
        self.cbCustomer.setText(_translate("MainWindow", "C. SERVICE"))
        self.cbScorecard.setText(_translate("MainWindow", "SCORECARD"))
        self.cbShip.setText(_translate("MainWindow", "SHIPPING"))
        self.cbNPI.setText(_translate("MainWindow", "NPI"))
        self.cbSPA.setText(_translate("MainWindow", "SPA"))
        self.cbUser.setText(_translate("MainWindow", "USER"))
        self.menuSETTING.setTitle(_translate("MainWindow", "SETTING"))
        self.menuHELP.setTitle(_translate("MainWindow", "HELP"))
        self.menuFILE.setTitle(_translate("MainWindow", "FILE"))
        self.actionTESTMODE.setText(_translate("MainWindow", "Test Mode"))
        self.actionTESTMODE.setShortcut(_translate("MainWindow", "Ctrl+T"))
        self.actionSMTP.setText(_translate("MainWindow", "SMTP"))
        self.actionSMTP.setShortcut(_translate("MainWindow", "Ctrl+Y"))
        self.actionMANUAL.setText(_translate("MainWindow", "Manual"))
        self.actionMANUAL.setShortcut(_translate("MainWindow", "Ctrl+H"))
        self.actionSave.setText(_translate("MainWindow", "Save"))
        self.actionSave.setShortcut(_translate("MainWindow", "Ctrl+S"))
        self.actionLoad.setText(_translate("MainWindow", "Load"))
        self.actionLoad.setShortcut(_translate("MainWindow", "Ctrl+L"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())