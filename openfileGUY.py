# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'OpenFile.ui'
#
# Created by: PyQt5 UI code generator 5.13.1
#
# WARNING! All changes made in this file will be lost!
from PyQt5 import QtCore, QtGui, QtWidgets
import tempfile
import os
import sys


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(320, 159)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(os.path.join(sys._MEIPASS, "Iconaip.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off) #sys._MEIPASS,
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setPixmap(QtGui.QPixmap(os.path.join(sys._MEIPASS, 'nt_no_1bg.png')).scaled(80,80,QtCore.Qt.KeepAspectRatio)) #sys._MEIPASS,
        self.label.move(110,89)
        self.openfile = QtWidgets.QPushButton(self.centralwidget)
        self.openfile.setGeometry(QtCore.QRect(30, 22, 260, 31))
        self.openfile.setObjectName("openfile")
        self.bouton_valider = QtWidgets.QPushButton(self.centralwidget)
        self.bouton_valider.setGeometry(QtCore.QRect(30, 55, 260, 31))
        self.bouton_valider.setObjectName("bouton_valider")
        self.bouton_valider.setEnabled(True)
        self.bouton_valider.setStyleSheet("background-color:#F39100;color:white;")
        # Menu barre et ses boutons
        extractAction = QtWidgets.QAction(QtGui.QIcon(os.path.join(sys._MEIPASS, 'aide.ico')) , "Aide", self.centralwidget)
        extractAction.setShortcut("F1")
        extractAction.setStatusTip('Documentation')
        extractAction.triggered.connect(self.doc)
        closeapp = QtWidgets.QAction(QtGui.QIcon(os.path.join(sys._MEIPASS, 'exit.png')) , "Fermer", self.centralwidget)
        closeapp.setShortcut("ESC")
        closeapp.setStatusTip('Fermer Instadoc')
        closeapp.triggered.connect(self.closeapp)
        self.statusBar = QtWidgets.QStatusBar(self.centralwidget)
        self.menuBar = QtWidgets.QMenuBar(self.centralwidget)
        fileMenu = self.menuBar.addMenu('Menu')
        fileMenu.addAction(extractAction)
        fileMenu.addAction(closeapp)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 321, 160))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.retranslateUi(MainWindow)

        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def doc(self):
        try:
       
            os.system('start '+os.path.join(sys._MEIPASS,'aipdoc.pdf'))
        except:
            # QtWidgets.QMainWindow.QMessageBox.information(self, "Erreur", "La documentation n'est plus accessible." , QtWidgets.QMessageBox.Ok)
            msgBox = QtWidgets.QMessageBox()
            msgBox.setIcon(QtWidgets.QMessageBox.Information)
            msgBox.setText("La documentation n'est plus accessible.")
            msgBox.setInformativeText(r'Vous pouvez la retrouver sous V:\Procédures.')
            msgBox.setWindowTitle("Erreur")
            msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msgBox.exec()

    def closeapp(self):
        sys.exit()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "AIP - IMPORT"))
        MainWindow.setStatusTip(_translate("MainWindow", "Version 0.1"))
        self.openfile.setText(_translate("MainWindow", "Sélectionnez le fichier Excel à ouvrir"))
        self.bouton_valider.setText(_translate("MainWindow", "Lancer le traitement"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
