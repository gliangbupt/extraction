# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'extraction_ui.ui',
# licensing of 'extraction_ui.ui' applies.
#
# Created: Tue Dec  4 18:01:27 2018
#      by: pyside2-uic  running on PySide2 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PySide2 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.setWindowModality(QtCore.Qt.NonModal)
        Form.resize(800, 200)
        Form.setMinimumSize(QtCore.QSize(800, 200))
        Form.setMaximumSize(QtCore.QSize(800, 200))
        self.Choose_btn = QtWidgets.QPushButton(Form)
        self.Choose_btn.setGeometry(QtCore.QRect(620, 50, 93, 31))
        self.Choose_btn.setObjectName("Choose_btn")
        self.dir_display = QtWidgets.QLabel(Form)
        self.dir_display.setGeometry(QtCore.QRect(80, 50, 81, 31))
        self.dir_display.setObjectName("dir_display")
        self.dir_line = QtWidgets.QLineEdit(Form)
        self.dir_line.setGeometry(QtCore.QRect(190, 50, 401, 31))
        self.dir_line.setObjectName("dir_line")
        self.ColorChecker_btn = QtWidgets.QPushButton(Form)
        self.ColorChecker_btn.setGeometry(QtCore.QRect(102, 121, 111, 28))
        self.ColorChecker_btn.setObjectName("ColorChecker_btn")
        self.Texture_btn = QtWidgets.QPushButton(Form)
        self.Texture_btn.setGeometry(QtCore.QRect(272, 121, 111, 28))
        self.Texture_btn.setObjectName("Texture_btn")
        self.Flash_btn = QtWidgets.QPushButton(Form)
        self.Flash_btn.setGeometry(QtCore.QRect(440, 121, 111, 28))
        self.Flash_btn.setObjectName("Flash_btn")
        self.GreyChart_btn = QtWidgets.QPushButton(Form)
        self.GreyChart_btn.setGeometry(QtCore.QRect(600, 121, 111, 28))
        self.GreyChart_btn.setObjectName("GreyChart_btn")

        appIcon = QtGui.QIcon("D:\My document\work\Execl提取数据\Data_extraction\\Excel.jpg")
        Form.setWindowIcon(appIcon)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(QtWidgets.QApplication.translate("Form", "MCE_Extraction", None, -1))
        self.Choose_btn.setText(QtWidgets.QApplication.translate("Form", "选择文件夹", None, -1))
        self.dir_display.setText(QtWidgets.QApplication.translate("Form", "<html><head/><body><p><span style=\" font-size:12pt;\">目录位置</span></p></body></html>", None, -1))
        self.ColorChecker_btn.setText(QtWidgets.QApplication.translate("Form", "ColorChecker", None, -1))
        self.Texture_btn.setText(QtWidgets.QApplication.translate("Form", "Texture", None, -1))
        self.Flash_btn.setText(QtWidgets.QApplication.translate("Form", "Flash", None, -1))
        self.GreyChart_btn.setText(QtWidgets.QApplication.translate("Form", "Grey Chart", None, -1))

