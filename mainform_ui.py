# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'd:\Документы\testgen2\mainform.ui'
#
# Created by: PyQt5 UI code generator 5.15.11
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(2000, 1000)
        MainWindow.setBaseSize(QtCore.QSize(2000, 800))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(10, 10, 781, 251))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.btn_loadfile = QtWidgets.QPushButton(self.centralwidget)
        self.btn_loadfile.setGeometry(QtCore.QRect(10, 270, 171, 41))
        self.btn_loadfile.setObjectName("btn_loadfile")
        self.btn_testgen = QtWidgets.QPushButton(self.centralwidget)
        self.btn_testgen.setGeometry(QtCore.QRect(10, 400, 161, 41))
        self.btn_testgen.setObjectName("btn_testgen")
        self.spinBox = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox.setGeometry(QtCore.QRect(130, 320, 61, 31))
        self.spinBox.setObjectName("spinBox")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 319, 121, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 360, 111, 31))
        self.label_2.setObjectName("label_2")
        self.spinBox_2 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_2.setGeometry(QtCore.QRect(130, 360, 61, 31))
        self.spinBox_2.setObjectName("spinBox_2")
        self.groupBox_col = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_col.setGeometry(QtCore.QRect(230, 310, 120, 80))
        self.groupBox_col.setObjectName("groupBox_col")
        self.radioButton_1col = QtWidgets.QRadioButton(self.groupBox_col)
        self.radioButton_1col.setGeometry(QtCore.QRect(10, 20, 83, 18))
        self.radioButton_1col.setObjectName("radioButton_1col")
        self.radioButton_2col = QtWidgets.QRadioButton(self.groupBox_col)
        self.radioButton_2col.setGeometry(QtCore.QRect(10, 50, 83, 18))
        self.radioButton_2col.setObjectName("radioButton_2col")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 2000, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action_load_file = QtWidgets.QAction(MainWindow)
        self.action_load_file.setObjectName("action_load_file")
        self.action_testgen = QtWidgets.QAction(MainWindow)
        self.action_testgen.setObjectName("action_testgen")

        self.retranslateUi(MainWindow)
        self.btn_loadfile.clicked.connect(self.action_load_file.trigger) # type: ignore
        self.btn_testgen.clicked.connect(self.action_testgen.trigger) # type: ignore
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "TESTGen2.0"))
        self.btn_loadfile.setText(_translate("MainWindow", "Загрузить из файла"))
        self.btn_testgen.setText(_translate("MainWindow", "Генерить тест"))
        self.label.setText(_translate("MainWindow", "Количество вариантов"))
        self.label_2.setText(_translate("MainWindow", "Количество вопросов"))
        self.groupBox_col.setTitle(_translate("MainWindow", "Количество колонок"))
        self.radioButton_1col.setText(_translate("MainWindow", "1 колонка"))
        self.radioButton_2col.setText(_translate("MainWindow", "2 колонки"))
        self.action_load_file.setText(_translate("MainWindow", "load_file"))
        self.action_testgen.setText(_translate("MainWindow", "testgen"))