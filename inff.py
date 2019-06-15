# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'inff.ui'
#
# Created by: PyQt5 UI code generator 5.10.1
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(272, 369)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName("textBrowser")
        self.verticalLayout.addWidget(self.textBrowser)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'.SF NS Text\'; font-size:13pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Добро пожаловать! Это программа для выявления благонадежности контрагентов. Для того, чтобы проверить компанию, введите название, ИНН или ОГРН в поля в центре экрана. Приложение произведет поиск по доступным в сети Интернет базам данным и выведет итоговую информацию. В появившемся после работы программы окне &quot;Результат&quot; можно будет увидеть финальную таблицу проверки на благонадежность. По нажатию на пиктограммы &quot;PDF&quot; и &quot;Excel&quot; итоговая сводка сохранится в соответствующий формат в корневой папке программы. Вопросы по работе приложения и предложения по исправлению ошибок предлагается отправлять по адресу <span style=\" font-family:\'Roboto,RobotoDraft,Helvetica,Arial,sans-serif\'; color:#222222; background-color:#ffffff;\">kontragentplus@gmail.com </span></p></body></html>"))

