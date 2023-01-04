from PyQt5 import QtCore, QtGui, QtWidgets
import sys

from PyQt5.QtWidgets import QTableWidgetItem
from openpyxl import load_workbook
import lib_gui

filename = 'assets\listado_pacientes.xlsx'

lista_map = ['Natalia', 'Leoncio Bernal', 'Maria Remesal', 'Cristina Plasencia', 'JA Fuentes', 'Susana Verdiu',
             'M Angeles Perez', \
             'Carmen Ballesteros', 'JA Peromingo', 'Rosario Gonzalez', 'Melisa Suelen']


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(771, 560)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.listView = QtWidgets.QTableWidget(self.centralwidget)
        self.listView.setGeometry(QtCore.QRect(30, 110, 362, 382))
        self.listView.setObjectName("listView")

        self.listView.setColumnCount(3)
        self.listView.setRowCount(self.extract_cells())
        item = QtWidgets.QTableWidgetItem('NOMBRE')
        self.listView.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem('FN')
        self.listView.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem('MAP')
        self.listView.setHorizontalHeaderItem(2, item)

        self.load_residents()

        self.label_name = QtWidgets.QLabel(self.centralwidget)
        self.label_name.setGeometry(QtCore.QRect(441, 286, 201, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_name.setFont(font)
        self.label_name.setObjectName("label_name")
        self.label_name.setVisible(False)

        self.textEdit_name = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_name.setGeometry(QtCore.QRect(441, 312, 251, 31))
        self.textEdit_name.setObjectName("textEdit_name")
        self.textEdit_name.setVisible(False)

        self.label_fn = QtWidgets.QLabel(self.centralwidget)
        self.label_fn.setGeometry(QtCore.QRect(441, 352, 251, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_fn.setFont(font)
        self.label_fn.setObjectName("label_fn")
        self.label_fn.setVisible(False)

        self.textEdit_fn = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_fn.setGeometry(QtCore.QRect(441, 378, 251, 31))
        self.textEdit_fn.setObjectName("textEdit_fn")
        self.textEdit_fn.setVisible(False)

        self.label_map = QtWidgets.QLabel(self.centralwidget)
        self.label_map.setGeometry(QtCore.QRect(441, 418, 251, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_map.setFont(font)
        self.label_map.setObjectName("label_map")
        self.label_map.setVisible(False)

        self.combobox_map = QtWidgets.QComboBox(self.centralwidget)
        self.combobox_map.setGeometry(QtCore.QRect(441, 444, 251, 31))
        self.combobox_map.setObjectName("textEdit_map")
        self.combobox_map.setVisible(False)
        for map in lista_map:
            self.combobox_map.addItem(map)

        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(30, 70, 251, 31))
        self.textEdit.setObjectName("textEdit")

        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setGeometry(QtCore.QRect(441, 110, 311, 81))
        self.textEdit_2.setObjectName("textEdit_2")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(441, 80, 201, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(30, 40, 131, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(30, 500, 121, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.add_resident)

        self.pushButton_add = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_add.setGeometry(QtCore.QRect(441, 500, 121, 41))
        self.pushButton_add.setObjectName("pushButton_add")
        self.pushButton_add.setVisible(False)
        self.pushButton_add.clicked.connect(self.confirm)

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(160, 500, 121, 41))
        self.pushButton_2.setObjectName("pushButton_2")

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(632, 200, 121, 41))
        self.pushButton_3.setObjectName("pushButton_3")

        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(441, 200, 121, 41))
        self.pushButton_4.setObjectName("pushButton_4")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 771, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def extract_cells(self):
        self.spreadsheet = load_workbook(filename)
        self.sheet = self.spreadsheet.active
        rows = self.sheet.max_row
        self.spreadsheet.save(filename)
        return rows

    def load_residents(self):
        self.spreadsheet = load_workbook(filename)
        self.sheet = self.spreadsheet.active
        a = 0
        for number in range(self.sheet.max_row):
            b = a + 2
            self.listView.setItem(a, 0, QtWidgets.QTableWidgetItem(str(self.sheet['A' + str(b)].value)))
            self.listView.setItem(a, 1, QtWidgets.QTableWidgetItem(
                self.format_date_of_birth(str(self.sheet['B' + str(b)].value))))
            self.listView.setItem(a, 2, QtWidgets.QTableWidgetItem(str(self.sheet['C' + str(b)].value)))
            a += 1

    def add_resident(self):
        self.label_name.setVisible(True)
        self.textEdit_name.setVisible(True)
        self.label_fn.setVisible(True)
        self.textEdit_fn.setVisible(True)
        self.label_map.setVisible(True)
        self.combobox_map.setVisible(True)
        self.pushButton_add.setVisible(True)

        # todo: escribir la logica para añadir residente con libreria openpyxl

    def confirm(self):
        self.sheet['A' + str(self.sheet.max_row + 1)] = self.textEdit_name.toPlainText()
        self.sheet['B' + str(self.sheet.max_row)] = self.textEdit_fn.toPlainText()
        self.sheet['C' + str(self.sheet.max_row)] = self.combobox_map.currentText()
        self.spreadsheet.save(filename)
        self.load_residents()

    def format_date_of_birth(self, date):
        self.format1 = date.split(' ')[0]
        self.format2 = self.format1.split('-')
        return '/'.join(self.format2[::-1])

    def remove_resident(self):
        pass
        # todo: escribir la logica para eliminar residente con libreria openpyxl, mover celdas de abajo hacia arriba

    def request(self):
        pass
        # todo: escribir logica para escribir en archivo doc con libreria docx y docx.share

    def reset_week(self):
        pass
        # todo: escribir logica para cambiar nombre y ruta de archivos con libreria sys y os

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SOLICITUDES A MAP"))
        self.textEdit.setHtml(_translate("MainWindow",
                                         "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                         "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                         "p, li { white-space: pre-wrap; }\n"
                                         "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                         "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Buscar</p></body></html>"))
        self.label.setText(_translate("MainWindow", "Escriba la solicitud a continuacion:"))
        self.label_2.setText(_translate("MainWindow", "Listado de residentes:"))
        self.label_name.setText(_translate("MainWindow", "Nombre y Apellidos:"))
        self.label_fn.setText(_translate("MainWindow", "Fecha de Nacimiento:"))
        self.label_map.setText(_translate("MainWindow", "MAP:"))
        self.pushButton.setText(_translate("MainWindow", "Añadir"))
        self.pushButton_2.setText(_translate("MainWindow", "Eliminar"))
        self.pushButton_3.setText(_translate("MainWindow", "SOLICITAR"))
        self.pushButton_4.setText(_translate("MainWindow", "RESETEAR\n"
                                                           "SEMANA"))
        self.pushButton_add.setText(_translate("MainWindow", "CONFIRMAR"))


def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
