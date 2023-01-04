from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QRect, QCoreApplication, QMetaObject
from PyQt5.QtGui import QFont, QCloseEvent
from PyQt5.QtWidgets import QMessageBox, QLabel, QFrame, QTextEdit, QCheckBox, QSpinBox, QWidget, QLineEdit, \
    QPushButton, QMenuBar, QDialog, QAction


class pop_up(object):

    def error(self, title, text):
        error = QMessageBox()
        error.setWindowTitle(title)
        error.setText(text)
        error.setIcon(QMessageBox.Critical)
        error.setStandardButtons(QMessageBox.Ok)
        error.exec_()

    def information(self, title, text):
        info = QMessageBox()
        info.setIcon(QMessageBox.Information)
        info.setWindowTitle(title)
        info.setText(text)
        info.setStandardButtons(QMessageBox.Ok)
        info.exec_()


def button(name, text, centralwidget, x, y, width, height, action, enabled=True):
    push_button = QtWidgets.QPushButton(centralwidget)
    push_button.setGeometry(QtCore.QRect(x, y, width, height))
    push_button.setObjectName(str(name))
    push_button.setText(text)
    push_button.clicked.connect(action)
    push_button.setEnabled(enabled)


def text_Edit(name, centralwidget, x, y, width, height, enabled=True):
    text_Edit = QTextEdit(centralwidget)
    text_Edit.setObjectName(str(name))
    text_Edit.setGeometry(QRect(x, y, width, height))
    text_Edit.setEnabled(enabled)


def label(name, text, font: QFont(), centralwidget, x, y, width, height, fontweight=75 ,bold=False, enabled=True):
    label = QLabel(centralwidget)
    label.setText(text)
    label.setObjectName(str(name))
    label.setGeometry(QRect(x, y, width, height))
    font.setBold(bold)
    font.setWeight(fontweight)
    label.setFont(font)
    label.setEnabled(enabled)