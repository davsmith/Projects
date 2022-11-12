"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=9101)
Detailed notes: https://tinyurl.com/y5db3c5c

This is a placeholder file for the db_show_data code in the example.  The code is not correct.

Created by converting the .ui file from Designer into a .py file using
    pyuic6 -x <input file>.ui -o <output file>.py

Concepts introduced:

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QVBoxLayout
    QLabel
    QLineEdit
    QHBoxLayout

    QSplitter
    QSpacerItem


Created November 12, 2022
"""
import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout
from PyQt6.QtWidgets import QTableWidget,QTableWidgetItem
from PyQt6.QtGui import QIcon, QFont

class Window(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('PyQt6 TableWidget')
        self.setWindowIcon(QIcon("app_icon.png"))
        self.setGeometry(500, 200, 500, 400)

        vbox = QVBoxLayout()

        # add table 
        tableWidget = QTableWidget(self)
        tableWidget.setColumnCount(3)
        tableWidget.setRowCount(3)
        tableWidget.setStyleSheet('background-color:yellow')
        tableWidget.setFont(QFont('Times New Roman', 14))

        tableWidget.setItem(0, 0, QTableWidgetItem("FName"))
        tableWidget.setItem(0, 1, QTableWidgetItem("LName"))
        tableWidget.setItem(0, 2, QTableWidgetItem("Email"))

        tableWidget.setItem(1, 0, QTableWidgetItem("Parwiz"))
        tableWidget.setItem(1, 1, QTableWidgetItem("Forogh"))
        tableWidget.setItem(1, 2, QTableWidgetItem("parwiz@gmail.com"))

        tableWidget.setItem(2, 0, QTableWidgetItem("John"))
        tableWidget.setItem(2, 1, QTableWidgetItem("Doe"))
        tableWidget.setItem(2, 2, QTableWidgetItem("john_doe@gmail.com"))

        vbox.addWidget(tableWidget)
        self.setLayout(vbox)

        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = Window()
    sys.exit(app.exec())