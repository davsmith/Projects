"""Example code from the PyQt Course for Beginners tutorial (YouTube).
------------------------------------------------------------------------
PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=5461)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates QTableWidget through code

Concepts introduced:
    - Create and style a table via the QTableWidget

Module/Class/Methods:
    QApplication        See detailed notes (linked above)
    QWidget             See detailed notes (linked above)
    QLabel              See detailed notes (linked above)
    QIcon               See detailed notes (linked above)
    QFont               See detailed notes (linked above)

    QTableWidget        Instantiates a table control
      setColumnCount    Sets the width of the table in columns
      setRowCount       Sets the height of the table in rows
      setItem           Sets the content at a cell location
      setStyleSheet     Defines the appearance of the table
      setFont           Defines the font and font size for text in the table

    QTableWidgetItem
      __init__          

___________________________________________________________________________________________________

Created November 20, 2022
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

        self.create_widgets()

    def create_widgets(self):
        vbox = QVBoxLayout()

        # add a table
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())