# https://youtu.be/ot94H3-d5d8?t=5952

import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout
from PyQt6.QtWidgets import QListWidget, QLabel
from PyQt6.QtGui import QIcon, QFont

class Window(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('PyQt6 ListWidget')
        self.setWindowIcon(QIcon("app_icon.png"))
        self.setGeometry(500, 200, 500, 400)

        vbox = QVBoxLayout()

        self.listWidget = QListWidget()
        self.listWidget.insertItem(0, "PyQt6")
        self.listWidget.insertItem(1, "wxPython")
        self.listWidget.insertItem(2, "Kivy")
        self.listWidget.insertItem(3, "TKinter")
        self.listWidget.insertItem(4, "PySide2")
        self.listWidget.setStyleSheet('background-color:yellow')
        self.listWidget.clicked.connect(self.item_selected)

        self.label = QLabel("Hello")
        self.label.setFont(QFont("Sanserif", 14))

        vbox.addWidget(self.listWidget)
        vbox.addWidget(self.label)
        self.setLayout(vbox)

        self.show()

    def item_selected(self):
        item = self.listWidget.currentItem()
        self.label.setText(f"You have selected {item.text()}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = Window()
    sys.exit(app.exec())