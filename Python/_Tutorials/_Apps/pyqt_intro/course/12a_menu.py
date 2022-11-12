"""
-   The menuBar attribute is part of the QMainWindow class


"""
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout
from PyQt6.QtWidgets import QTableWidget,QTableWidgetItem
from PyQt6.QtGui import QIcon, QFont, QAction

class Window(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt6 Menu")
        self.setWindowIcon(QIcon("app_icon.png"))
        self.setGeometry(500, 200, 500, 400)

        self.create_menu()

        vbox = QVBoxLayout()

        # vbox.addWidget(tableWidget)
        # self.setLayout(vbox)

        self.show()

    def create_menu(self):
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu("File")

        newAction = QAction(QIcon("images/new.png"), "New", self)
        newAction.setShortcut("Ctrl+N")
        fileMenu.addAction(newAction)

        saveAction = QAction(QIcon("images/save.png"), "Save", self)
        saveAction.setShortcut("Ctrl+S")
        fileMenu.addAction(saveAction)

        copyAction = QAction(QIcon("images/copy.png"), "Copy", self)
        copyAction.setShortcut("Ctrl+C")
        fileMenu.addAction(copyAction)

        fileMenu.addSeparator()


        exitAction = QAction(QIcon("images/exit.png"), "Exit", self)
        exitAction.setShortcut("Ctrl+Q")
        exitAction.triggered.connect(self.close_window)
        fileMenu.addAction(exitAction)

    def close_window(self):
        self.close()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = Window()
    sys.exit(app.exec())