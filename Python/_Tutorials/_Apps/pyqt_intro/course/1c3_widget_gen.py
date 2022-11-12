"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1217)
Detailed notes: https://tinyurl.com/y5db3c5c

Created by converting the .ui file from Designer into a .py file using
    pyuic6 -x <input file>.ui -o <output file>.py

Concepts introduced:
    - Converting a .ui file into a .py file using pyuic6

Module/Class/Methods:
    QApplication        Provides interface with OS for command line args and exit codes
        __init__        Pass sys.argv or []
        exec            Runs the application and returns exit code

    QWidget             Base application form with no menus
        __init__
        show            Displays the form
        setObjectName   Defines the name of the object used to reference the control
        resize          Sets the width and height of the window
        setStyleSheet   Sets the appearance (background color)of the app using CSS
        setWindowTitle  Sets the text to be displayed in the app title bar
        

    QIcon           Loads an image file and provides an instance of a PyQt icon class
        __init__    Receives an image file as an argument (e.g. .png)

Created November 12, 2022
"""

from PyQt6 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        Form.setStyleSheet("QWidget {\n"
            "background-color: rgb(67, 67, 255)\n"
        "}")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Qt Designer Window"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())