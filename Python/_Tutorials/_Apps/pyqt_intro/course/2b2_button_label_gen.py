"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=2045)
Detailed notes: https://tinyurl.com/y5db3c5c

Created by converting the .ui file from Designer into a .py file using
    pyuic6 -x <input file>.ui -o <output file>.py

Concepts introduced:
    - Adding signals and slots (via code)

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)

    QPushButton     Class to instantiate a push button control
      __init__      Receives the button text and host control
      clicked       Set the event handler for the click event (using .connect)

    QLabel          Class to instantiate a text control
      __init__      Receives the label text and host control
      setText       Sets the text string on the label
      setStyleSheet Formats the appearance of the label (background-color)
    
    QFont           Loads a font, sets the font size and provides an instance of a PyQt font class
    QIcon           See detailed notes (linked above)

Created November 12, 2022
"""
from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        Form.setStyleSheet("QPushButton {\n"
            "    background-color: \'red\'\n"
            "}\n"
            "\n"
            "QLabel {\n"
            "   color: \'green\'\n"
            "}\n"
        "")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(130, 20, 100, 100))
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(0, 200, 391, 50))
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(16)
        font.setBold(True)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.pushButton.clicked.connect(self.clicked_button)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def clicked_button(self):
        self.label.setText("The button was clicked!")
        self.label.setStyleSheet('background-color:yellow')
    
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pushButton.setText(_translate("Form", "Click Me"))
        self.label.setText(_translate("Form", "TextLabel"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())
