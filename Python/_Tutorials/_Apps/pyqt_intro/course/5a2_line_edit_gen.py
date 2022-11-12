"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=3058)
Detailed notes: https://tinyurl.com/y5db3c5c

Created by converting the .ui file from Designer into a .py file using
    pyuic6 -x <input file>.ui -o <output file>.py

Concepts introduced:
    - Adding edit box controls (QLineEdit)

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)

    QVBoxLayout     See detailed notes (linked above)
    QHBoxLayout     See detailed notes (linked above)
    QPushButton     See detailed notes (linked above)

    QLineEdit       Class to instantiate an edit box control
      __init__      Receives the button text and host control
      setObjectName Sets the name of the object used to reference the control
      text          Retrieves the string from the edit control
      setText       Sets the string in the edit control
    
    QFont           Loads a font, sets the font size and provides an instance of a PyQt font class
    QIcon           See detailed notes (linked above)

Created November 12, 2022
"""
from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setObjectName("lineEdit")
        self.verticalLayout.addWidget(self.lineEdit)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setObjectName("pushButton")

        self.pushButton.clicked.connect(self.on_clicked)

        self.horizontalLayout.addWidget(self.pushButton)
        self.label = QtWidgets.QLabel(Form)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.verticalLayout_2.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


    def on_clicked(self):
        textInput = self.lineEdit.text()
        self.label.setText(textInput)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pushButton.setText(_translate("Form", "Click"))
        self.label.setText(_translate("Form", ""))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())
