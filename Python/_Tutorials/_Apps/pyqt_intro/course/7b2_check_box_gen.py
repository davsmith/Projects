"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=4526)
Detailed notes: https://tinyurl.com/y5db3c5c

Created by converting the .ui file from Designer into a .py file using
    pyuic6 -x <input file>.ui -o <output file>.py

Created November 23, 2022
"""
from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 300)
        self.verticalLayout = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.checkBox = QtWidgets.QCheckBox(Form)
        self.checkBox.setObjectName("checkBox")
        self.checkBox.toggled.connect(self.item_selected)

        self.horizontalLayout.addWidget(self.checkBox)
        self.checkBox_2 = QtWidgets.QCheckBox(Form)
        self.checkBox_2.setObjectName("checkBox_2")
        self.checkBox_2.toggled.connect(self.item_selected)
        self.horizontalLayout.addWidget(self.checkBox_2)
        self.checkBox_3 = QtWidgets.QCheckBox(Form)
        self.checkBox_3.setObjectName("checkBox_3")
        self.checkBox_3.toggled.connect(self.item_selected)

        self.horizontalLayout.addWidget(self.checkBox_3)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.label = QtWidgets.QLabel(Form)
        self.label.setStyleSheet("QLabel {\n"
"color:\'red\'\n"
"}")
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


    def item_selected(self):
      value = ""

      if self.checkBox.isChecked():
        value = self.checkBox.text()

      if self.checkBox_2.isChecked():
        value = self.checkBox_2.text()

      if self.checkBox_3.isChecked():
        value = self.checkBox_3.text()

      self.label.setText("You have selected: " + value)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.checkBox.setText(_translate("Form", "Python"))
        self.checkBox_2.setText(_translate("Form", "Java"))
        self.checkBox_3.setText(_translate("Form", "C++"))
        self.label.setText(_translate("Form", "TextLabel"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec())
