"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=3820)
Detailed notes: https://tinyurl.com/y5db3c5c

Example implementation of a set of checkboxes using PyQt

Concepts introduced:
  - Using the QCheckBox class from the PyQt6.QtWidgets module

Created November 19, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QGroupBox, QLabel, QCheckBox, QVBoxLayout, QHBoxLayout
from PyQt6.QtGui import QIcon, QFont
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT."""

    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, setting up icons, dimensions, etc.
      """
      super().__init__()

      # Set window attributes
      self.setWindowTitle("PyQt6 CheckBox Example")
      self.setWindowIcon(QIcon("app_icon.png"))
      self.setGeometry(500, 200, 500, 400)

      self.create_widgets()

    def create_widgets(self):
      hbox = QHBoxLayout()

      self.check1 = QCheckBox("Python")
      self.check1.setFont(QFont("Times New Roman", 14))
      self.check1.toggled.connect(self.on_toggled)
      hbox.addWidget(self.check1)

      self.check2 = QCheckBox("Java")
      self.check2.setFont(QFont("Times New Roman", 14))
      self.check2.toggled.connect(self.on_toggled)
      hbox.addWidget(self.check2)

      self.check3 = QCheckBox("C#")
      self.check3.setFont(QFont("Times New Roman", 14))
      self.check3.toggled.connect(self.on_toggled)
      hbox.addWidget(self.check3)

      self.check4 = QCheckBox("C++")
      self.check4.setFont(QFont("Times New Roman", 14))
      self.check4.toggled.connect(self.on_toggled)
      hbox.addWidget(self.check4)

      vbox = QVBoxLayout()

      self.label = QLabel("Label")
      self.label.setFont(QFont("Sanserif", 15))
      self.label.setStyleSheet('color:red')

      vbox.addWidget(self.label)
      vbox.addLayout(hbox)

      self.setLayout(vbox)

    def on_toggled(self):
      value = ""

      if self.check1.isChecked():
        value = self.check1.text()

      if self.check2.isChecked():
        value = self.check2.text()

      if self.check3.isChecked():
        value = self.check3.text()

      if self.check4.isChecked():
        value = self.check4.text()

      self.label.setText("You have selected: " + value)


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())