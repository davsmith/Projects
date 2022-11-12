"""Example code from the PyQt Course for Beginners tutorial (YouTube).
------------------------------------------------------------------------
PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=4681)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates instantiation and use of a SpinBox control

Concepts introduced:
    - Use of QSpinBox

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QLabel          See detailed notes (linked above)
    QIcon           See detailed notes (linked above)
    QFont           See detailed notes (linked above)
    QHBoxLayout     See detailed notes (linked above)

    QSpinBox        Creates a spinner control with integer values (also see QDoubleSpinBox)

___________________________________________________________________________________________________

Created November 20, 2022
"""

from PyQt6.QtWidgets import QApplication, QWidget, QSpinBox, QLineEdit, QLabel, QHBoxLayout
from PyQt6.QtGui import QIcon, QFont
from os.path import join, realpath, dirname
import sys

class Window(QWidget):
    """Single line summary of the class with a period at the end.
 
    Methods:
        __init__(self, arg1, arg2)
        method2()
    """

    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, setting up icons, dimensions, etc.
      """
      super().__init__()

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))

      # Set window attributes
      self.setWindowTitle("PyQt6 SpinBox (QSpinBox)")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 500, 400)

      self.create_widgets()

    def create_widgets(self):
      """Creates the controls for this example"""
      hbox = QHBoxLayout()

      self.line_edit = QLineEdit()
      label = QLabel("Laptop price: ")
      self.spinbox = QSpinBox()
      self.spinbox.valueChanged.connect(self.on_value_changed)

      self.result = QLineEdit()

      hbox.addWidget(label)
      hbox.addWidget(self.line_edit)
      hbox.addWidget(self.spinbox)
      hbox.addWidget(self.result)

      self.setLayout(hbox)

    def on_value_changed(self):
      if self.line_edit.text() != 0:
        price = int(self.line_edit.text())

        total_price = self.spinbox.value() * price
        self.result.setText(str(total_price))


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())