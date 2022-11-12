"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=2220)
Detailed notes: https://tinyurl.com/y5db3c5c

PyQt app for use with layout management example

Concepts introduced:
  - Layout classes (QVBoxLayout)

Classes and methods used:
  - QApplication        Creates the application shell
      .exec

  - QWidget             Creates a form UI as part of a QApplication
      .setWindowTitle 
      .setWindowIcon
      .setGeometry
      .show
      .setLayout

  - QPushButton         Creates a push button control

  - QLabel              Creates a text label control
      .move
      .setStyleSheet
      .setFont

  - QVBoxLayout         Creates a VBox layout container
      .addWidget

  - QIcon               Helper class to create a normalized icon from a file (.png)

Created November 13, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton
from PyQt6.QtGui import QIcon
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT.

    Longer class information...
    Longer class information...

    Attributes:
      none
    """
    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, setting up icons, dimensions, etc.

      Args:
        none

      Returns:
        none

      Raises:
        none
    """
      super().__init__()

      self.setWindowTitle("PyQt6 Layouts")
      self.setWindowIcon(QIcon("app_icon.png"))
      self.setGeometry(500, 300, 400, 300)

      vbox = QVBoxLayout()

      btn1 = QPushButton("Button One")
      btn2 = QPushButton("Button Two")
      btn3 = QPushButton("Button Three")
      btn4 = QPushButton("Button Four")

      vbox.addWidget(btn1)
      vbox.addWidget(btn2)
      vbox.addWidget(btn3)
      vbox.addWidget(btn4)

      self.setLayout(vbox)

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())