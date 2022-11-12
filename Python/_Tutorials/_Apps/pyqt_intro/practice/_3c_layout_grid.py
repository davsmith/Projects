"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=2510)
Detailed notes: https://tinyurl.com/y5db3c5c

PyQt app for use with layout management example

Concepts introduced:
  - Layout classes (QGridLayout)

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

  - QGridLayout         Creates a HBox layout container
      .addWidget

  - QIcon               Helper class to create a normalized icon from a file (.png)
Created November 13, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton
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

      self.setWindowTitle("PyQt6 Grid Layout")
      self.setWindowIcon(QIcon("app_icon.png"))
      self.setGeometry(500, 300, 400, 300)

      grid = QGridLayout()

      btn1 = QPushButton("Button One")
      btn2 = QPushButton("Button Two")
      btn3 = QPushButton("Button Three")
      btn4 = QPushButton("Button Four")
      btn5 = QPushButton("Button Five")
      btn6 = QPushButton("Button Six")
      btn7 = QPushButton("Button Seven")
      btn8 = QPushButton("Button Eight")

      grid.addWidget(btn1, 0, 0)
      grid.addWidget(btn2, 0, 1)
      grid.addWidget(btn3, 0, 2)
      grid.addWidget(btn4, 0, 3)
      grid.addWidget(btn5, 1, 0)
      grid.addWidget(btn6, 1, 1)
      grid.addWidget(btn7, 1, 2)
      grid.addWidget(btn8, 1, 3)

      self.setLayout(grid)

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())