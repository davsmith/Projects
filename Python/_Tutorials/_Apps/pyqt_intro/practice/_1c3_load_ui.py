"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=410)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates using files from the Qt Designer

Concepts introduced:
  - Using the .loadUi method from the uic class to load a .ui file
  - Autogenerating a .py file from the .ui file (window_gen.py)

Modules used:
  - PyQt6.QtWidgets
  - PyQt6.uic

Classes and methods used:
  - QApplication        Creates the application shell
      .exec

  - QWidget             Creates a form UI as part of a QApplication
      .show

  - QIcon               Helper class to create a normalized icon from a file (.png)

  - uic.loadUi          Defines the UI for the widget from the Qt Designer created .ui file

Created November 12, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6 import uic
from PyQt6 import QtWidgets
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT.

    This example demonstrates using the loadUi method on the uic class to
    display a .ui file

    Attributes:
      none
    """
    def __init__(self):
      """Extends the init function on the QWidget class.

      Loads the window UI from a .ui file created in Qt Designer

      Args:
        none

      Returns:
        none

      Raises:
        none
    """
      super().__init__()

      uic.loadUi("_1c1_window.ui", self)


# *** Main ***
if __name__ == '__main__':
  app = QApplication(sys.argv)
  window = Window()
  window.show()
  sys.exit(app.exec())