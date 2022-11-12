"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1058)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates loading files at runtime from the Qt Designer

Concepts introduced:
  - Using the .loadUi method from the uic class to load a .ui file

Module/Class/Methods:
    QApplication    Provides interface with OS for command line args and exit codes
        __init__    Pass sys.argv or []
        exec        Runs the application and returns exit code

    QWidget         Base application form with no menus
        __init__
        show        Displays the form

    QIcon           Loads an image file and provides an instance of a PyQt icon class
        __init__    Receives an image file as an argument (e.g. .png)

    uic.loadUi      Loads a .ui file (from Qt Designer) and assigns it to the UI of
                    the specified form

Created November 20, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtGui import QIcon
from PyQt6 import uic
from os.path import join, realpath, dirname
import sys

class UI(QWidget):
    """Basic app window class extending the QWidget base class from PyQT.

    This example demonstrates using the loadUi method on the uic class to
    display a .ui file
    """
    def __init__(self):
      """Extends the init function on the QWidget class."""
      super().__init__()

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))
      app_ui = join(script_path, "1c1_widget.ui")

      uic.loadUi(app_ui, self)
      self.setWindowIcon(app_icon)


# *** Main ***
if __name__ == '__main__':
  app = QApplication(sys.argv)
  window = UI()
  window.show()
  sys.exit(app.exec())