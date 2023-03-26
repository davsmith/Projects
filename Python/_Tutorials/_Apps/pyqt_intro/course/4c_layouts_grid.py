"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=2510)
Detailed notes: https://tinyurl.com/y5db3c5c

Concepts introduced:
  - Layout classes (QGridLayout)

Module/Class/Methods:
    QApplication    Provides interface with OS for command line args and exit codes
        __init__    Pass sys.argv or []
        exec        Runs the application and returns exit code

    QWidget         Base application form with no menus
        __init__
        show        Displays the form

    QIcon           Loads an image file and provides an instance of a PyQt icon class
        __init__    Receives an image file as an argument (e.g. .png)

    QFont           Loads a font, sets the font size and provides an instance of a PyQt font class
    QGridLayout     Grid layout control


Created November 20, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton
from PyQt6.QtGui import QIcon
from os.path import join, realpath, dirname
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT."""
    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, setting up icons, dimensions, etc.
    """
      super().__init__()

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))

      # Define window appearance/attributes
      self.setWindowTitle("PyQt6 Grid Layout")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 400, 300)

      self.create_widgets()

    def create_widgets(self):
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