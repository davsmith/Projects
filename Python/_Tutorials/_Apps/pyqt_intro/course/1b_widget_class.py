"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=410)
Detailed notes: https://tinyurl.com/y5db3c5c

A basic PyQt application built by extending the QWidget class
Alternative window types are QDialog and QMainWindow

Concepts introduced:
  - Converting the app to object oriented
  - Extending the QWidget class
  - Setting the window icon and title
  - Using the QIcon class
  - Defining fixed window dimensions
  - Defining adjustable window dimenstions
  - Setting window parameters with a stylesheet

Module/Class/Methods:
    QApplication        Provides interface with OS for command line args and exit codes
        __init__        Pass sys.argv or []
        exec            Runs the application and returns exit code

    QWidget             Base application form with no menus
        __init__
        show            Displays the form
        setFixedWidth   Sets the width of the window and disables resize
        setFixedHeight  Sets the height of the window and disables resize
        setGeometry     Sets the width, height and location of the window
                        Allows for resizing the window
        setStyleSheet   Sets the appearance (background color)of the app using CSS

    QIcon           Loads an image file and provides an instance of a PyQt icon class
        __init__    Receives an image file as an argument (e.g. .png)

Created November 12, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtGui import QIcon
from os.path import join, realpath, dirname
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT."""
    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, including icons, dimensions, etc.
      """

      super().__init__()

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))

      self.setWindowTitle("PyQt6 Window (QWidget)")
      self.setWindowIcon(app_icon)

      # Use setFixedHeight and setFixedWidth to define a non-resizable window
      # self.setFixedHeight(200)
      # self.setFixedWidth(200)

      # Use setGeometry to set the window size, but allow resizing
      self.setGeometry(500, 300, 400, 300)

      # Appearance is controlled through CSS, either by specifying individual
      # parameters, or multiple parameters in a tuple
      self.setStyleSheet("background-color:blue")

      # stylesheet = ('background-color:red')
      # self.setStyleSheet(stylesheet)

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())