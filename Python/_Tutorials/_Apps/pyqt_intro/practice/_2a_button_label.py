"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1326)
Detailed notes: https://tinyurl.com/y5db3c5c

Example implementation of a button and label using PyQt

Concepts introduced:
  - Using the QPushButton and QLabel classes from the PyQt6.QtWidgets module
  - Using the QFont class from QtGui to set attributes on the text of a label
  - The button is displayed but no events are connected in this example

Classes and methods used:
  - QApplication        Creates the application shell
      .exec

  - QWidget             Creates a form UI as part of a QApplication
      .setWindowTitle 
      .setWindowIcon
      .setGeometry
      .show

  - QPushButton         Creates a push button control
      .setGeometry
      .setStyleSheet
      .setIcon

  - QLabel              Creates a text label control
      .move
      .setStyleSheet
      .setFont

  - QIcon               Helper class to create a normalized icon from a file (.png)
  - QFont               Helper class to load a font and set its size

Created November 12, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLabel
from PyQt6.QtGui import QIcon, QFont
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

      # Set window attributes
      self.setWindowTitle("PyQt6 Button & Label")
      self.setWindowIcon(QIcon("app_icon.png"))
      self.setGeometry(500, 300, 400, 300)

      self.create_widgets()

    def create_widgets(self):
      # Create and position a button, and add an icon
      btn = QPushButton("Click Me", self)

      # btn.move(100,100)
      btn.setGeometry(100, 100, 100, 100)
      btn.setStyleSheet('background-color:red')
      btn.setIcon(QIcon('button_icon.png'))

      # Create a text label and set attributes, including font
      label = QLabel("My Label", self)
      label.move(100,200)
      label.setStyleSheet('color:green')
      label.setFont(QFont("Times New Roman", 15))

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())