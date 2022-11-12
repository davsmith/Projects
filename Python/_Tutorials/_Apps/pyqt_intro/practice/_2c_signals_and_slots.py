"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1638)
Detailed notes: https://tinyurl.com/y5db3c5c

Example implementation of signals and slots (events and handlers)

Concepts introduced:
  - Using signals and slots to change the text on a label when a button is clicked
  - Storing controls as instance variables for the class

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
      .clicked.connect

  - QLabel              Creates a text label control
      .move
      .setStyleSheet
      .setFont
      .setText

  - QIcon               Helper class to create a normalized icon from a file (.png)
  - QFont               Helper class to load a font and set its size

Created November 13, 2022
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
      self.setWindowTitle('PyQt6 Signals & Slots')
      self.setWindowIcon(QIcon('app_icon.png'))
      self.setGeometry(500, 300, 400, 300)

      self.create_widgets()

    def create_widgets(self):
      # Create and position a button, and add an icon
      btn = QPushButton("Click Me", self)

      # btn.move(100,100)
      btn.setGeometry(100, 100, 100, 100)
      btn.setStyleSheet('background-color:red')
      btn.setIcon(QIcon('button_icon.png'))
      btn.clicked.connect(self.clicked_button)
      btn.pressed.connect(self.pressed_button)
      btn.released.connect(self.released_button)

      # Create a text label and set attributes, including font
      self.label = QLabel('My Label', self)
      self.label.move(100,200)
      self.label.setGeometry(100, 220, 200, 100)
      self.label.setStyleSheet('color:green')
      self.label.setFont(QFont('Times New Roman', 15))

    def clicked_button(self):
      self.label.setText("Text has changed")
      self.label.setStyleSheet('background-color:red')

    def pressed_button(self):
      print("The button was pressed")

    def released_button(self):
      print("The button was released")


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())