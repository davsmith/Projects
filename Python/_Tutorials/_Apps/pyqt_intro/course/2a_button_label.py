"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1326)
Detailed notes: https://tinyurl.com/y5db3c5c

Example implementation of the button and label controls using PyQt
The controls are for display only, and do not respond to events

Concepts introduced:
  - Using the QPushButton and QLabel classes from the PyQt6.QtWidgets module
  - Using the QFont class from QtGui to set attributes on the text of a label
  - Creating a set up function (create_widgets) for instantiating the controls for the app

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QIcon           See detailed notes (linked above)

    QPushButton     Class to instantiate a push button control
      __init__      Receives the button text and host control
      setStyleSheet Sets the display attributes for the control based on CSS
      setGeometry   Sets the dimensions and location of the control
      setIcon       Sets an icon which appears next to the button label

    QLabel          Class to instantiate a text control
      __init__      Receives the label text and host control
      setStyleSheet Sets the display attributes for the control based on CSS
      move          Sets the position of the control
      setFont       Sets the font and font size for the control.  Receives a QFont object.
    
    QFont           Loads a font, sets the font size and provides an instance of a PyQt font class
    
Created November 12, 2022
"""

from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLabel
from PyQt6.QtGui import QFont, QIcon
from os.path import join, realpath, dirname
import sys

class Window(QWidget):
    """Basic app window class extending the QWidget base class from PyQT.

    Methods:
        __init__(self)
        create_widgets(self)
    """
    def __init__(self):
      """Extends the init function on the QWidget class.

      Sets up the application window, setting up icons, dimensions, etc.
      """
      super().__init__()

      self.script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(self.script_path, "images/app_icon.png"))

      # Set window attributes
      self.setWindowTitle("PyQt6 Button & Label")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 400, 300)

      self.create_widgets()

    def create_widgets(self):
      button_icon = QIcon(join(self.script_path, "images/button_icon.png"))

      # Create and position a button, and add an icon
      btn = QPushButton("Click Me", self)

      # btn.move(100,100)
      btn.setGeometry(100, 100, 100, 100)
      btn.setStyleSheet('background-color:red')
      btn.setIcon(button_icon)

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