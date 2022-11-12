"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=1638)
Detailed notes: https://tinyurl.com/y5db3c5c

Example implementation of signals and slots (events and handlers)

Concepts introduced:
  - Using signals and slots to change the text on a label when a button is clicked
  - Storing controls as instance variables for the class
  - Dave's convention is to name slot functions (callbacks) as on_<event_name> (e.g. on_clicked)

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)

    QPushButton     Class to instantiate a push button control
      __init__      Receives the button text and host control
      clicked       Set the event handler for the click event (using .connect)

    QLabel          Class to instantiate a text control
      __init__      Receives the label text and host control
      setText       Sets the text string on the label
      setStyleSheet Formats the appearance of the label (background-color)
    
    QFont           Loads a font, sets the font size and provides an instance of a PyQt font class
    QIcon           See detailed notes (linked above)

Created November 12, 2022


Created November 13, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLabel
from PyQt6.QtGui import QIcon, QFont
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

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))

      # Set window attributes
      self.setWindowTitle("PyQt6 Signals & Slots")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 400, 300)
      self.setStyleSheet("background-color:blue")

      self.create_widgets()

    def create_widgets(self):
      # Create and position a button, and add an icon
      btn = QPushButton("Click Me", self)

      btn.setGeometry(100, 100, 100, 100)
      btn.setStyleSheet('background-color:red')
      btn.setIcon(QIcon('button_icon.png'))
      btn.clicked.connect(self.on_clicked)

      # Create a text label and set attributes, including font
      self.label = QLabel('My Label', self)
      self.label.move(100,200)
      self.label.setGeometry(100, 220, 200, 100)
      self.label.setStyleSheet('color:green')
      self.label.setFont(QFont('Times New Roman', 15))

    def on_clicked(self):
      self.label.setText("Text has changed")
      self.label.setStyleSheet('background-color:red')

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())