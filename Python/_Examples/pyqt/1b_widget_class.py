"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=410)
Detailed notes: https://tinyurl.com/y5db3c5c

Minimal PyQt app with icon and title, built from/as a class

Concepts introduced:
  - Converting the app to object oriented
  - Extending the QWidget class
  - Setting the window icon and title
  - Using the QIcon class
  - Defining fixed window dimensions
  - Defining adjustable window dimenstions
  - Setting window parameters with a stylesheet

Created November 12, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget
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

      self.setWindowTitle("PyQt6 Window")

      # The icon in the example is 426x480
      self.setWindowIcon(QIcon("app_icon.png"))

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