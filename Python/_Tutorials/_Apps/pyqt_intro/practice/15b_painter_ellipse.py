"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=10808)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates drawing an ellipse using the QPainter class

Concepts introduced:
    - Drawing an ellipse using the drawEllipse method on QPainter

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QIcon           See detailed notes (linked above)

    QPainter        Implements paintEvent and provides tools to draw shapes
    QPen            Object to define colors, width, linestyle, etc. for borders
    QBrush          Object to define colors, width, fill color, etc. for fills

Created November 20, 2022
"""

from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtGui import QIcon, QPainter, QPen, QBrush
from PyQt6.QtCore import Qt

from os.path import join, realpath, dirname
import sys

class Window(QWidget):
    """The top level window for the application.
 
    Methods:
        __init__(self)
        paintEvent(self, e)
    """

    def __init__(self):
      """This method sets the window appearance."""
      super().__init__()

      script_path = realpath(dirname(__file__))
      app_icon = QIcon(join(script_path, "images/app_icon.png"))

      # Set window attributes
      self.setWindowTitle("PyQt6 QPainter Ellipse")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 600, 500)
    #   self.setStyleSheet("background-color:blue")

    #   self.create_widgets()

    def paintEvent(self, e):
        """This method uses the QPainter class to draw shapes.

        Input Arguments: self, e
        """
    
        painter = QPainter(self)
        painter.setPen(QPen(Qt.GlobalColor.yellow, 5, Qt.PenStyle.SolidLine))
        painter.setBrush(QBrush(Qt.GlobalColor.green, Qt.BrushStyle.Dense1Pattern))

        painter.drawEllipse(100, 100, 400, 200)


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())