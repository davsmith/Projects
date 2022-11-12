"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=10388)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates drawing a rectangle using the QPainter class

Concepts introduced:
    - Drawing rectangles with different fills and border styles
    - Differences between PyQt5 and PyQt6
        - The method of using enums changed (e.g. from Qt.SolidLine to Qt.PenStyle.SolidLine)
    - The paintEvent which occurs when the canvas is repainted (window resize, minimize/restore)
    - Using enums (colors, brush styles, etc.) defined in the Qt module from PyQt6.QtCore 

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QLabel          See detailed notes (linked above)
    QIcon           See detailed notes (linked above)
    QFont           See detailed notes (linked above)

    QPainter        Draws shapes in a window (canvas)
      paintEvent    A method called when the canvas is repainted
___________________________________________________________________________________________________

Created November 22, 2022
"""

from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtGui import QIcon, QFont, QPainter, QPen, QBrush
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
      self.setWindowTitle("PyQt6 QPainter Rectangle")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 500, 300)

    def paintEvent(self, e):
        painter = QPainter(self)
        # painter.setPen(QPen(Qt.GlobalColor.blue, 5, Qt.PenStyle.DashDotDotLine))
        painter.setPen(QPen(Qt.GlobalColor.blue, 5, Qt.PenStyle.SolidLine))
        # painter.setBrush(QBrush(Qt.GlobalColor.green, Qt.BrushStyle.SolidPattern))
        painter.setBrush(QBrush(Qt.GlobalColor.green, Qt.BrushStyle.DiagCrossPattern))

        painter.drawRect(100, 15, 300, 100)
        print("paintEvent occurred")


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())