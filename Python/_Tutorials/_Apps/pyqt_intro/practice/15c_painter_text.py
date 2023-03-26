"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=10973)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates drawing text on a canvas using the QPainter class

Concepts introduced:
    - Drawing text using the drawText method on QPainter
    - Drawing text with HTML formatting using QTextDocument

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
    QIcon           See detailed notes (linked above)

    QPainter        Implements paintEvent and provides tools to draw shapes
    QPen            Object to define colors, width, linestyle, etc. for borders
    QBrush          Object to define colors, width, fill color, etc. for fills

    QTextDocument
      setTextWidth  Sets the window width in which the text is drawn
      setHtml       Sets the contents and HTML formatting for the string
      drawContents  Renders the HTML string on the canvas

Created November 23, 2022
"""

from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtGui import QIcon, QPainter, QPen, QBrush, QTextDocument
from PyQt6.QtCore import Qt, QRect, QRectF

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
      self.setWindowTitle("PyQt6 QPainter Text")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 600, 500)
    #   self.setStyleSheet("background-color:blue")

    def paintEvent(self, e):
        """This method uses the QPainter class to draw text.

        Input Arguments: self, e
        """
    
        painter = QPainter(self)
        painter.setPen(QPen(Qt.GlobalColor.blue, 5, Qt.PenStyle.SolidLine))
        painter.setBrush(QBrush(Qt.GlobalColor.green, Qt.BrushStyle.Dense1Pattern))

        painter.drawText(100, 100, "PyQt Course")

        rect = QRect(100, 150, 250, 25) # Left, top, width and height
        painter.drawRect(rect)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "PyQt Course")

        htmlString = "<b>Welcome to the PyQt6 tutorial</b> - \
            <i>November 23, 2022</i> \
            <font size='15' color='red'> \
            <br>Enjoy the course!"

        document = QTextDocument()
        rect2 = QRectF(0, 0, 600, 250)
        document.setTextWidth(rect2.width())
        document.setHtml(htmlString)
        document.drawContents(painter, rect2)

# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())