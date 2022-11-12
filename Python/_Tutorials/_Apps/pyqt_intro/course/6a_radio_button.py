"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=3256)
Detailed notes: https://tinyurl.com/y5db3c5c

Demonstrates radio button controls (QRadioButton)

Concepts introduced:
    - Instantiating radio button controls with QRadioButton
    - Grouping radio buttons in a named frame with QGroupBox
    - Responding to a change of button state with a 'toggled' event
    - Using the sender method to determine which control fire an event

Module/Class/Methods:
    QApplication    See detailed notes (linked above)
    QWidget         See detailed notes (linked above)
      sender        The object which fired the event (signal)

    QLabel          See detailed notes (linked above)
    QIcon           See detailed notes (linked above)
    QFont           See detailed notes (linked above)

    QGroupBox       Container for other controls (e.g. RadioButtons).  Draws a labelled frame around them.
    QRadioButton    
      text          The text label of the radio button
      isChecked     A bool indicating whether the radio button is selected/checked
      toggled       A signal indicating a radio button changed state (from on to off, or off to on)

Created November 20, 2022
"""

' Minimal PyQt6 window with no icon or title, using classes '
from PyQt6.QtWidgets import QApplication, QWidget, QGroupBox, QLabel, QRadioButton, QVBoxLayout, QHBoxLayout
from PyQt6.QtGui import QIcon, QFont
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

      # Set window attributes
      self.setWindowTitle("PyQt6 Radio Buttons")
      self.setWindowIcon(app_icon)
      self.setGeometry(500, 300, 400, 300)

      self.create_widgets()

    def create_widgets(self):
      self.label = QLabel("Hello")
      self.label.setFont(QFont("Times New Roman", 14))

      self.grpbox = QGroupBox("Choose Programming Language", self)
      self.grpbox.setFont(QFont("Times New Roman", 15))

      vbox = QVBoxLayout()
      vbox.addWidget(self.grpbox)
      vbox.addWidget(self.label)

      self.setLayout(vbox)

      hbox = QHBoxLayout()
      self.rad1 = QRadioButton("Python")
      self.rad1.setFont(QFont("Times New Roman", 14))
      self.rad1.toggled.connect(self.on_toggled)
      hbox.addWidget(self.rad1)

      self.rad2 = QRadioButton("C++")
      self.rad2.setFont(QFont("Times New Roman", 14))
      self.rad2.toggled.connect(self.on_toggled)
      hbox.addWidget(self.rad2)

      self.rad3 = QRadioButton("Java")
      self.rad3.setFont(QFont("Times New Roman", 14))
      self.rad3.toggled.connect(self.on_toggled)
      hbox.addWidget(self.rad3)

      self.rad4 = QRadioButton("C#")
      self.rad4.setFont(QFont("Times New Roman", 14))
      self.rad4.toggled.connect(self.on_toggled)
      hbox.addWidget(self.rad4)

      self.grpbox.setLayout(hbox)

    def on_toggled(self):
      radio = self.sender()

      if radio.isChecked():
        self.label.setText("You have selected: " + radio.text())


# *** Main ***
if __name__ == '__main__':
  app = QApplication([])
  window = Window()
  window.show()
  sys.exit(app.exec())