"""Example code from PyQt Course for Beginners tutorial (YouTube).

PyQt6 Course for Beginners - Parwiz Forogh (https://youtu.be/ot94H3-d5d8?t=193)
Detailed notes: https://tinyurl.com/y5db3c5c

A basic PyQt application built by instantiating and calling methods on QWidget
Alternative window types are QDialog and QMainWindow

Concepts introduced:
    - Importing QApplication and QWidget classes
    - Using QWidget
    - Displaying a window with the .show method
    - Passing arguments using sys.argv

Module/Class/Methods:
    QApplication    Provides interface with OS for command line args and exit codes
        __init__    Pass sys.argv or []
        exec        Runs the application and returns exit code

    QWidget         Base application form with no menus
        __init__
        show        Displays the form


Created November 12, 2022
"""

from  PyQt6.QtWidgets import QApplication, QWidget
import sys

app = QApplication(sys.argv)

# Can be QWidget, QDialog, or QMainWindow
window = QWidget()
window.show()

sys.exit(app.exec())

