"""Example code from the PyQt tutorial at https://tinyurl.com/4st2u3yp.
------------------------------------------------------------------------
Demonstrates creating a table in PyQt using a QTableWidget
This example is _not_ part of the PyQt Course for Beginners tutorial.

Concepts introduced:
    - Concept 1
    - Concept 2

Module/Class/Methods:
    QApplication                    
    QWidget                         
    QVBoxLayout

    QTableWidget
      setColumnCount
      setRowCount
      setHorizontalHeaderLabels
      setItem

    QTableWidgetItem

___________________________________________________________________________________________________

Created November 20, 2022
"""
import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout
from PyQt6.QtWidgets import QLineEdit
from PyQt6.QtWidgets import QTableWidget,QTableWidgetItem
class YuTextFrame(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('TutorialExample Add QTableWidget Example')  
        vbox = QVBoxLayout()
        
        # add table 
        table = QTableWidget(self)
        table.setColumnCount(5)
        table.setRowCount(10)
        #set table header
        table.setHorizontalHeaderLabels(['id','Name','Age','Sex','Address'])
        #add data
        table_data = []
        item_1 = ['001', 'John', 30, 'Male', 'Street No 2']
        item_2 = ['002', 'Lily', 32, 'Female', 'Street No 1']
        item_3 = ['003', 'Kate', 20, 'Male', 'Street No 3']
        item_4 = ['004', 'Tom', 22, 'Male', 'Street No 4']
        table_data.append(item_1)
        table_data.append(item_2)
        table_data.append(item_3)
        table_data.append(item_4)
        
        row = 0
        for r in table_data:
            col = 0
            for item in r:
                cell = QTableWidgetItem(str(item))
                table.setItem(row, col, cell)
                col += 1
            row += 1
        vbox.addWidget(table)

        self.setLayout(vbox)
        self.setGeometry(300,400,500,400)
        self.show()

if __name__ == '__main__':

    app = QApplication(sys.argv)
    frame = YuTextFrame()
    sys.exit(app.exec())