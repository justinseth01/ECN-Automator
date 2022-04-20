"""
errorWindow.py

When an error is encountered within a try catch,this document 
is called to display a pop up window with the error.

Justin Seth
3/15/2022
"""

from PyQt5.QtWidgets import *


class error(QMessageBox):
    def __init__(self, message, detailedMessage=None):
        super().__init__()
        msgBox = QMessageBox()
        #msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msgBox.setText("ERROR")
        msgBox.setInformativeText(message)

        # Only add a detailed message if given one
        if detailedMessage != None:
            msgBox.setDetailedText(repr(detailedMessage))
        msgBox.setWindowTitle("Error")
        returnValue = msgBox.exec()

        if returnValue == QMessageBox.Ok:
            print("okay was pressed")
            self.close()
        else:
            print("Okay was not pressed")
            self.close()
