"""
customWidget.py

This document is the template for adding/removing
a "change". This is is called using the "add" 
and "remove" button in the right column of the UI.

Contains 6 different data entry boxes including
    -   MM number
    -   Description
    -   Part NUmber
    -   Old Quantity
    -   New Quantity
    -   Status Code

Justin Seth
3/15/2022
"""


from PyQt5.QtWidgets import *


class customWidget(QWidget):
    def __init__(self, parent=None):
        QWidget.__init__(self, parent=parent)
        self.layout = QGridLayout()

        self.mmNumber = QLineEdit()
        self.mmNumber.setPlaceholderText("MM Number")

        self.description = QLineEdit()
        self.description.setPlaceholderText("Description")

        self.partNumber = QLineEdit()
        self.partNumber.setPlaceholderText("Part Number")

        self.oldQuantity = QLineEdit()
        self.oldQuantity.setPlaceholderText("Old Quantity")

        self.newQuantity = QLineEdit()
        self.newQuantity.setPlaceholderText("New Quantity")

        self.statusCode = QLineEdit()
        self.statusCode.setPlaceholderText("Status Code")

        self.layout.addWidget(self.mmNumber, 0, 0)
        self.layout.addWidget(self.partNumber, 0, 1, 1, 2)
        self.layout.addWidget(self.description, 1, 0, 1, 3)
        self.layout.addWidget(self.oldQuantity, 2, 0)
        self.layout.addWidget(self.newQuantity, 2, 1)
        self.layout.addWidget(self.statusCode, 2, 2)

        self.setLayout(self.layout)
