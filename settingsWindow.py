"""
settingsWindow.py

This defines the layout and all anything calls
involving reading/writing to the settings file

Contains
    -   Contstructor setting general layout
    -   readSettings() == read settings function
    -   writeSettings() == a helper function for saveSettings()
    -   createSettingsBox() == function for specific button layout
        - Creates specific layout of buttons
        - Connects buttons and data entry lines to variables
    -   saveSettings() == a function that organizes settings variables 
        into an array and sends them to the writeSettings() function
    

Justin Seth
3/15/2022
"""

from PyQt5.QtWidgets import *
import json
#import mainWindow


class settingsWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Settings")
        self.setGeometry(100, 100, 500, 400)
        self.settings = self.readSettings()

        self.createSettingsBox()

        self.buttonBox = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel)

        # Connect save/cance with respective function
        self.buttonBox.accepted.connect(self.saveSettings)
        self.buttonBox.rejected.connect(self.close)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.settingsBox)
        mainLayout.addWidget(self.buttonBox)
        self.setLayout(mainLayout)

    def readSettings(self):
        with open("settings.json") as settings_file:
            settings_loaded = json.load(settings_file)

        settings = []
        for row in settings_loaded:
            settings.append(settings_loaded[row])

        return settings

    def writeSettings(self, settings):
        # Load settings as a dictionary
        with open("settings.json", "r") as settings_file:
            settings_loaded = json.load(settings_file)

        # Change settings dictionary to new settings
        for i, row in enumerate(settings_loaded):
            settings_loaded[row] = settings[i]

        # Write settings to settings.json
        with open("settings.json", "w") as outfile:
            json.dump(settings_loaded, outfile)

    def createSettingsBox(self):
        self.firstName = QLineEdit(self.settings[0].capitalize())
        self.firstName.setPlaceholderText("First")

        self.lastName = QLineEdit(self.settings[1].capitalize())
        self.lastName.setPlaceholderText("Last")

        self.nameLayout = QGridLayout()

        self.nameLayout.addWidget(self.firstName, 0, 0)
        self.nameLayout.addWidget(self.lastName, 0, 1)

        self.emailLabel = QLabel("Enter Emails:")
        self.fromEmail = QLineEdit(self.settings[2])
        self.toEmail = QLineEdit(self.settings[3])
        self.ecnLogPath = QLineEdit(self.settings[4])
        self.ecnLogPath.setDisabled(True)
        self.ecnFormatPath = QLineEdit(self.settings[5])
        self.ecnSavePath = QLineEdit(self.settings[6])

        self.settingsBox = QGroupBox("Settings")

        # creating a form layout
        layout = QFormLayout()

        layout.addRow(QLabel("Name"))
        layout.addRow(self.nameLayout)
        layout.addRow(QLabel("From Email"))
        layout.addRow(self.fromEmail)
        layout.addRow(QLabel("To Email"))
        layout.addRow(self.toEmail)
        layout.addRow(QLabel("ECN Log Path"))
        layout.addRow(self.ecnLogPath)
        layout.addRow(QLabel("ECN Format Path"))
        layout.addRow(self.ecnFormatPath)
        layout.addRow(QLabel("ECN Save Path"))
        layout.addRow(self.ecnSavePath)

        # setting layout
        self.settingsBox.setLayout(layout)

    def saveSettings(self):
        print("saving settings window")
        newSettings = []
        newSettings.append(self.firstName.text())
        newSettings.append(self.lastName.text())
        newSettings.append(self.fromEmail.text())
        newSettings.append(self.toEmail.text())
        newSettings.append(self.ecnLogPath.text())
        newSettings.append(self.ecnFormatPath.text())
        newSettings.append(self.ecnSavePath.text())

        # Write settings to settings file
        self.writeSettings(newSettings)
        self.close()
