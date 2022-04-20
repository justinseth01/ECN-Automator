# ECN-Automator
This project is a project that automates the engineering documentation process at an
engineering company. This is written in python and uses the following
    -   sendgrid api for sending automated emails
    -   PYQT5 for creating an easy-to-use user interface
    -   openpyxl for reading, analyzing, and writing to excel documents
    -   python-docx for creating and editing word documents

This program is designed for engineering companies that still use the dated "ECN" (Engineering
Change Notice) system. Specifically, this program is designed to integrate seamlessly into
For companies that want to still use the dated "ECN" (Engineering Change Notice) documentation system,
this program helps significantly reduce engineering documentation time from 10 minutes to ~1 minute.

Specifically, this program has been designed to do the following:
    -   Obtain user's ECN information through a UI
    -   Read, write, and analyze excel sheets based on given inputted information
        from the user.
    -   Complete preformatted word documents with the engineering documentation
        information entered by the user.
    -   Send automated emails to respective personnel with newly created documentation

All of these tasks originally needed to be done manually, and was a 
rather tedious task due to its simplicity and how often it needed to be done. 
This program simplifies this process as much as possible by automating every step that is able to
be automated. This projecect also contains a settings window so the user does not need to 
re enter information like their name and email every time the program is run. It also
has built in UI error windows for when incorrect information is entered or the code blows
up in a try-catch block.
