# ECN-Automator
This project is a project that automates the engineering documentation process at an
engineering company. This is written in python and uses the following <br>
&emsp;-&emsp;sendgrid api for sending automated emails<br>
&emsp;-&emsp;PYQT5 for creating an easy-to-use user interface<br>
&emsp;-&emsp;openpyxl for reading, analyzing, and writing to excel documents<br>
&emsp;-&emsp;python-docx for creating and editing word documents<br>

This program is designed for engineering companies that still use the dated "ECN" (Engineering
Change Notice) system. Specifically, this program is designed to integrate seamlessly into
For companies that want to still use the dated "ECN" (Engineering Change Notice) documentation system,
this program helps significantly reduce engineering documentation time from 10 minutes to ~1 minute.

Specifically, this program has been designed to do the following:<br>
&emsp;-&emsp;Obtain user's ECN information through a UI<br>
&emsp;-&emsp;Read, write, and analyze excel sheets based on given inputted information<br>
&emsp;&emsp;from the user.<br>
&emsp;-&emsp;Complete preformatted word documents with the engineering documentation<br>
&emsp;&emsp;information entered by the user.<br>
&emsp;-&emsp;Send automated emails to respective personnel with newly created documentation<br>

All of these tasks originally needed to be done manually, and was a 
rather tedious task due to its simplicity and how often it needed to be done. 
This program simplifies this process as much as possible by automating every step that is able to
be automated. This projecect also contains a settings window so the user does not need to 
re enter information like their name and email every time the program is run. It also
has built in UI error windows for when incorrect information is entered or the code blows
up in a try-catch block.
