### This repo contains the code for the optical fiber billing management application

This application is coded using the python library tkinter library and is generated as a .exe file using Pyinstaller

The project folder should contain : 
gfa_luthern.py (python file with the code)
gfa_luth.ico (icon for the application) 


To generate the .exe file with pyinstaller, the following command is used in the command terminal :  

pyinstaller  --icon=gfa_luth.ico --add-data "gfa_luth.ico;." --windowed path/to/gfa_luthern.py

